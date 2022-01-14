//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using SevenZip.Compression.LZMA;
using System.Threading;
using System.Globalization;

internal static class ResourceHelper
{
    internal enum TypeName
    {
        CONFIG = -1,
        ASSEMBLY = 0,
        DNA = 1,
        IMAGE = 2,
        SOURCE = 3,
        PDB = 4,
    }

    // TODO: Learn about locales
    private const ushort localeNeutral		= 0;
	private const ushort localeEnglishUS	= 1033;
	private const ushort localeEnglishSA	= 7177;

	[DllImport("kernel32.dll")]
	private static extern IntPtr BeginUpdateResource(
		string pFileName,
		bool bDeleteExistingResources);

	[DllImport("kernel32.dll")]
	private static extern bool EndUpdateResource(
		IntPtr hUpdate,
		bool fDiscard);
	
	//, EntryPoint="UpdateResourceA", CharSet=CharSet.Ansi,
	[DllImport("kernel32.dll", SetLastError=true)]
	private static extern bool UpdateResource(
		IntPtr hUpdate,
		string lpType,
		string lpName,
		ushort wLanguage,
		IntPtr lpData,
		uint cbData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool UpdateResource(
        IntPtr hUpdate,
        string lpType,
        IntPtr intResource,
        ushort wLanguage,
        IntPtr lpData,
        uint cbData);

    // This overload provides the resource type and name conversions that would be done by MAKEINTRESOURCE
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool UpdateResource(
        IntPtr hUpdate,
        uint lpType,
        uint lpName,
        ushort wLanguage,
        IntPtr lpData,
        uint cbData);

    [DllImport("version.dll", SetLastError = true)]
    private static extern uint GetFileVersionInfoSize(
        string lptstrFilename,
        out uint lpdwHandle);

    [DllImport("version.dll", SetLastError = true)]
    private static extern bool GetFileVersionInfo(
        string lptstrFilename,
        uint dwHandle,
        uint dwLen,
        byte[] lpData);

	[DllImport("kernel32.dll")]
	private static extern uint GetLastError();

    internal class ResourceUpdater
    {
        int typelibIndex = 0;
        IntPtr _hUpdate;
        List<object> updateData = new List<object>();

        object lockResource = new object();

        Queue<ManualResetEvent> finishedTask = new Queue<ManualResetEvent>();

        public ResourceUpdater(string fileName)
        {
            _hUpdate = BeginUpdateResource(fileName, false);
            if (_hUpdate == IntPtr.Zero)
            {
                throw new Win32Exception();
            }
        }

        private void CompressDoUpdateHelper(byte[] content, string name, TypeName typeName, bool compress)
        {
            if (compress)
                content = SevenZipHelper.Compress(content);
            DoUpdateResource(typeName.ToString() + (compress ? "_LZMA" : ""), name, content);
        }

        public string AddFile(byte[] content, string name, TypeName typeName, bool compress, bool multithreading)
        {
            XorRecode(content);

            Debug.Assert(name == name.ToUpperInvariant());

            if (multithreading)
            {
                var mre = new ManualResetEvent(false);
                finishedTask.Enqueue(mre);
                ThreadPool.QueueUserWorkItem(delegate
                    {
                        CompressDoUpdateHelper(content, name, typeName, compress);
                        mre.Set();
                    }
                );
            }
            else
            {
                CompressDoUpdateHelper(content, name, typeName, compress);
            }

            return name;
        }

        public string AddAssembly(string path, bool compress, bool multithreading, bool includePdb)
        {
            try
            {
                byte[] assemblyBytes = File.ReadAllBytes(path);
                // Not just into the Reflection context, because this Load is used to get the name and also to 
                // check that the assembly can Load from bytes (mixed assemblies can't).
                Assembly assembly = Assembly.Load(assemblyBytes);
                AssemblyName assemblyName = assembly.GetName();
                CultureInfo cultureInfo = assemblyName.CultureInfo;
                string name = assemblyName.Name.ToUpperInvariant(); // .ToUpperInvariant().Replace(".", "_");

                // For .resources assemblies we add the Culture name to the packed name
                if (name.EndsWith(".RESOURCES") && cultureInfo != null && !string.IsNullOrEmpty(cultureInfo.Name))
                {
                    name += "." + cultureInfo.Name.ToUpperInvariant();
                }

                AddFile(assemblyBytes, name, TypeName.ASSEMBLY, compress, multithreading);

                string pdbFile = Path.ChangeExtension(path, "pdb");
                if (includePdb && File.Exists(pdbFile))
                {
                    byte[] pdbBytes = File.ReadAllBytes(pdbFile);
                    AddFile(pdbBytes, name, TypeName.PDB, compress, multithreading);
                }
                return name;
            }
            catch (Exception e)
            {
                Console.WriteLine("Assembly at " + path + " could not be packed. Possibly a mixed assembly? (These are not supported yet.)\r\nException: " + e);
                return null;
            }
        }

        public int AddTypeLib(byte[] data)
        {
            lock (lockResource)
            {
                string typeName = "TYPELIB";
                typelibIndex++;

                Console.WriteLine(string.Format("  ->  Updating typelib: Type: {0}, Index: {1}, Length: {2}", typeName, typelibIndex, data.Length));
                GCHandle pinHandle = GCHandle.Alloc(data, GCHandleType.Pinned);
                updateData.Add(pinHandle);

                bool result = ResourceHelper.UpdateResource(_hUpdate, typeName, (IntPtr)typelibIndex, localeNeutral, pinHandle.AddrOfPinnedObject(), (uint)data.Length);
                if (!result)
                {
                    throw new Win32Exception();
                }

            }
            return typelibIndex;
        }

        public void DoUpdateResource(string typeName, string name, byte[] data)
        {
            lock (lockResource)
            {
                Console.WriteLine(string.Format("  ->  Updating resource: Type: {0}, Name: {1}, Length: {2}", typeName, name, data.Length));
                GCHandle pinHandle = GCHandle.Alloc(data, GCHandleType.Pinned);
                updateData.Add(pinHandle);

                bool result = ResourceHelper.UpdateResource(_hUpdate, typeName, name, localeNeutral, pinHandle.AddrOfPinnedObject(), (uint)data.Length);
                if (!result)
                {
                    throw new Win32Exception();
                }
            }
        }

        public void RemoveResource(string typeName, string name)
        {
            lock (lockResource)
            {
                bool result = ResourceHelper.UpdateResource(_hUpdate, typeName, name, localeNeutral, IntPtr.Zero, 0);
                if (!result)
                {
                    throw new Win32Exception();
                }
            }
        }

        public void CopyFileVersion(string fromFile)
        {
            uint ignored;
            uint versionSize = ResourceHelper.GetFileVersionInfoSize(fromFile, out ignored);
            if (versionSize == 0)
            {
                throw new Win32Exception();
            }

            byte[] versionBuf = new byte[versionSize];
            bool result = ResourceHelper.GetFileVersionInfo(fromFile, ignored, versionSize, versionBuf);
            if (!result)
            {
                throw new Win32Exception();
            }

            GCHandle versionBufHandle = GCHandle.Alloc(versionBuf, GCHandleType.Pinned);
            try
            {
                lock (lockResource)
                {
                    uint versionResourceType = 16;
                    uint versionResourceId = 1;
                    result = ResourceHelper.UpdateResource(
                        _hUpdate,
                        versionResourceType,
                        versionResourceId,
                        localeNeutral,
                        versionBufHandle.AddrOfPinnedObject(),
                        versionSize);
                    if (!result)
                    {
                        throw new Win32Exception();
                    }
                }
            }
            finally
            {
                versionBufHandle.Free();
            }
        }

        public void EndUpdate()
        {
            EndUpdate(false);
        }

        public void EndUpdate(bool discard)
        {
            if (finishedTask.Count > 0)
            {
                while (finishedTask.Count > 0)
                {
                    int cnt = finishedTask.Count;
                    // WaitAll accepts a maximum of 64 WaitHandles
                    if (cnt > 64) cnt = 64;

                    ManualResetEvent[] mre = new ManualResetEvent[cnt];

                    for (int i = 0; i < cnt; i++)
                        mre[i] = finishedTask.Dequeue();

                    WaitHandle.WaitAll(mre);
                }
            }

            bool result = EndUpdateResource(_hUpdate, discard);
            if (!result)
            {
                throw new Win32Exception();
            }
        }

        static readonly byte[] _xorKeys = System.Text.Encoding.ASCII.GetBytes("ExcelDna");
        static void XorRecode(byte[] data)
        {
            var keys = _xorKeys;
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(keys[i % keys.Length] ^ data[i]);
            }
        }
    }
}
