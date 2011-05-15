using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using SevenZip.Compression.LZMA;

internal unsafe static class ResourceHelper
{
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

	[DllImport("kernel32.dll")]
	private static extern uint GetLastError();
	internal unsafe class ResourceUpdater
	{
        int typelibIndex = 0;
		IntPtr _hUpdate;
		List<object> updateData = new List<object>();

		public ResourceUpdater(string fileName)
		{
			_hUpdate = BeginUpdateResource(fileName, false);
			if (_hUpdate == IntPtr.Zero)
			{
				throw new Win32Exception();
			}
		}

		public string AddAssembly(string path)
		{
			try
			{
				byte[] assBytes = File.ReadAllBytes(path);
				// Not just into the Reflection context, bacause this Load is used to get the name and also to 
				// check that the assembly can Load from bytes (mixed assemblies can't).
				Assembly ass = Assembly.Load(assBytes);
				string name = ass.GetName().Name.ToUpperInvariant(); // .ToUpperInvariant().Replace(".", "_");
				byte[] data = SevenZipHelper.Compress(assBytes);
				DoUpdateResource("ASSEMBLY_LZMA", name, data);
				return name;
			}
			catch (Exception e)
			{
				Console.WriteLine("Assembly at " + path + " could not be packed. Possibly a mixed assembly? (These are not supported yet.)\r\nException: " + e);
				return null;
			}
		}

        public void AddDnaFileUncompressed(byte[] dnaContent, string name)
        {
            Debug.Assert(name == name.ToUpperInvariant());
            DoUpdateResource("DNA", name, dnaContent);
        }

		public void AddDnaFile(byte[] dnaContent, string name)
		{
			Debug.Assert(name == name.ToUpperInvariant());
			byte[] data = SevenZipHelper.Compress(dnaContent);
			DoUpdateResource("DNA_LZMA", name, data);
		}

        public void AddImage(byte[] imageBytes, string name)
        {
            Debug.Assert(name == name.ToUpperInvariant());
            byte[] data = SevenZipHelper.Compress(imageBytes);
            DoUpdateResource("IMAGE_LZMA", name, data);
        }

		public void AddConfigFile(byte[] configContent, string name)
		{
			DoUpdateResource("CONFIG", name, configContent);
		}

        public int AddTypeLib(byte[] data)
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

            return typelibIndex;
        }

		public void DoUpdateResource(string typeName, string name, byte[] data)
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

		public void RemoveResource(string typeName, string name)
		{
			bool result = ResourceHelper.UpdateResource(_hUpdate, typeName, name, localeEnglishUS, IntPtr.Zero, 0);
			if (!result)
			{
				throw new Win32Exception();
			}
		}

		public void EndUpdate()
		{
			EndUpdate(false);
		}

		public void EndUpdate(bool discard)
		{
			bool result = EndUpdateResource(_hUpdate, discard);
			if (!result)
			{
				throw new Win32Exception();
			}
		}
	}
}