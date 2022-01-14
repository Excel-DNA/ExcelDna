//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using SevenZip.Compression.LZMA;

namespace ExcelDna.ManagedHost
{
    // TODO: Lots more to make a flexible loader.
    internal static class AssemblyManager
    {
        static IntPtr hModule;
        static string pathXll;
        static Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string,Assembly>();
#if NETCOREAPP
        static ExcelDnaAssemblyLoadContext alc;
#endif

        internal static void Initialize(IntPtr hModule, string pathXll
#if NETCOREAPP
            , ExcelDnaAssemblyLoadContext alc
#endif
            )
        {
            AssemblyManager.hModule = hModule;
            AssemblyManager.pathXll = pathXll;
#if NETCOREAPP
            AssemblyManager.alc = alc;
#endif
            loadedAssemblies.Add(Assembly.GetExecutingAssembly().FullName, Assembly.GetExecutingAssembly());
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        internal static Assembly AssemblyResolve(AssemblyName assemblyName)
        {
			string name;
			byte[] assemblyBytes;
            Assembly loadedAssembly = null;

            CultureInfo cultureInfo = assemblyName.CultureInfo;
			name = assemblyName.Name.ToUpperInvariant();
			
			if (name == "EXCELDNA.MANAGEDHOST")
			{
				// Loader must have been loaded from bytes.
				// But I have seen the Loader, and it is us.
				return Assembly.GetExecutingAssembly();
			}

            bool isResourceAssembly = name.EndsWith(".RESOURCES");

            // This check and mapping must match that done when packing (in ResourceHelper.cs : ResourceUpdate.AddAssembly)
            if (isResourceAssembly && cultureInfo != null && !string.IsNullOrEmpty(cultureInfo.Name))
            {
                name += "." + cultureInfo.Name.ToUpperInvariant();
            }

            // Check our AssemblyResolve cache
            if (loadedAssemblies.TryGetValue(name, out loadedAssembly))
                return loadedAssembly;

            // Check if it is loaded in the AppDomain already, 
            // e.g. from resources as an ExternalLibrary
            loadedAssembly = GetAssemblyIfLoaded(assemblyName);
            if (loadedAssembly != null)
            {
                Logger.Initialization.Info("Assembly {0} was found to already be loaded into the AppDomain.", name);
                loadedAssemblies[name] = loadedAssembly;
                return loadedAssembly;
            }

            // Now check in resources ...
            // We expect failures when loading .resources assemblies, so only log at the Verbose level.
            // From: http://blogs.msdn.com/b/suzcook/archive/2003/05/29/57120.aspx
            // "Note: Unless you are explicitly debugging the failure of a resource to load, 
            //        you will likely want to ignore failures to find assemblies with the ".resources" extension 
            //        with the culture set to something other than "neutral". Those are expected failures when the 
            //        ResourceManager is probing for satellite assemblies."

            if (isResourceAssembly)
                Logger.Initialization.Verbose("Attempting to load {0} from resources.", name);
            else
                Logger.Initialization.Info("Attempting to load {0} from resources.", name);

			assemblyBytes = GetResourceBytes(name, 0);
			if (assemblyBytes == null)
			{
                if (isResourceAssembly)
                    Logger.Initialization.Verbose("Assembly {0} could not be loaded from resources (ResourceManager probing for satellite assemblies).", name);
                else
                    Logger.Initialization.Warn("Assembly {0} could not be loaded from resources.", name);
				return null;
			}

            byte[] pdbBytes = GetResourceBytes(name, 4);
            if (pdbBytes == null)
                Logger.Initialization.Info("Trying Assembly.Load for {0} (from {1} bytes, without pdb).", name, assemblyBytes.Length);
            else
                Logger.Initialization.Info("Trying Assembly.Load for {0} (from {1} bytes, with {2} bytes of pdb).", name, assemblyBytes.Length, pdbBytes.Length);
			try
			{
                loadedAssembly = LoadFromAssemblyBytes(assemblyBytes, pdbBytes);
                loadedAssemblies[name] = loadedAssembly;
				return loadedAssembly;
			}
			catch (Exception e)
			{
                Logger.Initialization.Error(e, "Error during Assembly Load from bytes");
			}
			return null;
        }

        internal static Assembly LoadFromAssemblyPath(string assemblyPath)
        {
#if NETCOREAPP
            return alc.LoadFromAssemblyPath(assemblyPath);
#else
            return Assembly.LoadFrom(assemblyPath);
#endif
        }

        internal static Assembly LoadFromAssemblyBytes(byte[] assemblyBytes, byte[] pdbBytes)
        {
#if NETCOREAPP
            return alc.LoadFromAssemblyBytes(assemblyBytes, pdbBytes);
#else
            return (pdbBytes == null) ? Assembly.Load(assemblyBytes) : Assembly.Load(assemblyBytes, pdbBytes);
#endif
        }

        // TODO: This method probably should not be here.
        internal static byte[] GetResourceBytes(string resourceName, int type) // types: 0 - Assembly, 1 - Dna file, 2 - Image
		{
            // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
            // Logger.Initialization.Info("GetResourceBytes for resource {0} of type {1}", resourceName, type);
			string typeName;
			if (type == 0)
			{
				typeName = "ASSEMBLY";
			}
			else if (type == 1)
			{
				typeName = "DNA";
			}
            else if (type == 2)
            {
                typeName = "IMAGE";
            }
            else if (type == 3)
            {
                typeName = "SOURCE";
            }
            else if (type == 4)
            {
                typeName = "PDB";
            }
            else
            {
                throw new ArgumentOutOfRangeException("type", "Unknown resource type. Only types 0 (Assembly), 1 (Dna file), 2 (Image) or 3 (Source) are valid.");
            }
			return ResourceHelper.LoadResourceBytes(hModule, typeName, resourceName);
		}

        // A copy of this method lives in ExcelDna.Integration - ExternalLibrary.cs
        private static Assembly GetAssemblyIfLoaded(AssemblyName assemblyName)
        {
#if NETCOREAPP
            IEnumerable<Assembly> assemblies = alc.Assemblies;
#else
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
#endif
            foreach (Assembly loadedAssembly in assemblies)
            {
                AssemblyName loadedName = loadedAssembly.GetName();
                if (loadedName.Name.Equals(assemblyName.Name, StringComparison.OrdinalIgnoreCase))
                {
                    // For resources, also check the culture
                    if (loadedName.Name.EndsWith(".RESOURCES", StringComparison.OrdinalIgnoreCase))
                    {
                        if ((loadedName.CultureInfo == null) && (assemblyName.CultureInfo != null) ||
                            (loadedName.CultureInfo != null) && (assemblyName.CultureInfo == null) ||
                            !string.Equals(loadedName.CultureInfo.Name, assemblyName.CultureInfo.Name))
                        {
                            continue; // next loadedAssembly
                        }
                    }
                    return loadedAssembly;
                }
            }
            return null;
        }
    }

    internal static class ResourceHelper
    {
		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr FindResource(
			IntPtr hModule,
			string lpName,
			string lpType);
        [DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr LoadResource(
			IntPtr hModule,
			IntPtr hResInfo);
        [DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr LockResource(
			IntPtr hResData);
        [DllImport("kernel32.dll", SetLastError = true)]
		private static extern uint SizeofResource(
			IntPtr hModule,
			IntPtr hResInfo);

        [DllImport("kernel32.dll")]
		private static extern uint GetLastError();

		// Load the resource, trying also as compressed if no uncompressed version is found.
		// If the resource type ends with "_LZMA", we decompress from the LZMA format.
		internal static byte[] LoadResourceBytes(IntPtr hModule, string typeName, string resourceName)
		{
            // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
            // Logger.Initialization.Info("LoadResourceBytes for resource {0} of type {1}", resourceName, typeName);
			IntPtr hResInfo	= FindResource(hModule, resourceName, typeName);
			if (hResInfo == IntPtr.Zero)
			{
				// We expect this null result value when the resource does not exists.

				if (!typeName.EndsWith("_LZMA"))
				{
					// Try the compressed name.
					typeName += "_LZMA";
					hResInfo = FindResource(hModule, resourceName, typeName);
				}
				if (hResInfo == IntPtr.Zero)
				{
                    // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
                    // Logger.Initialization.Info("Resource not found - resource {0} of type {1}", resourceName, typeName);
                    Debug.Print("ResourceHelper.LoadResourceBytes - Resource not found - resource {0} of type {1}", resourceName, typeName);
					// Return null to indicate that the resource was not found.
					return null;
				}
			}
            IntPtr hResData	= LoadResource(hModule, hResInfo);
			if (hResData == IntPtr.Zero)
			{
				// Unexpected error - this should not happen
                // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
                //Logger.Initialization.Error("Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
                Debug.Print("ResourceHelper.LoadResourceBytes - Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
				throw new Win32Exception();
			}
            uint   size	= SizeofResource(hModule, hResInfo);
            IntPtr pResourceBytes = LockResource(hResData);
            byte[] resourceBytes = new byte[size];
			Marshal.Copy(pResourceBytes, resourceBytes, 0, (int)size);

            byte[] resultBytes;
			if (typeName.EndsWith("_LZMA"))
				resultBytes = Decompress(resourceBytes);
			else 
				resultBytes = resourceBytes;

            XorRecode(resultBytes);
            return resultBytes;
		}

		private static byte[] Decompress(byte[] inputBytes)
		{
			MemoryStream newInStream = new MemoryStream(inputBytes);
			Decoder decoder = new Decoder();
			newInStream.Seek(0, 0);
			MemoryStream newOutStream = new MemoryStream();
			byte[] properties2 = new byte[5];
			if (newInStream.Read(properties2, 0, 5) != 5)
				throw (new Exception("input .lzma is too short"));
			long outSize = 0;
			for (int i = 0; i < 8; i++)
			{
				int v = newInStream.ReadByte();
				if (v < 0)
					throw (new Exception("Can't Read 1"));
				outSize |= ((long)(byte)v) << (8 * i);
			}
			decoder.SetDecoderProperties(properties2);
			long compressedSize = newInStream.Length - newInStream.Position;
			decoder.Code(newInStream, newOutStream, compressedSize, outSize, null);
			byte[] b = newOutStream.ToArray();
			return b;
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
