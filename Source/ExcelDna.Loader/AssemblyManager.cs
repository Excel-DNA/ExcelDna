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
using ExcelDna.Loader.Logging;
using SevenZip.Compression.LZMA;

namespace ExcelDna.Loader
{
    // TODO: Lots more to make a flexible loader.
    internal static class AssemblyManager
    {
        internal static class AssemblyDictionary
        {
            static Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();

            private static string GetCleanAssemblyName(string assemblyName)
            {
                AssemblyName assName = new AssemblyName(assemblyName);
                // remove Version
                assName.Version = null;
                // remove Public Key
                assName.SetPublicKey(null);
                // remove Public KeyToken
                assName.SetPublicKeyToken(null);
                // remove Flags
                assName.Flags = AssemblyNameFlags.None;

                if (assName.CultureInfo == null)
                    assName.CultureInfo = CultureInfo.InvariantCulture;

                // other Variant
                //return assName.Name + ", Culture=" + assName.CultureInfo.ToString();

                return assName.FullName;
            }

            public static Assembly GetAssemblyByName(string assemblyName)
            {
                var assName = GetCleanAssemblyName(assemblyName);
             
                // Check our AssemblyResolve cache
                if (loadedAssemblies.ContainsKey(assName))
                    return loadedAssemblies[assName];

                //update AssemblyResolve Cache
                foreach (Assembly loadedAssemblyItem in AppDomain.CurrentDomain.GetAssemblies())
                {
                    AddLoadedAssembly(loadedAssemblyItem);
                }

                //// Check our AssemblyResolve cache after update again
                if (loadedAssemblies.ContainsKey(assName))
                    return loadedAssemblies[assName];

                return null;
            }

            public static void AddLoadedAssembly(Assembly assembly)
            {
                var assName = GetCleanAssemblyName(assembly.FullName);

                if (!loadedAssemblies.ContainsKey(assName))
                loadedAssemblies.Add(assName, assembly);
            }
        }

        static string pathXll;
        static IntPtr hModule;
     
        internal static void Initialize(IntPtr hModule, string pathXll)
        {
            AssemblyManager.pathXll = pathXll;
            AssemblyManager.hModule = hModule;
            
            // TODO: Load up the DnaFile and Assembly names ?

            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        private static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            byte[] assemblyBytes= null;
                  
            Logger.Initialization.Info("Attempting to load {0} from resources.", args.Name);

            // try to get Assembly from already loaded assemblies
            var loadedAssembly = AssemblyDictionary.GetAssemblyByName(args.Name);
            if (loadedAssembly != null)
                return loadedAssembly;

            AssemblyName assName = new AssemblyName(args.Name);
            ushort lcid = 0;
            // if we don't have a culture info or of it is the Neutral culture take 0, otherwise take LCID
            if (assName.CultureInfo != null &&  !string.IsNullOrEmpty(assName.CultureInfo.Name))
                lcid = (ushort)assName.CultureInfo.TextInfo.LCID;
            assemblyBytes = ResourceHelper.LoadResourceBytes(hModule, "ASSEMBLY", assName.Name.ToUpper(), lcid);
          
            if (assemblyBytes == null)
            {                
                // if the missing Assembly is only a Resource use a lower LogLevel
                if (assName.Name.EndsWith(".RESOURCES", StringComparison.InvariantCultureIgnoreCase))
                    Logger.Initialization.Verbose("Assembly {0} could not be loaded from resources.", args.Name);
                else
                    Logger.Initialization.Warn("Assembly {0} could not be loaded from resources.", args.Name);
                return null;
            }

            Logger.Initialization.Info("Trying Assembly.Load for {0} (from {1} bytes).", args.Name, assemblyBytes.Length);
            //File.WriteAllBytes(@"c:\Temp\" + name + ".dll", assemblyBytes);
            
            try
            {              
                loadedAssembly = Assembly.Load(assemblyBytes);
                Logger.Initialization.Info("Assembly Loaded from bytes. FullName: {0}", loadedAssembly.FullName);
                AssemblyDictionary.AddLoadedAssembly(loadedAssembly);
                return loadedAssembly;
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "Error during Assembly Load from bytes");
            }
            return null;
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
            else
            {
                throw new ArgumentOutOfRangeException("type", "Unknown resource type. Only types 0 (Assembly), 1 (Dna file), 2 (Image) or 3 (Source) are valid.");
            }
			return ResourceHelper.LoadResourceBytes(hModule, typeName, resourceName,0);
		}     
    }

    internal unsafe static class ResourceHelper
    {
		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr FindResource(
			IntPtr hModule,
			string lpName,
			string lpType);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr FindResourceEx(
          IntPtr hModule,
          string lpType,
          string lpName,
          uint wLanguage);

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

        internal static IntPtr FindResourceIndependet(IntPtr hModule, string resourceName, string typeName, uint lcid)
        {
            IntPtr hResInfo = FindResourceEx(hModule, typeName, resourceName, lcid);
            if (hResInfo == IntPtr.Zero)
                hResInfo = FindResource(hModule, resourceName, typeName);
            return hResInfo;
        }

        // Load the resource, trying also as compressed if no uncompressed version is found.
        // If the resource type ends with "_LZMA", we decompress from the LZMA format.
        internal static byte[] LoadResourceBytes(IntPtr hModule, string typeName, string resourceName, ushort lcid)
		{
            // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
            // Logger.Initialization.Info("LoadResourceBytes for resource {0} of type {1}", resourceName, typeName);
			IntPtr hResInfo	= FindResourceIndependet(hModule, resourceName, typeName, lcid);
			if (hResInfo == IntPtr.Zero)
			{
				// We expect this null result value when the resource does not exists.

				if (!typeName.EndsWith("_LZMA"))
				{
					// Try the compressed name.
					typeName += "_LZMA";
					hResInfo = FindResourceIndependet(hModule, resourceName, typeName, lcid);
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
			
			if (typeName.EndsWith("_LZMA"))
				return Decompress(resourceBytes);
			else 
				return resourceBytes;
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

	}
}