/*
  Copyright (C) 2005-2011 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
//using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Resources;
using System.IO;
using SevenZip.Compression.LZMA;
using System.ComponentModel;
using System.Security.Permissions;
using System.Security.Policy;
using System.Security;

namespace ExcelDna.Loader
{
    // TODO: Lots more to make a flexible loader.
    internal static class AssemblyManager
    {
        static string pathXll;
        static IntPtr hModule;
        static Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string,Assembly>();

        internal static void Initialize(IntPtr hModule, string pathXll)
        {
            AssemblyManager.pathXll = pathXll;
            AssemblyManager.hModule = hModule;
            loadedAssemblies.Add(Assembly.GetExecutingAssembly().FullName, Assembly.GetExecutingAssembly());

            // TODO: Load up the DnaFile and Assembly names ?

            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;
        }
       
        private static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
			string name;
			byte[] assemblyBytes;

			AssemblyName assName = new AssemblyName(args.Name);
			name = assName.Name.ToUpperInvariant();

			if (name == "EXCELDNA") /* Special case for pre-0.14 versions of ExcelDna */
			{
				name = "EXCELDNA.INTEGRATION";
			}
			
			if (name == "EXCELDNA.LOADER")
			{
				// Loader must have been loaded from bytes.
				// But I have seen the Loader, and it is us.
				return Assembly.GetExecutingAssembly();
			}

            // Check cache
            if (loadedAssemblies.ContainsKey(name))
                return loadedAssemblies[name];

            // Now check in resources ...
			Debug.Print("Attempting to load {0} from resources.", name);
			assemblyBytes = GetResourceBytes(name, 0);
			if (assemblyBytes == null)
			{
				Debug.Print("Assembly {0} could not be loaded from resources.", name);
				return null;
			}

			Debug.Print("Trying Assembly.Load for {0} (from {1} bytes).", name, assemblyBytes.Length);
			//File.WriteAllBytes(@"c:\Temp\" + name + ".dll", assemblyBytes);
			try
			{
				Assembly loadedAssembly = Assembly.Load(assemblyBytes);
				Debug.Print("Assembly Loaded from bytes. FullName: {0}", loadedAssembly.FullName);
				loadedAssemblies.Add(name, loadedAssembly);
				return loadedAssembly;
			}
			catch (Exception e)
			{
				Debug.Print("Exception during Assembly Load from bytes. Exception: {0}", e);
				// TODO: Trace / Log.
			}
			return null;
        }

        // TODO: This method probably should not be here.
		internal static byte[] GetResourceBytes(string resourceName, int type) // types: 0 - Assembly, 1 - Dna file, 2 - Image
		{
			Debug.Print("GetResourceBytes for resource {0} of type {1}", resourceName, type);
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
            else
            {
                throw new ArgumentOutOfRangeException("type", "Unknown resource type. Only types 0 (Assembly), 1 (Dna file) and 2 (Image) are valid.");
            }
			return ResourceHelper.LoadResourceBytes(hModule, typeName, resourceName);
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
			Debug.Print("LoadResourceBytes for resource {0} of type {1}", resourceName, typeName);
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
					Debug.Print("Resource not found - resource {0} of type {1}", resourceName, typeName);
					// Return null to indicate that the resource was not found.
					return null;
				}
			}
            IntPtr hResData	= LoadResource(hModule, hResInfo);
			if (hResData == IntPtr.Zero)
			{
				// Unexpected error - this should not happen
				Debug.Print("Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
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