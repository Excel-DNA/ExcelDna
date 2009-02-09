/*
  Copyright (C) 2005-2008 Govert van Drimmelen

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
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Resources;
using System.IO;

namespace ExcelDna.Loader
{
    // TODO: Lots more to make a flexible loader.
    internal class AssemblyManager
    {
        static string pathXll;
        static IntPtr hModule;
        static Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string,Assembly>();

        internal static void Initialize(IntPtr hModule, string pathXll)
        {
            AssemblyManager.pathXll = pathXll;
            AssemblyManager.hModule = hModule;
            loadedAssemblies.Add(Assembly.GetExecutingAssembly().FullName, Assembly.GetExecutingAssembly());

            //// Testing ...
            //Assembly a = ResourceHelper.LoadAssemblyFromResources(hModule, "EXCELDNA_INTEGRATION");
            //loadedAssemblies.Add(a.FullName, a);

            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;
        }
        
        private static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // Check cache - includes special case of ExcelDna.Loader.
            if (loadedAssemblies.ContainsKey(args.Name))
            {
                return loadedAssemblies[args.Name];
            }

            // Now check in resources ...
            if (args.Name.StartsWith("ExcelDna.Integration") || 
                args.Name.StartsWith("ExcelDna,") /* Special case for pre-0.14 versions of ExcelDna */)
            {
                foreach (Assembly ass in loadedAssemblies.Values)
                {
                    if (ass.FullName.StartsWith("ExcelDna.Integration"))
                        return ass;
                }

                Debug.Print("Loading ExcelDna.Integration from resources.");
                Assembly loadedAssembly = ResourceHelper.LoadAssemblyFromResources(hModule, "EXCELDNA_INTEGRATION");
                loadedAssemblies.Add(loadedAssembly.FullName, loadedAssembly);
                return loadedAssembly;
            }

            // TODO: Other Assemblies.
            Debug.WriteLine("AssemblyManager.AssemblyResolve failed: " + args.Name);
            return null;
        }

        internal static byte[] GetAssemblyBytes(string assemblyName)
        {
            // TODO: Other assemblies
            if (assemblyName.StartsWith("ExcelDna.Integration"))
            {
                return ResourceHelper.ReadAssemblyFromResources(hModule, "EXCELDNA_INTEGRATION");
            }
            return null;
        }
    }

    internal unsafe static class ResourceHelper
    {
        [DllImport("KERNEL32.DLL")]
        internal static extern IntPtr FindResource(
            IntPtr hModule,
            string lpName,
            string lpType);
        [DllImport("KERNEL32.DLL")]
        internal static extern IntPtr LoadResource(
            IntPtr hModule,
            IntPtr hResInfo);
        [DllImport("KERNEL32.DLL")]
        internal static extern IntPtr LockResource(
            IntPtr hResData);
        [DllImport("KERNEL32.DLL")]
        internal static extern uint SizeofResource(
            IntPtr hModule,
            IntPtr hResInfo);

        internal unsafe static byte[] ReadAssemblyFromResources(IntPtr hModule, string resourceName)
        {
            IntPtr hResInfo = FindResource(hModule, resourceName, "ASSEMBLY");
            IntPtr hResData = LoadResource(hModule, hResInfo);
            uint size = SizeofResource(hModule, hResInfo);
            IntPtr pAssemblyBytes = LockResource(hResData);
            byte[] assemblyBytes = new byte[size];
            Marshal.Copy(pAssemblyBytes, assemblyBytes, 0, (int)size);
            return assemblyBytes;
        }

        internal unsafe static Assembly LoadAssemblyFromResources(IntPtr hModule, string resourceName)
        {
            byte[] assemblyBytes = ReadAssemblyFromResources(hModule, resourceName);
            Assembly loadedAssembly = Assembly.Load(assemblyBytes);
            return loadedAssembly;
        }
    }

}
