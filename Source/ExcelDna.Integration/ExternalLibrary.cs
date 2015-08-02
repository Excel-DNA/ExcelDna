//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
	// TODO: Allow Com References (via TlbImp?)/Exported Libraries
	// DOCUMENT When loading ExternalLibraries, we check first the path given in the Path attribute:
	// if there is no such file, we try to find a file with the right name in the same 
	// directory as the .xll.
	// We load files with .dna extension as Dna Libraries

	[Serializable]
	[XmlType(AnonymousType = true)]
	public class ExternalLibrary
	{
		private string _Path;
		[XmlAttribute]
		public string Path
		{
			get { return _Path; }
			set { _Path = value; }
		}

        private string _TypeLibPath;
        [XmlAttribute]
        public string TypeLibPath
        {
            get { return _TypeLibPath; }
            set { _TypeLibPath = value; }
        }

        private bool _ComServer;
        [XmlAttribute]
        public bool ComServer
        {
            get { return _ComServer; }
            set { _ComServer = value; }
        }

		private bool _Pack = false;
		[XmlAttribute]
		public bool Pack
		{
			get { return _Pack; }
			set { _Pack = value; }
		}

        private bool _LoadFromBytes = false;
        [XmlAttribute]
        public bool LoadFromBytes
        {
            get { return _LoadFromBytes; }
            set { _LoadFromBytes = value; }
        }

		private bool _ExplicitExports = false;
		[XmlAttribute]
		public bool ExplicitExports
		{
			get { return _ExplicitExports; }
			set { _ExplicitExports = value; }
		}

        private bool _ExplicitRegistration = false;
        [XmlAttribute]
        public bool ExplicitRegistration
        {
            get { return _ExplicitRegistration; }
            set { _ExplicitRegistration = value; }
        }

		internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot, DnaLibrary dnaLibrary)
		{
			List<ExportedAssembly> list = new List<ExportedAssembly>();

			try
			{
                string realPath = Path;
				if (Path.StartsWith("packed:"))
				{
                    // The ExternalLibrary is packed.
                    // We'll have to load it from resources.
					string resourceName = Path.Substring(7);
					if (Path.EndsWith(".DNA", StringComparison.OrdinalIgnoreCase))
					{
						byte[] dnaContent = ExcelIntegration.GetDnaFileBytes(resourceName);
						DnaLibrary lib = DnaLibrary.LoadFrom(dnaContent, pathResolveRoot);
						if (lib == null)
						{
                            Logger.Initialization.Error("External library could not be registered - Path: {0}\r\n - Packed DnaLibrary could not be loaded", Path);
							return list;
						}

						return lib.GetAssemblies(pathResolveRoot);
					}
					else
					{
                        // DOCUMENT: TypeLibPath which is a resource in a library is denoted as fileName.dll\4
                        // For packed assemblies, we set TypeLibPath="packed:2"
                        string typeLibPath = null;
                        if (!string.IsNullOrEmpty(TypeLibPath) && TypeLibPath.StartsWith("packed:"))
                        {
                            typeLibPath = DnaLibrary.XllPath + @"\" + TypeLibPath.Substring(7);
                        }

                        // It would be nice to check here whether the assembly is loaded already.
                        // But because of the name mangling in the packing we can't easily check.

                        // So we make the following assumptions:
                        // 1. Packed assemblies won't also be loadable from files (else they might be loaded twice)
                        // 2. ExternalLibrary loads will happen before reference loads via AssemblyResolve.
                        // Under these assumptions we should not have assemblies loaded more than once, 
                        // even if not checking here.
                        byte[] rawAssembly = ExcelIntegration.GetAssemblyBytes(resourceName);
					    Assembly assembly = Assembly.Load(rawAssembly);
						list.Add(new ExportedAssembly(assembly, ExplicitExports, ExplicitRegistration, ComServer, false, typeLibPath, dnaLibrary));
						return list;
					}
				}
				if (Uri.IsWellFormedUriString(Path, UriKind.Absolute))
				{
                    // Here is support for loading ExternalLibraries from http.
                    Uri uri = new Uri(Path, UriKind.Absolute);
                    if (uri.IsUnc)
                    {
                        realPath = uri.LocalPath;
                        // Will continue to load later with the regular file load part below...
                    }
                    else
                    {
                        string scheme = uri.Scheme.ToLowerInvariant();
                        if (scheme != "http" && scheme != "file" && scheme != "https")
                        {
                            Logger.Initialization.Error("The ExternalLibrary path {0} is not a valid Uri scheme.", Path);
                            return list;
                        }
                        else
                        {
                            if (uri.AbsolutePath.EndsWith("dna", StringComparison.InvariantCultureIgnoreCase))
                            {
                                DnaLibrary lib = DnaLibrary.LoadFrom(uri);
                                if (lib == null)
                                {
                                    Logger.Initialization.Error("External library could not be registered - Path: {0} - DnaLibrary could not be loaded" + Path);
                                    return list;
                                }
                                // CONSIDER: Should we add a resolve story for .dna files at Uris?
                                return lib.GetAssemblies(null); // No explicit resolve path
                            }
                            else
                            {
                                // Load as a regular assembly - TypeLib not supported.
                                Assembly assembly = Assembly.LoadFrom(Path);
                                list.Add(new ExportedAssembly(assembly, ExplicitExports, ExplicitRegistration, ComServer, false, null, dnaLibrary));
                                return list;
                            }
                        }
                    }
                }
                // Keep trying with the current value of realPath.
                string resolvedPath = DnaLibrary.ResolvePath(realPath, pathResolveRoot);
                if (resolvedPath == null)
                {
                    Logger.Initialization.Error("External library could not be registered - Path: {0} - The library could not be found at this location" + Path);
                    return list;
				}
                if (System.IO.Path.GetExtension(resolvedPath).Equals(".DNA", StringComparison.OrdinalIgnoreCase))
				{
					// Load as a DnaLibrary
                    DnaLibrary lib = DnaLibrary.LoadFrom(resolvedPath);
					if (lib == null)
					{
                        Logger.Initialization.Error("External library could not be registered - Path: {0} - DnaLibrary could not be loaded" + Path);
                        return list;
					}

                    string pathResolveRelative = System.IO.Path.GetDirectoryName(resolvedPath);
					return lib.GetAssemblies(pathResolveRelative);
				}
				else
                {
                    Assembly assembly;
                    // Load as a regular assembly
                    // First check if it is already loaded (e.g. as a reference from another assembly)
                    // DOCUMENT: Some cases might still have assemblies loaded more than once.
                    // E.g. for an assembly that is both ExternalLibrary and references from another assembly,
                    // having the assembly LoadFromBytes and in the file system would load it twice, 
                    // because LoadFromBytes here happens before the .NET loaders assembly resolution.
                    string assemblyName = System.IO.Path.GetFileNameWithoutExtension(resolvedPath);
                    assembly = GetAssemblyIfLoaded(assemblyName);
                    if (assembly == null)
                    {
                        // Really have to load it.
                        if (LoadFromBytes)
                        {
                            // We need to be careful here to not re-load the assembly if it had already been loaded, 
                            // e.g. as a dependency of an assembly loaded earlier.
                            // In that case we won't be able to have the library 'LoadFromBytes'.
                            byte[] bytes = File.ReadAllBytes(resolvedPath);

                            string pdbPath = System.IO.Path.ChangeExtension(resolvedPath, "pdb");
                            if (File.Exists(pdbPath))
                            {
                                byte[] pdbBytes = File.ReadAllBytes(pdbPath);
                                assembly = Assembly.Load(bytes, pdbBytes);
                            }
                            else
                            {
                                assembly = Assembly.Load(bytes);
                            }
                        }
                        else
                        {
                            assembly = Assembly.LoadFrom(resolvedPath);
                        }
                    }
                    string resolvedTypeLibPath = null;
                    if (!string.IsNullOrEmpty(TypeLibPath))
                    {
                        resolvedTypeLibPath = DnaLibrary.ResolvePath(TypeLibPath, pathResolveRoot); // null is unresolved
                        if (resolvedTypeLibPath == null)
                        {
                            resolvedTypeLibPath = DnaLibrary.ResolvePath(TypeLibPath, System.IO.Path.GetDirectoryName(resolvedPath));
                        }
                    }
                    else
                    {
                        // Check for .tlb with same name next to resolvedPath
                        string tlbCheck = System.IO.Path.ChangeExtension(resolvedPath, "tlb");
                        if (System.IO.File.Exists(tlbCheck))
                        {
                            resolvedTypeLibPath = tlbCheck;
                        }
                    }
                    list.Add(new ExportedAssembly(assembly, ExplicitExports, ExplicitRegistration, ComServer, false, resolvedTypeLibPath, dnaLibrary));
					return list;
				}
			}
			catch (Exception e)
			{
				// Assembly could not be loaded.
				Logger.Initialization.Error(e, "External library could not be registered - Path: {0}", Path);
                return list;
			}
		}
        
        // A copy of this method lives in ExcelDna.Loader - AssemblyManager.cs
        private static Assembly GetAssemblyIfLoaded(string assemblyName)
        {
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (Assembly loadedAssembly in assemblies)
            {
                string loadedAssemblyName = loadedAssembly.FullName.Split(',')[0];
                if (string.Equals(assemblyName, loadedAssemblyName, StringComparison.OrdinalIgnoreCase))
                    return loadedAssembly;
            }
            return null;
        }
	}
}
