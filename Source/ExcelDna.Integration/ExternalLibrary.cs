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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
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

		internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot, DnaLibrary dnaLibrary)
		{
			List<ExportedAssembly> list = new List<ExportedAssembly>();

			try
			{
                string realPath = Path;
				if (Path.StartsWith("packed:"))
				{
					string resourceName = Path.Substring(7);
					if (Path.ToUpperInvariant().EndsWith(".DNA"))
					{
						byte[] dnaContent = Integration.GetDnaFileBytes(resourceName);
						DnaLibrary lib = DnaLibrary.LoadFrom(dnaContent, pathResolveRoot);
						if (lib == null)
						{
                            LogDisplay.WriteLine("External library could not be registered - Path: " + Path);
                            LogDisplay.WriteLine("    Error: Packed DnaLibrary could not be loaded.");
							return list;
						}

						return lib.GetAssemblies(pathResolveRoot);
					}
					else
					{
						byte[] rawAssembly = Integration.GetAssemblyBytes(resourceName);
                        // DOCUMENT: TypeLibPath which is a resource in a library is denoted as fileName.dll\4
                        // For packed assemblies, we set TypeLibPath="packed:2"
                        string typeLibPath = null;
                        if (!string.IsNullOrEmpty(TypeLibPath) && TypeLibPath.StartsWith("packed:"))
                        {
                            typeLibPath = DnaLibrary.XllPath + @"\" + TypeLibPath.Substring(7);
                        }
						list.Add(new ExportedAssembly(Assembly.Load(rawAssembly), ExplicitExports, ComServer, false, typeLibPath, dnaLibrary));
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
                            Logging.LogDisplay.WriteLine("The ExternalLibrary path {0} is not a valid Uri scheme.", Path);
                            return list;
                        }
                        else
                        {
                            if (uri.AbsolutePath.EndsWith("dna", StringComparison.InvariantCultureIgnoreCase))
                            {
                                DnaLibrary lib = DnaLibrary.LoadFrom(uri);
                                if (lib == null)
                                {
                                    LogDisplay.WriteLine("External library could not be registered - Path: " + Path);
                                    LogDisplay.WriteLine("    Error: DnaLibrary could not be loaded.");
                                    return list;
                                }
                                // CONSIDER: Should we add a resolve story for .dna files at Uris?
                                return lib.GetAssemblies(null); // No explicit resolve path
                            }
                            else
                            {
                                // Load as a regular assembly - TypeLib not supported.
                                list.Add(new ExportedAssembly(Assembly.LoadFrom(Path), ExplicitExports, ComServer, false, null, dnaLibrary));
                                return list;
                            }
                        }
                    }
                }
                // Keep trying with the current value of realPath.
                string resolvedPath = DnaLibrary.ResolvePath(realPath, pathResolveRoot);
                if (resolvedPath == null)
                {
                    LogDisplay.WriteLine("External library could not be registered - Path: " + Path);
                    LogDisplay.WriteLine("    Error: The library could not be found at this location.");
                    return list;
				}
                if (System.IO.Path.GetExtension(resolvedPath).ToUpperInvariant() == ".DNA")
				{
					// Load as a DnaLibrary
                    DnaLibrary lib = DnaLibrary.LoadFrom(resolvedPath);
					if (lib == null)
					{
                        LogDisplay.WriteLine("External library could not be registered - Path: " + Path);
                        LogDisplay.WriteLine("    Error: DnaLibrary could not be loaded.");
                        return list;
					}

                    string pathResolveRelative = System.IO.Path.GetDirectoryName(resolvedPath);
					return lib.GetAssemblies(pathResolveRelative);
				}
				else
				{
                    Assembly assembly;
					// Load as a regular assembly
                    if (LoadFromBytes)
                    {
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
                    string resolvedTypeLibPath = null;
                    if (!string.IsNullOrEmpty(TypeLibPath))
                    {
                        resolvedTypeLibPath = DnaLibrary.ResolvePath(TypeLibPath, pathResolveRoot); // null is unresolved
                        if (resolvedTypeLibPath == null && resolvedPath != null)
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
                    list.Add(new ExportedAssembly(assembly, ExplicitExports, ComServer, false, resolvedTypeLibPath, dnaLibrary));
					return list;
				}
			}
			catch (Exception e)
			{
				// Assembly could not be loaded.
				LogDisplay.WriteLine("External library could not be registered - Path: " + Path);
                LogDisplay.WriteLine("    Error: " + e.Message);
                return list;
			}
		}
	}
}
