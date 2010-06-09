/*
  Copyright (C) 2005-2010 Govert van Drimmelen

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

namespace ExcelDna.Integration
{
	// TODO: Allow Com References/Exported Libraries
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

		private bool _Pack = false;
		[XmlAttribute]
		public bool Pack
		{
			get { return _Pack; }
			set { _Pack = value; }
		}

		private bool _ExplicitExports = false;
		[XmlAttribute]
		public bool ExplicitExports
		{
			get { return _ExplicitExports; }
			set { _ExplicitExports = value; }
		}

		internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot)
		{
			List<ExportedAssembly> list = new List<ExportedAssembly>();

			try
			{
				if (Path.StartsWith("packed:"))
				{
					string resourceName = Path.Substring(7);
					if (Path.ToUpperInvariant().EndsWith(".DNA"))
					{
						byte[] dnaContent = Integration.GetDnaFileBytes(resourceName);
						DnaLibrary lib = DnaLibrary.LoadFrom(dnaContent, pathResolveRoot);
						if (lib == null)
						{
							// Problems during load.
							return list;
						}

						return lib.GetAssemblies(pathResolveRoot);
					}
					else
					{
						byte[] rawAssembly = Integration.GetAssemblyBytes(resourceName);
						list.Add(new ExportedAssembly(Assembly.Load(rawAssembly), ExplicitExports));
						return list;
					}
				}
				else
				{
					string realPath = Path;

                    string resolvedPath = DnaLibrary.ResolvePath(realPath, pathResolveRoot);
                    if (resolvedPath == null)
                    {
                        // Give up.
					    return list;
					}
                    if (System.IO.Path.GetExtension(resolvedPath).ToUpperInvariant() == ".DNA")
					{
						// Load as a DnaLibrary
                        DnaLibrary lib = DnaLibrary.LoadFrom(resolvedPath);
						if (lib == null)
						{
							// Problems during load.
							return list;
						}

                        string pathResolveRelative = System.IO.Path.GetDirectoryName(resolvedPath);
						return lib.GetAssemblies(pathResolveRelative);
					}
					else
					{
						// Load as a regular assembly
						// CONSIDER: Rather load into the Load context?
                        list.Add(new ExportedAssembly(Assembly.LoadFrom(resolvedPath), ExplicitExports));
						return list;
					}
				}
			}
			catch (Exception e)
			{
				// Assembly could not be loaded.
				Debug.Print("Assembly load exception for file: " + Path + "\n" + e.ToString());
				return list;
			}
		}
	}
}
