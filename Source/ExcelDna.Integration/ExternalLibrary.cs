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
	public class ExternalLibrary : IAssemblyDefinition
	{
		[XmlAttribute]
		public string Path;

		public List<Assembly> GetAssemblies()
		{
			try
			{
				string realPath = Path;
				if (!File.Exists(realPath))
				{
                    string xllName = DnaLibrary.XllPath;
					string localDirectory = System.IO.Path.GetDirectoryName(xllName);
					if (System.IO.Path.IsPathRooted(realPath))
					{
						// Rooted path -- try locally
						string fileName = System.IO.Path.GetFileName(realPath);
						realPath = System.IO.Path.Combine(localDirectory, fileName);
					}
					else
					{
						// Try a path relative to local directory
						realPath = System.IO.Path.Combine(localDirectory, realPath);
					}
					// Check again
					if (!File.Exists(realPath))
					{
						// Give up.
						Debug.Print("Could not find file " + Path);
						return new List<Assembly>();
					}
				}
				if (System.IO.Path.GetExtension(realPath).ToLower() == ".dna")
				{
					// Load as a DnaLibrary
					DnaLibrary lib = DnaLibrary.LoadFrom(realPath);
                    if (lib == null)
                    {
                        // Problems during load.
                        return new List<Assembly>();
                    }

                    return lib.GetAssemblies();
				}
				else
				{
					// Load as a regular assembly
					List<Assembly> list = new List<Assembly>();
					list.Add(Assembly.LoadFrom(realPath));
					return list;
				}
			}
			catch (Exception e)
			{
				// Assembly could not be loaded.
				Debug.Print("Assembly load exception for file: " + Path + "\n" + e.ToString());
				return new List<Assembly>();
			}
		}
	}
}
