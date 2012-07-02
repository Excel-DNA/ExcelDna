/*
  Copyright (C) 2005-2012 Govert van Drimmelen

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
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.IO;
using System.Diagnostics;

namespace ExcelDna.Integration
{
	[Serializable]
	[XmlType(AnonymousType = true)]
	public class SourceItem
	{
		private string _Name;
		[XmlAttribute]
		public string Name
		{
			get { return _Name; }
			set { _Name = value; }
		}

		private string _Code;
		[XmlText]
		public string Code
		{
			get { return _Code; }
			set	{ _Code = value; }
		}

        private string _Path;
        [XmlAttribute]
        public string Path
        {
            get { return _Path; }
            set { _Path = value; }
        }

        [XmlAttribute]
        public bool Pack;

        // Returns the resulting source for this SourceItem.
        // If Path is filled in, and file exists, takes source from there.
        // Else returns Code value.
        public string GetSource(string pathResolveRoot)
        {
            if (!string.IsNullOrEmpty(Path))
            {
                if (Path.StartsWith("packed:"))
                {
                    string resourceName = Path.Substring(7);
                    byte[] sourceBytes = ExcelIntegration.GetSourceBytes(resourceName);
                    return Encoding.UTF8.GetString(sourceBytes);
                }
                else
                {
                    // Try to read from a file.
                    string resolvedPath = DnaLibrary.ResolvePath(Path, pathResolveRoot);
                    if (resolvedPath == null)
                    {
                        Debug.Print("Source path {0} could not be resolved.", Path);
                    }
                    else
                    {
                        return File.ReadAllText(resolvedPath).Trim();
                    }
                }
            }
            return Code;
        }

        public SourceItem()
        {
        }

        internal SourceItem(string code)
        {
            Code = code;
        }
	}
}
