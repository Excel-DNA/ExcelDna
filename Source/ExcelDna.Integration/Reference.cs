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
using System.Text;
using System.Xml.Serialization;

namespace ExcelDna.Integration
{
	[Serializable]
	[XmlType(AnonymousType = true)]
	public class Reference
	{
        // DOCUMENT: AssemblyPath is obsolete, and only used as Path if the Path itself is null.
        // If Path is empty or cannot be resolved to a real file, LoadWithPartialName is called with Name, so that Name="System.Windows.Forms" and "Microsoft.Office.Interop.Excel" will work.

        [XmlAttribute]
        public string Name;

        // [Obsolete("Please use Path attribute.")]
        // (Can't mark it as Obsolete since serilizer will then ignore, breaking backward compatibility.)
        [XmlAttribute]
        public string AssemblyPath;

        private string _path;
        [XmlAttribute]
        public string Path
        {
            get
            {
                if (_path == null)
                {
                    return AssemblyPath;
                }
                return _path;
            }
            set
            {
                _path = value;
            }
        }

		[XmlAttribute]
		public bool Pack;

        public Reference()
        {
        }

        public Reference(string path)
        {
            Path = path;
        }
	}

    [Serializable]
    [XmlType(AnonymousType = true)]
    public class Image
    {
        [XmlAttribute]
        public string Name;

        [XmlAttribute]
        public string Path;

        [XmlAttribute]
        public bool Pack;

        public Image()
        {
        }
    }

}
