//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
