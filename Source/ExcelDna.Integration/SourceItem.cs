//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
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
