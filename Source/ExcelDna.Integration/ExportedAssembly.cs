using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace ExcelDna.Integration
{
	internal class ExportedAssembly
	{
		internal Assembly Assembly;
		internal bool     ExplicitExports;
        internal bool     ComServer;
        internal bool     IsDynamic;
        internal string   TypeLibPath;
        internal DnaLibrary ParentDnaLibrary;

		internal ExportedAssembly(Assembly assembly, bool explicitExports, bool comServer, bool isDynamic, string typeLibPath, DnaLibrary parentDnaLibrary)
		{
			this.Assembly = assembly;
			this.ExplicitExports = explicitExports;
            this.ComServer = comServer;
            this.IsDynamic = isDynamic;
            this.TypeLibPath = typeLibPath;
            this.ParentDnaLibrary = parentDnaLibrary;
		}
	}
}
