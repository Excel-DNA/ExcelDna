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
        internal DnaLibrary ParentDnaLibrary;

		internal ExportedAssembly(Assembly assembly, bool explicitExports, DnaLibrary parentDnaLibrary)
		{
			this.Assembly = assembly;
			this.ExplicitExports = explicitExports;
            this.ParentDnaLibrary = parentDnaLibrary;
		}
	}
}
