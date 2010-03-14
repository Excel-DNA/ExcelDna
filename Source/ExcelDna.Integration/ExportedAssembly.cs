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

		internal ExportedAssembly(Assembly assembly, bool explicitExports)
		{
			this.Assembly = assembly;
			this.ExplicitExports = explicitExports;
		}
	}
}
