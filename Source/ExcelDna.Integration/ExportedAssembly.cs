//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Reflection;

namespace ExcelDna.Integration
{
	internal class ExportedAssembly
	{
		internal Assembly Assembly;
        internal bool     ExplicitExports;
        internal bool     ExplicitRegistration;
        internal bool     ComServer;
        internal bool     IsDynamic;
        internal string   TypeLibPath;
        internal DnaLibrary ParentDnaLibrary;

		internal ExportedAssembly(Assembly assembly, bool explicitExports, bool explicitRegistration, bool comServer, bool isDynamic, string typeLibPath, DnaLibrary parentDnaLibrary)
		{
			this.Assembly = assembly;
            this.ExplicitExports = explicitExports;
            this.ExplicitRegistration = explicitRegistration;
            this.ComServer = comServer;
            this.IsDynamic = isDynamic;
            this.TypeLibPath = typeLibPath;
            this.ParentDnaLibrary = parentDnaLibrary;
		}
	}
}
