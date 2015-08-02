//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExcelDna.Integration
{
	// This class is used for the dynamic assembly references.
	// It manages the collection of assemblies that are referenced by each project,
	// and registers with the AppDomain to assist in assembly resolution.
	// I want it to work as I expected it to initially.
	// This means that the assembly found at the full path provided in the references section
	// when the dynamic assembly was compiled, 
	// will be resolved as an assembly when the dynamic code runs.
	// For now, it seems I cannot intercept the assembly resolution to prevent it from 
	// working in the usual way, but I intercept failures and fix up there.

	// TODO: This class might grow to also manage COM and DnaProject references.
	// CONSIDER: Fold into Reference class?

	public class AssemblyReference
	{
		public string Path;
		public Assembly Assembly;

		static Dictionary<string, AssemblyReference> AssemblyReferences = new Dictionary<string, AssemblyReference>();
		static AssemblyReference()
		{
			AppDomain.CurrentDomain.AssemblyResolve += Resolve;
		}

		public static void AddAssembly(string path)
		{
			if (path != null && !AssemblyReferences.ContainsKey(path) && File.Exists(path))
			{
				AssemblyReference ar = new AssemblyReference();
				ar.Path = path;
				ar.Assembly = Assembly.LoadFrom(path);
				AssemblyReferences.Add(path, ar);
			}
		}

		// Way to sort out the assembly resolve to an assembly that was referenced
		// but is not now accessible.
		// CONSIDER: How to do this better?
		internal static Assembly Resolve(object sender, ResolveEventArgs args)
		{
			foreach (AssemblyReference ar in AssemblyReferences.Values)
			{
				if (ar.Assembly.FullName == args.Name)
					return ar.Assembly;
			}
			return null;
		}
	}
}
