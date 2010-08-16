/*
  Copyright (C) 2005-2010 Govert van Drimmelen

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
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Xml.Serialization;

namespace ExcelDna.Integration
{
	[Serializable]
	[XmlType(AnonymousType = true)]
	public class Project
	{

        private string _Name;
        [XmlAttribute]
		public string Name
		{
			get { return _Name; }
			set { _Name = value; }
		}

        private string _Language;
        [XmlAttribute]
		public string Language
		{
			get { return _Language; }
			set { _Language = value; }
		}

		private string _CompilerVersion;
		[XmlAttribute]
		public string CompilerVersion
		{
			get { return _CompilerVersion; }
			set { _CompilerVersion = value; }
		}

        private List<Reference> _References;
        [XmlElement("Reference", typeof(Reference))]
		public List<Reference> References
		{
			get { return _References; }
			set	{ _References = value; }
		}

        // Sets whether references to System, System.Data and System.Xml are added automatically
        private bool _DefaultReferences = true;
        [XmlAttribute]
        public bool DefaultReferences
        {
            get { return _DefaultReferences; }
            set { _DefaultReferences = value; }
        }

        // Used for VB Projects only
        // Sets whether Imports to ExcelDna.Integration, 
        // Microsoft.VisualBasic, System, etc. added automatically
        private bool _DefaultImports = true;
        [XmlAttribute]
        public bool DefaultImports
        {
            get { return _DefaultImports; }
            set { _DefaultImports = value; }
        }

		private bool _ExplicitExports = false;
		[XmlAttribute]
		public bool ExplicitExports
		{
			get { return _ExplicitExports; }
			set { _ExplicitExports = value; }
		}

        private List<SourceItem> _SourceItems;
        [XmlElement("SourceItem", typeof(SourceItem))]
		public List<SourceItem> SourceItems
		{
			get { return _SourceItems; }
			set	{ _SourceItems = value; }
		}

        private string _Code;
        [XmlText]
        public string Code
        {
            get { return _Code; }
            set { _Code = value; }
        }

        internal Project()
        {
        }

        // Used by DnaLibrary to quickly make a default project
        internal Project(string language, string compilerVersion, List<Reference> references, string code,
                bool defaultReferences, bool defaultImports, bool explicitExports )
        {
            Language = language;
			CompilerVersion = compilerVersion;
            References = references;
            Code = code;
            DefaultReferences = defaultReferences;
            DefaultImports = defaultImports;
			ExplicitExports = explicitExports;
        }

        // Get projects explicit and implicitly present in the library
        private List<SourceItem> GetSourceItems()
        {
            List<SourceItem> sourceItems = new List<SourceItem>();
            if (SourceItems != null)
                sourceItems.AddRange(SourceItems);
            if (Code != null && Code.Trim() != "")
                sourceItems.Add(new SourceItem(Code));
            return sourceItems;
        }

        private List<string> tempAssemblyPaths = new List<string>();

        public List<Reference> GetReferences()
        {
            List<Reference> references = new List<Reference>();
			if (References != null)
			{
				foreach (Reference rf in References)
				{
					if (rf.AssemblyPath != null && rf.AssemblyPath.StartsWith("packed:"))
					{
						string assName = rf.AssemblyPath.Substring(7);
						string assPath = Path.GetTempFileName();
						tempAssemblyPaths.Add(assPath);
						File.WriteAllBytes(assPath, Integration.GetAssemblyBytes(assName));
						references.Add(new Reference(assPath));
					}
					else
					{
						references.Add(rf);
					}
				}
			}
            if (DefaultReferences)
            {
                references.Add(new Reference("System.dll"));
                references.Add(new Reference("System.Data.dll"));
                references.Add(new Reference("System.Xml.dll"));
            }
            // DOCUMENT: Reference to the xll is always added
            // Sort out "ExcelDna.Integration.dll"
            // TODO: URGENT: Revisit this...
            string location = Assembly.GetExecutingAssembly().Location;
            if (location != "")
            {
                references.Add(new Reference(location));
            }
            else
            {
				string assPath = Path.GetTempFileName();
                tempAssemblyPaths.Add(assPath);
                File.WriteAllBytes(assPath, Integration.GetAssemblyBytes("EXCELDNA.INTEGRATION"));
                references.Add(new Reference(assPath));
            }

			Debug.WriteLine("Compiler References: ");
			foreach (Reference rf in references)
			{
				Debug.WriteLine("\t" + rf.AssemblyPath);
			}

            return references;
        }

        // TODO: Move compilation stuff elsewhere.
		internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot /*currently unused - for references?*/)
		{
			List<ExportedAssembly> list = new List<ExportedAssembly>();
			// Dynamically compile this project to an in-memory assembly

			CodeDomProvider provider = GetProvider();
			if (provider == null)
				return list;

			CompilerParameters cp = new CompilerParameters();

			// TODO: Debug build ?
			// cp.IncludeDebugInformation = true;
			cp.GenerateExecutable = false;
			//cp.OutputAssembly = Name; // TODO: Keep track of built assembly for the project
			cp.GenerateInMemory = true;
			cp.TreatWarningsAsErrors = false;

			// This is attempt to fix the bug reported on the group, where the add-in compilation fails if the add-in is put into c:\
			// It is caused by a quirk of the 'Path.GetDirectoryName' function when dealing with the path "c:\test.abc" 
			// - it leaves the last DirectorySeparator in the path in this special case.
			// Thanks to Nemo for the great fix.
			//local variable to hold the quoted/unquoted version of the executing dirction
			string ProcessedExecutingDirectory = DnaLibrary.ExecutingDirectory;
            if (ProcessedExecutingDirectory.IndexOf(' ') != -1)
                ProcessedExecutingDirectory = "\"" + ProcessedExecutingDirectory + "\"";

			//set compiler command line vars as needed
            if (provider is Microsoft.VisualBasic.VBCodeProvider)
            {
                cp.CompilerOptions = " /libPath:" + ProcessedExecutingDirectory;
                if (DefaultImports)
                {
                    string importsList = "Microsoft.VisualBasic,System,System.Collections,System.Collections.Generic,System.Data,System.Diagnostics,ExcelDna.Integration";
                    cp.CompilerOptions += " /imports:" + importsList;
                }
            }
            else if ( (provider is Microsoft.CSharp.CSharpCodeProvider) || (provider.GetType().FullName.ToLower().IndexOf(".jscript.") != -1))
            {
                cp.CompilerOptions = " /lib:" + ProcessedExecutingDirectory;
            }
            else if (provider.GetType().FullName == "Microsoft.FSharp.Compiler.CodeDom.FSharpCodeProvider")
            {
                cp.CompilerOptions = " --nologo -I " + ProcessedExecutingDirectory;
            }

			List<Reference> references = GetReferences();
			List<string> refNames = new List<string>();
			foreach (Reference item in references)
			{
				if (item.AssemblyPath != null && item.AssemblyPath != "")
				{
					refNames.Add(item.AssemblyPath);
				}
				if (item.Name != null && item.Name != "")
				{
					refNames.Add(item.Name);
				}
			}

			cp.ReferencedAssemblies.AddRange(refNames.ToArray());

            List<string> sources = GetSourceItems().ConvertAll<string>(delegate(SourceItem item) { return item.Code.Trim(); });
			CompilerResults cr = provider.CompileAssemblyFromSource(cp, sources.ToArray());

			foreach (string path in tempAssemblyPaths)
			{
				File.Delete(path);
			}
			tempAssemblyPaths.Clear();

			if (cr.Errors.HasErrors)
			{
                ExcelDna.Logging.LogDisplay.WriteLine("There were errors when compiling project: " + Name);
				foreach (CompilerError err in cr.Errors)
				{
                    ExcelDna.Logging.LogDisplay.WriteLine(err.ToString());
				}
				return list;
			}

			// Success !!
			// Now add all the references
			// TODO: How to remove again??
			foreach (Reference r in References)
			{
				AssemblyReference.AddAssembly(r.AssemblyPath);
			}

			list.Add(new ExportedAssembly(cr.CompiledAssembly, ExplicitExports));
			return list;
		}

		private CodeDomProvider GetProvider()
		{
			// DOCUMENT: Currently accepted languages: 
			// CS / CSHARP / C# / VB / VISUAL BASIC / VISUALBASIC / FS /F# / FSHARP / F SHARP
			// or a fully qualified TypeName that derives from CodeDomProvider
			// DOCUMENT: CompilerVersion usage

			Dictionary<string, string> providerOptions = null; 
			if (!string.IsNullOrEmpty(CompilerVersion))
			{
				providerOptions = new Dictionary<string, string>();
				providerOptions.Add("CompilerVersion", CompilerVersion);
			}

            string lang;
            if (Language == null)
                lang = "vb";
            else
			    lang = Language.ToLower();

			if (lang == "cs" || lang == "csharp" || lang == "c#" || lang == "c sharp")
			{
				if (providerOptions == null)
				{
					return new Microsoft.CSharp.CSharpCodeProvider();
				}
				else
				{
					Assembly sys = Assembly.GetAssembly(typeof(Microsoft.CSharp.CSharpCodeProvider));
					return (CodeDomProvider)sys.CreateInstance("Microsoft.CSharp.CSharpCodeProvider", false, BindingFlags.CreateInstance, null, new object[] {providerOptions}, null, null);
				}
			}
			else if (lang == "vb" || lang == "visual basic" || lang == "visualbasic")
			{
				if (providerOptions == null)
				{
					return new Microsoft.VisualBasic.VBCodeProvider();
				}
				else
				{
					Assembly sys = Assembly.GetAssembly(typeof(Microsoft.VisualBasic.VBCodeProvider));
					return (CodeDomProvider)sys.CreateInstance("Microsoft.VisualBasic.VBCodeProvider", false, BindingFlags.CreateInstance, null, new object[] {providerOptions}, null, null);
				}
			}
            else if (lang == "fs" || lang == "fsharp" || lang == "f#" || lang == "f sharp")
            {
                try
                {
                    // TODO: Reconsider how to support F#
                    Assembly fsharp = Assembly.LoadWithPartialName("FSharp.Compiler.CodeDom" );
                    return (CodeDomProvider)fsharp.CreateInstance("Microsoft.FSharp.Compiler.CodeDom.FSharpCodeProvider");
                }
                catch
                {
                    // TODO: Log this error to the display?
                    Debug.Fail("FSharp.Compiler.CodeDom could not be loaded.");
                    return null;
                }
            }

			// Else try to load the language as a type
			// TODO: Test this !?
			try
			{
				Type t = Type.GetType(Language);
				if (t.IsSubclassOf(typeof(CodeDomProvider)))
				{
					ConstructorInfo ci = t.GetConstructor(new Type[] {} );
					CodeDomProvider p = (CodeDomProvider)ci.Invoke(new object[] { });
					return p;
				}

				return null;
			}
			catch (Exception e)
			{
				Debug.Fail("Unknown Project Language: " + Language);
				Logging.LogDisplay.WriteLine("Unknown Project Language: " + Language + " Exception: " + e.Message);
			}
			return null;
		}
	}
}
