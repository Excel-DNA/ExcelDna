/*
  Copyright (C) 2005, 2006 Govert van Drimmelen

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
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Xml.Serialization;

namespace ExcelDna.Integration
{
	[Serializable]
	[XmlType(AnonymousType = true)]
	public class Project : IAssemblyDefinition
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
        internal Project(string language, List<Reference> references, string code,
                bool defaultReferences, bool defaultImports )
        {
            Language = language;
            References = references;
            Code = code;
            DefaultReferences = defaultReferences;
            DefaultImports = defaultImports;
        }

        // Get projects explicit and implicitly prosent in the library
        private List<SourceItem> GetSourceItems()
        {
            List<SourceItem> sourceItems = new List<SourceItem>();
            if (SourceItems != null)
                sourceItems.AddRange(SourceItems);
            if (Code != null && Code.Trim() != "")
                sourceItems.Add(new SourceItem(Code));
            return sourceItems;
        }

        public List<Reference> GetReferences()
        {
            List<Reference> references = new List<Reference>();
            if (References != null)
                references.AddRange(References);
            if (DefaultReferences)
            {
                references.Add(new Reference("System.dll"));
                references.Add(new Reference("System.Data.dll"));
                references.Add(new Reference("System.Xml.dll"));
            }
            // DOCUMENT: Reference to the xll is always added
            references.Add(new Reference((string)XlCall.Excel(XlCall.xlGetName)));
            return references;
        }

        // TODO: Move compilation stuff elsewhere.
		public List<Assembly> GetAssemblies()
		{
			List<Assembly> list = new List<Assembly>();
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

            if (provider is Microsoft.VisualBasic.VBCodeProvider && DefaultImports)
            {
                string importsList = "Microsoft.VisualBasic,System,System.Collections,System.Collections.Generic,System.Data,System.Diagnostics,ExcelDna.Integration";
                cp.CompilerOptions = "/imports:" + importsList;
            }

            List<string> references = GetReferences().ConvertAll<string>(delegate(Reference item) { return item.AssemblyPath; });
			cp.ReferencedAssemblies.AddRange(references.ToArray());

            List<string> sources = GetSourceItems().ConvertAll<string>(delegate(SourceItem item) { return item.Code; });
			CompilerResults cr = provider.CompileAssemblyFromSource(cp, sources.ToArray());

			if (cr.Errors.HasErrors)
			{
				// TODO: Manage Errors / Output
				StringBuilder Errors = new StringBuilder();
				Errors.AppendLine("There were errors when compiling project: " + Name);
				foreach (CompilerError err in cr.Errors)
				{
					Errors.AppendLine(err.ToString());
				}
                ErrorDisplay.DisplayErrorMessage(Errors.ToString());
				return list;
			}

			// Success !!
			// Now add all the references
			// TODO: How to remove again??
			foreach (Reference r in References)
			{
				AssemblyReference.AddAssembly(r.AssemblyPath);
			}

			list.Add(cr.CompiledAssembly);
			return list;
		}

		private CodeDomProvider GetProvider()
		{
			// DOCUMENT: Currently accepted languages: 
			// CS / CSHARP / C#, VB, VISUAL BASIC, VISUALBASIC
			// or a fully qualified TypeName that derives from CodeDomProvider
            string lang;
            if (Language == null)
                lang = "vb";
            else
			    lang = Language.ToLower();

			if (lang == "cs" || lang == "csharp" || lang == "c#" )
			{
				return new Microsoft.CSharp.CSharpCodeProvider();
			}
			else if (lang == "vb" || lang == "visual basic" || lang == "visualbasic")
			{
				return new Microsoft.VisualBasic.VBCodeProvider();
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
			catch
			{
				Debug.Fail("Unknown Project Language: " + Language);
			}
			return null;
		}
	}
}
