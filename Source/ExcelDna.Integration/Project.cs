/*
  Copyright (C) 2005-2012 Govert van Drimmelen

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
using System.Runtime.InteropServices;
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

        private bool _ComServer = false;
        [XmlAttribute]
        public bool ComServer
        {
            get { return _ComServer; }
            set { _ComServer = value; }
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
        private List<string> GetSources(string pathResolveRoot)
        {
            List<string> sources = new List<string>();
            if (SourceItems != null)
            {
                // Deal with explicit SourceItem tags.
                foreach (SourceItem sourceItem in SourceItems)
                {
                    string source = sourceItem.GetSource(pathResolveRoot);
                    if (!string.IsNullOrEmpty(source))
                    {
                        sources.Add(source.Trim());
                    }
                }
            }
            // Implicit Source
            if (Code != null && Code.Trim() != "")
            {
                sources.Add(Code.Trim());
            }
            return sources;
        }

        private List<string> tempAssemblyPaths = new List<string>();

        public List<string> GetReferencePaths(string pathResolveRoot, CodeDomProvider provider)
        {
            List<string> refPaths = new List<string>();
			if (References != null)
			{
				foreach (Reference rf in References)
				{
                    bool isResolved = false;
                    if (rf.Path != null)
                    {
                        if (rf.Path.StartsWith("packed:"))
                        {
                            string assName = rf.Path.Substring(7);
                            string assPath = Path.GetTempFileName();
                            tempAssemblyPaths.Add(assPath);
                            File.WriteAllBytes(assPath, ExcelIntegration.GetAssemblyBytes(assName));
                            refPaths.Add(assPath);
                            isResolved = true;
                        }
                        else
                        {
                            // Call ResolvePath - check relative to pathResolveRoot and in framework directory.
                            string refPath = DnaLibrary.ResolvePath(rf.Path, pathResolveRoot);
                            if (!string.IsNullOrEmpty(refPath))
                            {
                                refPaths.Add(refPath);
                                isResolved = true;
                            }
                        }
                    }
                    if (!isResolved && rf.Name != null)
                    {
                        // Try to resolve by Name
                        #pragma warning disable 0618
                        Assembly refAssembly = Assembly.LoadWithPartialName(rf.Name);
                        #pragma warning restore 0618
                        if (refAssembly != null)
                        {
                            if (!string.IsNullOrEmpty(refAssembly.Location))
                            {
                                refPaths.Add(refAssembly.Location);
                                isResolved = true;
                            }
                        }
                    }
                    if (!isResolved)
                    {
                        // Must have been loaded by us from the packing....?
                        Debug.Print("Assembly resolve failure - Reference Name: {0}, Path: {1}", rf.Name, rf.Path);
                    }

				}
			}
            if (DefaultReferences)
            {
                // CONSIDER: Should these be considered more carefully? I'm just putting in what the default templates in Visual Studio 2010 put in.
                refPaths.Add("System.dll");
                refPaths.Add("System.Data.dll");
                refPaths.Add("System.Xml.dll");
                if (Environment.Version.Major >= 4)
                {
                    refPaths.Add("System.Core.dll");
                    refPaths.Add("System.Data.DataSetExtensions.dll");
                    refPaths.Add("System.Xml.Linq.dll");
                    if (provider is Microsoft.CSharp.CSharpCodeProvider)
                    {
                        refPaths.Add("Microsoft.CSharp.dll");
                    }
                }
            }
            // DOCUMENT: Reference to the xll is always added
            // Sort out "ExcelDna.Integration.dll" copy which causes problems
            // TODO: URGENT: Revisit this... (don't know what to do yet - maybe full AssemblyName).
            string location = Assembly.GetExecutingAssembly().Location;
            if (location != "")
            {
                refPaths.Add(location);
            }
            else
            {
				string assPath = Path.GetTempFileName();
                tempAssemblyPaths.Add(assPath);
                File.WriteAllBytes(assPath, ExcelIntegration.GetAssemblyBytes("EXCELDNA.INTEGRATION"));
                refPaths.Add(assPath);
            }

			Debug.WriteLine("Compiler References: ");
			foreach (string rfPath in refPaths)
			{
				Debug.WriteLine("\t" + rfPath);
			}

            return refPaths;
        }

        // TODO: Move compilation stuff elsewhere.
        // TODO: Consider IronPython support: http://www.ironpython.info/index.php/Using_Compiled_Python_Classes_from_.NET/CSharp_IP_2.6
		internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot, DnaLibrary dnaLibrary)
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
                    if (Environment.Version.Major >= 4)
                    {
                        importsList += ",System.Linq,System.Xml.Linq";
                    }
                    cp.CompilerOptions += " /imports:" + importsList;
                }
            }
            else if (provider is Microsoft.CSharp.CSharpCodeProvider) 
            {
                cp.CompilerOptions = " /lib:" + ProcessedExecutingDirectory;
            }
            else if (provider.GetType().FullName.ToLower().IndexOf(".jscript.") != -1)
            {
                cp.CompilerOptions = " /lib:" + ProcessedExecutingDirectory;
            }
            else if (provider.GetType().FullName == "Microsoft.FSharp.Compiler.CodeDom.FSharpCodeProvider")
            {
                cp.CompilerOptions = " --nologo -I " + ProcessedExecutingDirectory; // In F# 2.0, the --nologo is redundant - I leave it because it does no harm.

                // FSharp 2.0 compiler will target .NET 4 unless we do something to ensure .NET 2.0.
                // It seems adding an explicit reference to the .NET 2.0 version of mscorlib.dll is good enough.
                if (Environment.Version.Major < 4)
                {
                    // Explicitly add a reference to the mscorlib version from the currently running .NET version
                    string libPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "mscorlib.dll");
                    cp.ReferencedAssemblies.Add(libPath);
                }
            }

            // TODO: Consider what to do if we can't resolve some of the Reference paths -- do we try to compile anyway, throw an exception, ...what?
			List<string> refPaths = GetReferencePaths(pathResolveRoot, provider);
			cp.ReferencedAssemblies.AddRange(refPaths.ToArray());

            List<string> sources = GetSources(pathResolveRoot);
            
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
				AssemblyReference.AddAssembly(r.Path);
			}

            // TODO: Create TypeLib for execution-time compiled assemblies.
			list.Add(new ExportedAssembly(cr.CompiledAssembly, ExplicitExports, ComServer, true, null, dnaLibrary));
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

                    // This is my best plan to attempt 'future' compatibility.
                    #pragma warning disable 0618
                    Assembly fsharp = Assembly.LoadWithPartialName("FSharp.Compiler.CodeDom" );
                    #pragma warning restore 0618
                    if (fsharp != null)
                    {
                        return (CodeDomProvider)fsharp.CreateInstance("Microsoft.FSharp.Compiler.CodeDom.FSharpCodeProvider");
                    }
                    else
                    {
                        Logging.LogDisplay.WriteLine("The F# code provider could not be loaded.");
                        Logging.LogDisplay.WriteLine("Please ensure that both the F# Compiler and the F# PowerPack are installed.");
                        Logging.LogDisplay.WriteLine("    The F# Compiler (August 2010 CTP) can be found at: http://go.microsoft.com/fwlink/?LinkId=151924.");
                        Logging.LogDisplay.WriteLine("    The F# PowerPack can be found at: http://fsharppowerpack.codeplex.com/.");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Logging.LogDisplay.WriteLine("Error while loading F# code provider.");
                    Logging.LogDisplay.WriteLine(" Exception: " + ex.Message);
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
				Logging.LogDisplay.WriteLine("Unknown Project Language: " + Language);
                Logging.LogDisplay.WriteLine(" Exception: " + e.Message);
			}
			return null;
		}
	}
}
