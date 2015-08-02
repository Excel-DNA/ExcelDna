//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using ExcelDna.Logging;

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

        private bool _ExplicitRegistration = false;
        [XmlAttribute]
        public bool ExplicitRegistration
        {
            get { return _ExplicitRegistration; }
            set { _ExplicitRegistration = value; }
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
                        Logger.DnaCompilation.Error("Assembly resolve failure - Reference Name: {0}, Path: {1}", rf.Name, rf.Path);
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

            Logger.DnaCompilation.Verbose("Compiler References: ");
			foreach (string rfPath in refPaths)
			{
                Logger.DnaCompilation.Verbose("    " + rfPath);
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
            bool isFsharp = false;

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
                isFsharp = true;
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

            CompilerResults cr;
            try
            {
                cr = provider.CompileAssemblyFromSource(cp, sources.ToArray());
            }
            catch (Win32Exception wex)
            {
                if (isFsharp)
                {
                    Logger.DnaCompilation.Error("There was an error in loading the add-in " + DnaLibrary.CurrentLibraryName + " (" + DnaLibrary.XllPath + "):");
                    string fsBinPath = Environment.GetEnvironmentVariable("FSHARP_BIN");
                    string msg;
                    if (fsBinPath == null)
                    {
                        msg = "    Calling the F# compiler failed (\"" + wex.Message + "\").\r\n" + 
                              "    Please check that the F# compiler is correctly installed.\r\n" + 
                              "    This error can sometimes be fixed by creating an FSHARP_BIN environment variable.\r\n" +
                              "    Create an environment variable FSHARP_BIN with the full path to the directory containing \r\n" +
                              "    the F# compiler fsc.exe - for example \r\n" +
                              "        \"" + @"C:\Program Files (x86)\Microsoft SDKs\F#\3.0\Framework\v4.0\""";
                    }
                    else
                    {
                        msg = "    Calling the F# compiler failed (\"" + wex.Message + "\").\r\n" +
                              "    Please check that the F# compiler is correctly installed, and that the FSHARP_BIN environment variable is correct\r\n" +
                              "    (it currently points to " + fsBinPath + ").";

                    }
                    Logger.DnaCompilation.Error(msg);
                    return list;
                }
                throw;
            }
            

			foreach (string path in tempAssemblyPaths)
			{
				File.Delete(path);
			}
			tempAssemblyPaths.Clear();
            

			if (cr.Errors.HasErrors)
			{
                Logger.DnaCompilation.Error("There was an error in loading the add-in " + DnaLibrary.CurrentLibraryName + " (" + DnaLibrary.XllPath + "):");
                Logger.DnaCompilation.Error("There were errors when compiling project: " + Name);
				foreach (CompilerError err in cr.Errors)
				{
                    Logger.DnaCompilation.Error("    " + err.ToString());
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
			list.Add(new ExportedAssembly(cr.CompiledAssembly, ExplicitExports, ExplicitRegistration, ComServer, true, null, dnaLibrary));
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
                        Logger.DnaCompilation.Error("There was an error in loading the add-in " + DnaLibrary.CurrentLibraryName + " (" + DnaLibrary.XllPath + "):");
                        Logger.DnaCompilation.Error("    The F# CodeDom provider (FSharp.Compiler.CodeDom.dll) could not be loaded.");
                        Logger.DnaCompilation.Error("        Please ensure that the F# Compiler is installed and that the");
                        Logger.DnaCompilation.Error("        FSharp.Compiler.CodeDom.dll assembly (part of the F# PowerPack) can be loaded by the add-in.");
                        Logger.DnaCompilation.Error("        (Placing a copy of FSharp.Compiler.CodeDom.dll in the same directory as the .xll file should work.)");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Logger.DnaCompilation.Error("There was an error in loading the add-in " + DnaLibrary.CurrentLibraryName + " (" + DnaLibrary.XllPath + "):");
                    Logger.DnaCompilation.Error("Error in loading the F# CodeDom provider.");
                    Logger.DnaCompilation.Error("    Exception: " + ex.Message);
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
				Logger.DnaCompilation.Error("Unknown Project Language: {0}", Language);
                Logger.DnaCompilation.Error(e, "   CodeDomProvider load error");
			}
			return null;
		}
	}
}
