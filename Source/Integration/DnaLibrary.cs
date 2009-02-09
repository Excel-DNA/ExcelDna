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
	[XmlRoot(Namespace = "", IsNullable = false)]
	public class DnaLibrary : IAssemblyDefinition
	{
		[XmlIgnore]
		public string Name;
		[XmlIgnore]
		public string FileName;


		private List<ExternalLibrary> _ExternalLibraries;
		[XmlElement("ExternalLibrary", typeof(ExternalLibrary))]
        public List<ExternalLibrary> ExternalLibraries
		{
			get { return _ExternalLibraries; }
			set	{ _ExternalLibraries = value; }
		}

		private List<Project> _Projects;
		[XmlElement("Project", typeof(Project))]
        public List<Project> Projects
		{
			get { return _Projects; }
			set { _Projects = value; }
		}

		private string _Description;
		[XmlAttribute]
		public string Description
		{
			get
			{
				if (_Description == null || _Description == "")
				{
					string dllName = Assembly.GetExecutingAssembly().Location;
					string xllFileRoot = Path.GetFileNameWithoutExtension(dllName);
					_Description = xllFileRoot + " (ExcelDna)";
				}
				return _Description;
			}
			set { _Description = value; }
		}

        // Next three are abbreviations for Project contents
        private List<Reference> _References;
        [XmlElement("Reference", typeof(Reference))]
        public List<Reference> References
        {
            get { return _References; }
            set { _References = value; }
        }

        private string _Language; // Default is VB
        [XmlAttribute]
        public string Language
        {
            get { return _Language; }
            set { _Language = value; }
        }

        private bool _DefaultReferences = true;
        [XmlAttribute]
        public bool DefaultReferences
        {
            get { return _DefaultReferences; }
            set { _DefaultReferences = value; }
        }

        private bool _DefaultImports = true;
        [XmlAttribute]
        public bool DefaultImports
        {
            get { return _DefaultImports; }
            set { _DefaultImports = value; }
        }
        
        private string _Code;
        [XmlText]
        public string Code
        {
            get { return _Code; }
            set { _Code = value; }
        }

        // Get projects explicit and implicitly present in the library
        private List<Project> GetProjects()
        {
            List<Project> projects = new List<Project>();
            if (Projects != null)
                projects.AddRange(Projects);
            if (_Code != null && Code.Trim() != "")
                projects.Add(new Project(Language, References, _Code, DefaultReferences, DefaultImports));
            return projects;
        }

        private List<Assembly> _Assemblies;
		public List<Assembly> GetAssemblies()
		{
			if (_Assemblies == null)
			{
				_Assemblies = new List<Assembly>();
				List<IAssemblyDefinition> assemblyDefs = new List<IAssemblyDefinition>();
				if (ExternalLibraries != null)
					assemblyDefs.AddRange(ExternalLibraries.ToArray());
				assemblyDefs.AddRange(GetProjects().ToArray());
				foreach (IAssemblyDefinition def in assemblyDefs)
				{
					_Assemblies.AddRange(def.GetAssemblies());
				}
			}
			return _Assemblies;
		}

		internal List<XlMethodInfo> GetExcelMethods()
		{
			// HACK: Add an assembly resolve override, 
			// since the dynamic assembly otherwise cannot resolve this assembly !?
			// This is needed for the Custom Marshaler tagged into the dynamic assembly to work.
			AppDomain.CurrentDomain.AssemblyResolve += LocalResolve;

			List<MethodInfo> methods = new List<MethodInfo>();
			foreach (Assembly assembly in GetAssemblies()) 
			{
				methods.AddRange(AssemblyLoader.GetExcelMethods(assembly));
			}

			List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods);

			AppDomain.CurrentDomain.AssemblyResolve -= LocalResolve;

			return xlMethods;
		}

		internal List<AssemblyLoader.ExcelAddInInfo> GetExcelAddIns()
		{
			// Aggregate from assemblies
            List<AssemblyLoader.ExcelAddInInfo> addIns = new List<AssemblyLoader.ExcelAddInInfo>();
			foreach (Assembly assembly in GetAssemblies())
			{
				addIns.AddRange(AssemblyLoader.GetExcelAddIns(assembly));
			}

			return addIns;
		}
		
		// Statics
		private static DnaLibrary currentLibrary;
		internal static void Initialize()
		{
			// Load the current library
			// Get the .xll filename
			string xllFileLocation = Assembly.GetExecutingAssembly().Location;
			string xllDirectory =  Path.GetDirectoryName(xllFileLocation);
			string xllFileRoot = Path.GetFileNameWithoutExtension(xllFileLocation);
			string dnaFileName = Path.Combine(xllDirectory, xllFileRoot + ".dna");
			if (File.Exists(dnaFileName))
			{
				currentLibrary = LoadFrom(dnaFileName);
			}
			else
			{
				ErrorDisplay.DisplayErrorMessage(string.Format("The required .dna script file {0} does not exist.", dnaFileName));
			}
            // If there have been problems, ensure that there is at lease some current library.
            if (currentLibrary == null)
                currentLibrary = new DnaLibrary();
		}

		// HACK: See above - need this for resolving custom marshaler
 		private static Assembly LocalResolve(object sender, ResolveEventArgs args)
		{
			if (args.Name.StartsWith("ExcelDna,"))
				return Assembly.GetExecutingAssembly();
			else
				return null;
		}

		internal static DnaLibrary LoadFrom(string fileName)
		{
			DnaLibrary dnaLibrary;
//			XmlSerializer serializer = new XmlSerializer(typeof(DnaLibrary));
			XmlSerializer serializer = new Microsoft.Xml.Serialization.GeneratedAssembly.DnaLibrarySerializer();
			using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
			{
                try
                {
                    dnaLibrary = (DnaLibrary)serializer.Deserialize(fileStream);
                }
                catch (Exception e)
                {
                    string errorMessage = string.Format("There was an error while processing {0}:\r\n{1}\r\n{2}", fileName, e.Message,e.InnerException.Message);
                    ErrorDisplay.DisplayErrorMessage(errorMessage);
                    return null;
                }
			}

			dnaLibrary.Name = Path.GetFileNameWithoutExtension(fileName);
			dnaLibrary.FileName = fileName;
			return dnaLibrary;
		}

		internal static void Save(string fileName, DnaLibrary dnaProject)
		{
//			XmlSerializer serializer = new XmlSerializer(typeof(DnaLibrary));
			XmlSerializer serializer = new Microsoft.Xml.Serialization.GeneratedAssembly.DnaLibrarySerializer(); 
			using (FileStream fileStream = new FileStream(fileName, FileMode.Truncate))
			{
				serializer.Serialize(fileStream, dnaProject);
			}
		}

		public static DnaLibrary CurrentLibrary
		{
			get
			{
				return currentLibrary;
			}
		}

	}
}
