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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Schema;
using System.Drawing;
using System.Net;

using ExcelDna.Serialization;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Rtd;
using ExcelDna.ComInterop;

namespace ExcelDna.Integration
{
    [Serializable]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class DnaLibrary
    {
        private List<ExternalLibrary> _ExternalLibraries;
        [XmlElement("ExternalLibrary", typeof(ExternalLibrary))]
        public List<ExternalLibrary> ExternalLibraries
        {
            get { return _ExternalLibraries; }
            set { _ExternalLibraries = value; }
        }

        private List<Project> _Projects;
        [XmlElement("Project", typeof(Project))]
        public List<Project> Projects
        {
            get { return _Projects; }
            set { _Projects = value; }
        }

        private static string _XllPath = null;
        [XmlIgnore]
        internal static string XllPath
        {
            get
            {
                if (_XllPath == null)
                {
                    _XllPath = (string)XlCall.Excel(XlCall.xlGetName);
                }
                return _XllPath;
            }
            set
            {
                _XllPath = value;
            }
        }

        private string _Name;
        [XmlAttribute]
        public string Name
        {
            get
            {
                if (string.IsNullOrEmpty(_Name))
                {
                    _Name = Path.GetFileNameWithoutExtension(_XllPath);
                }
                return _Name;
            }
            set { _Name = value; }
        }

        private string _RuntimeVersion; // default is effectively v2.0.50727 ?
        [XmlAttribute]
        public string RuntimeVersion
        {
            get { return _RuntimeVersion; }
            set { _RuntimeVersion = value; }
        }

        private bool _ShadowCopyFiles = false;
        [XmlAttribute]
        public bool ShadowCopyFiles
        {
            get { return _ShadowCopyFiles; }
            set { _ShadowCopyFiles = value; }
        }

        // Used directly by the unmanaged loader.
        // If not ('true' or 'false') the sandbox is created only under runtime versions >= 4
        private string _CreateSandboxedAppDomain;
        [XmlAttribute]
        public string CreateSandboxedAppDomain
        {
            get { return _CreateSandboxedAppDomain; }
            set { _CreateSandboxedAppDomain = value; }
        }

        // Next bunch are abbreviations for Project contents
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

        private string _CompilerVersion;
        [XmlAttribute]
        public string CompilerVersion
        {
            get { return _CompilerVersion; }
            set { _CompilerVersion = value; }
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

        // No ExplicitExports flag on the DnaLibrary (for now), because it might cause confusion when mixed with ExternalLibraries.
        // Projects can be marked as ExplicitExports by adding an explicit <Project> tag.
        //private bool _ExplicitExports = false;
        //[XmlAttribute]
        //public bool ExplicitExports
        //{
        //    get { return _ExplicitExports; }
        //    set { _ExplicitExports = value; }
        //}

        private string _Code;
        [XmlText]
        public string Code
        {
            get { return _Code; }
            set { _Code = value; }
        }


        // DOCUMENT: The three different elements are their namespaces.
        //           The CustomUI and inner customUI are case sensitive.
        private List<XmlNode> _CustomUIs;
        [XmlElement("CustomUI", typeof(XmlNode))]
        public List<XmlNode> CustomUIs
        {
            get { return _CustomUIs; }
            set { _CustomUIs = value; }
        }

        private List<Image> _Images;
        [XmlElement("Image", typeof(Image))]
        public List<Image> Images
        {
            get { return _Images; }
            set { _Images = value; }
        }

        private string dnaResolveRoot;

        // Get projects explicit and implicitly present in the library
        public List<Project> GetProjects()
        {
            List<Project> projects = new List<Project>();
            if (Projects != null)
                projects.AddRange(Projects);
            if (_Code != null && Code.Trim() != "")
                projects.Add(new Project(Language, CompilerVersion, References, _Code, DefaultReferences, DefaultImports, false));
            return projects;
        }

        internal List<ExportedAssembly> GetAssemblies(string pathResolveRoot)
        {
            List<ExportedAssembly> assemblies = new List<ExportedAssembly>();
            try
            {
                if (ExternalLibraries != null)
                {
                    foreach (ExternalLibrary lib in ExternalLibraries)
                    {
                        assemblies.AddRange(lib.GetAssemblies(pathResolveRoot, this));
                    }
                }
                foreach (Project proj in GetProjects())
                {
                    assemblies.AddRange(proj.GetAssemblies(pathResolveRoot, this));
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error while loading assemblies. Exception: " + e.Message);
            }
            return assemblies;
        }

        // Managed AddIns related to this DnaLibrary
        // Kept so that AutoClose can be called later.
        [XmlIgnore]
        private List<AssemblyLoader.ExcelAddInInfo> _addIns = new List<AssemblyLoader.ExcelAddInInfo>();
        [XmlIgnore]
        private List<MethodInfo> _methods = new List<MethodInfo>();

        // The idea is that initialize compiles, loads and sorts out the assemblies,
        //    but does not depend on any calls to Excel.
        // Then Initialize can be called during RTD registration or loading, 
        //    without 'AutoOpening' the add-in
        internal void Initialize()
        {
            // Get MethodsInfos and AddIn classes from assemblies
            List<Type> rtdServerTypes = new List<Type>();
            List<ExcelComClassType> comClassTypes = new List<ExcelComClassType>();

            // Recursively get assemblies down .dna tree.
            List<ExportedAssembly> assemblies = GetAssemblies(dnaResolveRoot);
            AssemblyLoader.ProcessAssemblies(assemblies, _methods, _addIns, rtdServerTypes, comClassTypes);

            // Register RTD Server Types immediately
            RtdRegistration.RegisterRtdServerTypes(rtdServerTypes);

            // CAREFUL: This interacts with the implementation of ExcelRtdServer to implement the thread-safe synchronization.
            // Check whether we have an ExcelRtdServer type, and need to install the Sync Window
            // Uninstalled in the AutoClose
            bool registerSyncManager = false;
            foreach (Type rtdType in rtdServerTypes)
            {
                if (rtdType.IsSubclassOf(typeof(ExcelRtdServer)))
                {
                    registerSyncManager = true;
                    break;
                }
            }
            if (registerSyncManager) SynchronizationManager.Install();

            // Register COM Server Types immediately
            ComServer.RegisterComClassTypes(comClassTypes);
        }

        internal void AutoOpen()
        {
            // Register my Methods
            ExcelIntegration.RegisterMethods(_methods);

            // Invoke AutoOpen in all assemblies
            foreach (AssemblyLoader.ExcelAddInInfo addIn in _addIns)
            {
                try
                {
                    if (addIn.AutoOpenMethod != null)
                    {
                        addIn.AutoOpenMethod.Invoke(addIn.Instance, null);
                    }
                }
                catch (Exception e)
                {
                    // TODO: What to do here?
                    Debug.Print(e.Message);
                }
            }

            // Load my UI
            LoadCustomUI();
        }

        internal void AutoClose()
        {
            UnloadCustomUI();

            foreach (AssemblyLoader.ExcelAddInInfo addIn in _addIns)
            {
                try
                {
                    if (addIn.AutoCloseMethod != null)
                    {
                        addIn.AutoCloseMethod.Invoke(addIn.Instance, null);
                    }
                }
                catch (Exception e)
                {
                    // TODO: What to do here?
                    Debug.WriteLine(e.Message);
                }
            }
            // This is safe, even if never registered
            SynchronizationManager.Uninstall();
            _addIns.Clear();
        }

        internal void LoadCustomUI()
        {
            bool uiLoaded = false;
            if (ExcelDnaUtil.ExcelVersion >= 12.0)
            {
                // Load ComAddIns
                foreach (AssemblyLoader.ExcelAddInInfo addIn in _addIns)
                {
                    if (addIn.IsCustomUI)
                    {
                        // Load ExcelRibbon classes
                        ExcelRibbon excelRibbon = addIn.Instance as ExcelRibbon;
                        excelRibbon.DnaLibrary = addIn.ParentDnaLibrary;
                        ExcelComAddInHelper.LoadComAddIn(excelRibbon);
                        uiLoaded = true;
                    }
                }

                // CONSIDER: Really not sure if this is a good idea - seems to interfere with unloading somehow.
                //if (uiLoaded == false && CustomUIs != null)
                //{
                //    // Check whether we should add an empty ExcelCustomUI instance to load a Ribbon interface?
                //    bool loadEmptyAddIn = false;
                //    if (CustomUIs != null)
                //    {
                //        foreach (XmlNode xmlCustomUI in CustomUIs)
                //        {
                //            if (xmlCustomUI.LocalName == "customUI" &&
                //                (xmlCustomUI.NamespaceURI == ExcelRibbon.NamespaceCustomUI2007 ||
                //                 (ExcelDnaUtil.ExcelVersion >= 14.0 &&
                //                  xmlCustomUI.NamespaceURI == ExcelRibbon.NamespaceCustomUI2010)))
                //            {
                //                loadEmptyAddIn = true;
                //            }
                //            if (loadEmptyAddIn)
                //            {
                //                // There will be Ribbon xml to load. Make a temp add-in and load it.
                //                ExcelRibbon customUI = new ExcelRibbon();
                //                customUI.DnaLibrary = this;
                //                ExcelComAddInHelper.LoadComAddIn(customUI);
                //                uiLoaded = true;
                //            }
                //        }
                //    }
                //}
            }

            // should we load CommandBars?
            if (uiLoaded == false && CustomUIs != null)
            {
                foreach (XmlNode xmlCustomUI in CustomUIs)
                {
                    if (xmlCustomUI.LocalName == "commandBars")
                    {
                        ExcelCommandBarUtil.LoadCommandBars(xmlCustomUI, this.GetImage);
                    }
                }
            }
        }

        internal void UnloadCustomUI()
        {
            // This is safe, even if no Com Add-Ins were loaded.
            ExcelComAddInHelper.UnloadComAddIns();
            ExcelCommandBarUtil.UnloadCommandBars();
        }

        // Statics
        private static DnaLibrary rootLibrary;
        internal static void InitializeRootLibrary(string xllPath)
        {
            // Loads the primary .dna library
            // Load sequence is:
            // 1. Look for a packed .dna file named "__MAIN__" in the .xll.
            // 2. Look for the .dna file in the same directory as the .xll file, with the same name and extension .dna.

            Debug.WriteLine("Enter DnaLibrary.InitializeRootLibrary");
            XllPath = xllPath;
            Logging.LogDisplay.CreateInstance();
            byte[] dnaBytes = ExcelIntegration.GetDnaFileBytes("__MAIN__");
            if (dnaBytes != null)
            {
                Debug.WriteLine("Got Dna file from resources.");
                string pathResolveRoot = Path.GetDirectoryName(DnaLibrary.XllPath);
                rootLibrary = LoadFrom(dnaBytes, pathResolveRoot);
                // ... would have displayed error and returned null if there was an error.
            }
            else
            {
                Debug.WriteLine("No Dna file in resources - looking for file.");
                // No packed .dna file found - load from a .dna file.
                string dnaFileName = Path.ChangeExtension(XllPath, ".dna");
                rootLibrary = LoadFrom(dnaFileName);
                // ... would have displayed error and returned null if there was an error.
            }

            // If there have been problems, ensure that there is at lease some current library.
            if (rootLibrary == null)
            {
                Debug.WriteLine("No Dna Library found.");
                rootLibrary = new DnaLibrary();
            }

            rootLibrary.Initialize();
            Debug.WriteLine("Exit DnaLibrary.Initialize");
        }

        internal static void DeInitialize()
        {
            // Called to shut down the Add-In.
            // Free whatever possible
            rootLibrary = null;
        }

        public static DnaLibrary LoadFrom(Uri uri)
        {
            DnaLibrary dnaLibrary;
            XmlSerializer serializer = new DnaLibrarySerializer();

            // The uri might be file or http.
            try
            {
                WebRequest req = WebRequest.CreateDefault(uri);
                WebResponse res = req.GetResponse();
                Stream s = res.GetResponseStream();
                dnaLibrary = (DnaLibrary)serializer.Deserialize(s);
            }
            catch (Exception e)
            {
                string errorMessage = string.Format("There was an error while loading the .dna file from a Uri:\r\n{0}\r\n{1}\r\nUri:{2}", e.Message, e.InnerException != null ? e.InnerException.Message : string.Empty, uri.ToString());
                Debug.WriteLine(errorMessage);
                ExcelDna.Logging.LogDisplay.WriteLine(errorMessage);
                return null;
            }

            dnaLibrary.dnaResolveRoot = null;
            return dnaLibrary;

        }

        public static DnaLibrary LoadFrom(byte[] bytes, string pathResolveRoot)
        {
            DnaLibrary dnaLibrary;
            XmlSerializer serializer = new DnaLibrarySerializer();

            try
            {
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    dnaLibrary = (DnaLibrary)serializer.Deserialize(ms);
                }
            }
            catch (Exception e)
            {
                string errorMessage = string.Format("There was an error while processing DnaFile bytes:\r\n{0}\r\n{1}", e.Message, e.InnerException != null ? e.InnerException.Message : string.Empty);
                Debug.WriteLine(errorMessage);
                //ExcelDna.Logging.LogDisplay.SetText(errorMessage);
                return null;
            }
            dnaLibrary.dnaResolveRoot = pathResolveRoot;
            return dnaLibrary;
        }

        public static DnaLibrary LoadFrom(string fileName)
        {
            DnaLibrary dnaLibrary;

            if (!File.Exists(fileName))
            {
                ExcelDna.Logging.LogDisplay.WriteLine("The required .dna script file {0} does not exist.", fileName);
                return null;
            }

            try
            {
                XmlSerializer serializer = new DnaLibrarySerializer();
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    dnaLibrary = (DnaLibrary)serializer.Deserialize(fileStream);
                }
            }
            catch (Exception e)
            {
                ExcelDna.Logging.LogDisplay.WriteLine("There was an error during processing of {0}:\r\n{1}\r\n{2}", fileName, e.Message, e.InnerException != null ? e.InnerException.Message : string.Empty);
                return null;
            }
            dnaLibrary.dnaResolveRoot = Path.GetDirectoryName(fileName);
            return dnaLibrary;
        }

        public static void Save(string fileName, DnaLibrary dnaLibrary)
        {
            //			XmlSerializer serializer = new XmlSerializer(typeof(DnaLibrary));
            XmlSerializer serializer = new DnaLibrarySerializer();
            using (FileStream fileStream = new FileStream(fileName, FileMode.Truncate))
            {
                serializer.Serialize(fileStream, dnaLibrary);
            }
        }

        public static byte[] Save(DnaLibrary dnaLibrary)
        {
            //			XmlSerializer serializer = new XmlSerializer(typeof(DnaLibrary));
            XmlSerializer serializer = new DnaLibrarySerializer();
            using (MemoryStream ms = new MemoryStream())
            {
                serializer.Serialize(ms, dnaLibrary);
                return ms.ToArray();
            }
        }

        public static DnaLibrary CurrentLibrary
        {
            // Should not be called before Initialize
            get
            {
                if (rootLibrary == null)
                {
                    throw new InvalidOperationException("No CurrentLibrary set.");
                }
                return rootLibrary;
            }
        }

        // Called during initialize when displaying log
        // TODO: Clean up - inserted to prevent recursion when displaying log during initialize
        internal static string CurrentLibraryName
        {
            get
            {
                if (rootLibrary == null)
                {
                    string dllName = XllPath;
                    return Path.GetFileNameWithoutExtension(dllName);
                }
                return CurrentLibrary.Name;
            }
        }

        internal static string ExecutingDirectory
        {
            get
            {
                return Path.GetDirectoryName(XllPath);
            }
        }

        public string ResolvePath(string path)
        {
            return ResolvePath(path, dnaResolveRoot);
        }

        // ResolvePath tries to figure out the actual path to a file - either a .dna file or an 
        // assembly to be packed.
        // Resolution sequence:
        // 1. Check the path - if not rooted that will be relative to working directory.
        // 2. If the path is rooted, try the filename part relative to the .dna file location.
        //    If the path is not rooted, try the whole path relative to the .dna file location.
        // 3. Try step 2 against the appdomain.
        // dnaDirectory can be null, in which case we don't check against it.
        // This might be the case when we're loading from a Uri.
        // CONSIDER: Should the Uri case be handled better?
        public static string ResolvePath(string path, string dnaDirectory)
        {

            Debug.Print("ResolvePath: Resolving {0} from DnaDirectory: {1}", path, dnaDirectory);
            if (File.Exists(path))
            {
                string fullPath = Path.GetFullPath(path);
                Debug.Print("ResolvePath: Found at {0}", fullPath);
                return fullPath;
            }

            string fileName = Path.GetFileName(path);
            if (dnaDirectory != null)
            {
                // Try relative to dna directory
                string dnaPath;
                if (Path.IsPathRooted(path))
                {
                    // It was a rooted path -- try locally instead
                    dnaPath = Path.Combine(dnaDirectory, fileName);
                }
                else
                {
                    // Not rooted - try a path relative to local directory
                    dnaPath = System.IO.Path.Combine(dnaDirectory, path);
                }
                Debug.Print("ResolvePath: Checking at {0}", dnaPath);
                if (File.Exists(dnaPath))
                {
                    Debug.Print("ResolvePath: Found at {0}", dnaPath);
                    return dnaPath;
                }
            }
            // try relative to AppDomain BaseDirectory
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (baseDirectory != dnaDirectory)
            {
                string basePath;
                if (Path.IsPathRooted(path))
                {
                    basePath = Path.Combine(baseDirectory, fileName);
                }
                else
                {
                    basePath = System.IO.Path.Combine(baseDirectory, path);
                }
                // ... and check again
                Debug.Print("ResolvePath: Checking at {0}", basePath);
                if (File.Exists(basePath))
                {
                    Debug.Print("ResolvePath: Found at {0}", basePath);
                    return basePath;
                }
            }

            // CONSIDER: Do we really need this? 
            // - it is here mainly for backward compatibility, so that Path="System.Windows.Forms.dll" will still work.
            // Try in .NET framework directory
            // Last try - check in current version of .NET's directory, 
            string frameworkBase = RuntimeEnvironment.GetRuntimeDirectory();
            string frameworkPath = Path.Combine(frameworkBase, fileName);
            if (File.Exists(frameworkPath))
            {
                Debug.Print("ResolvePath: Found at {0}", frameworkPath);
                return frameworkPath;
            }

            // Else give up (maybe try load from GAC for assemblies?)
            Debug.Print("ResolvePath: Could not find {0} from DnaDirectory {1}", path, dnaDirectory);
            return null;
        }

        public Bitmap GetImage(string imageId)
        {
            // We expect these to be small images.

            // First check if imageId is in the DnaLibrary's Image list.
            // DOCUMENT: Case sensitive match.
            foreach (Image image in Images)
            {
                if (image.Name == imageId && image.Path != null)
                {
                    byte[] imageBytes = null;
                    System.Drawing.Image imageLoaded;
                    if (image.Path.StartsWith("packed:"))
                    {
                        string resourceName = image.Path.Substring(7);
                        imageBytes = ExcelIntegration.GetImageBytes(resourceName);
                    }
                    else
                    {
                        string imagePath = ResolvePath(image.Path);
                        if (imagePath == null)
                        {
                            // This is the image but we could not find it !?
                            Debug.Print("DnaLibrary.LoadImage - For image {0} the path resolution failed: {1}", image.Name, image.Path);
                            return null;
                        }
                        imageBytes = File.ReadAllBytes(imagePath);
                    }
                    using (MemoryStream ms = new MemoryStream(imageBytes, false))
                    {
                        imageLoaded = Bitmap.FromStream(ms);
                        if (imageLoaded is Bitmap)
                        {
                            return (Bitmap)imageLoaded;
                        }
                        Debug.Print("Image {0} read from {1} was not a bitmap!?", image.Name, image.Path);
                    }
                }
            }
            return null;
        }

    }

    //public class CustomUI : IXmlSerializable
    //{
    //    public string Content { get; set; }
    //    public string NsUri { get; set; }


    //    public XmlSchema GetSchema()
    //    {
    //        return null;
    //    }

    //    public void WriteXml(XmlWriter w)
    //    {
    //        w.WriteStartElement("customUI");
    //        w.WriteAttributeString("xmlns", NsUri);
    //        w.WriteRaw(Content);
    //        w.WriteEndElement();
    //    }

    //    public void ReadXml(XmlReader r)
    //    {
    //        NsUri = r.NamespaceURI;
    //        Content = r.ReadOuterXml();
    //    }
    //
    //}
}


/*
public class CDataField : IXmlSerializable
    {
        private string elementName;
        private string elementValue;

        public CDataField(string elementName, string elementValue)
        {
            this.elementName = elementName;
            this.elementValue = elementValue;
        }

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void WriteXml(XmlWriter w)
        {
            w.WriteStartElement(this.elementName);
            w.WriteCData(this.elementValue);
            w.WriteEndElement();
        }

        public void ReadXml(XmlReader r)
        {                      
            throw new NotImplementedException("This method has not been implemented");
        }
    }
 */

// Not sure if this is still valid - current version is not customised....?????
// CAREFUL: Customization of the deserialization in this if clause....
//else if (((object)Reader.LocalName == (object)id11_customUI /*&& (object)Reader.NamespaceURI == (object)id2_Item */))
//{
//    a_9.Add((global::System.Xml.XmlNode)ReadXmlNode(false /*true*/));
//}

