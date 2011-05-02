using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Reflection;
using ExcelDna.ComInterop.ComRegistration;
using ExcelDna.Integration.Extensibility;

using HRESULT = System.Int32;
using IID = System.Guid;
using CLSID = System.Guid;
using DWORD = System.UInt32;
using System.Globalization;

namespace ExcelDna.Integration.CustomUI
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ExcelComAddIn : IDTExtensibility2 
    {
        internal DnaLibrary DnaLibrary { get; set; }
        private string _progId;
        protected string ProgId
        {
            get { return _progId; }
        }

        internal void SetProgId(string progId)
        {
            _progId = progId;
        }

        #region IDTExtensibility2 interface - not used by ExcelDna
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnConnection");
    
            
            // Grab an Application reference here and keep around...?
            // TODO: Check that Excel shuts down in various settings.
            ExcelDnaUtil.Application = Application;
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnBeginShutdown");
        }
        #endregion
    }

    [ComVisible(true)]
    //    [Guid("4B0C96EC-9740-4F78-8396-84B0FABB0E74")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ExcelRibbon : ExcelComAddIn, IRibbonExtensibility //, ICustomTaskPaneConsumer
    {
        public const string NamespaceCustomUI2010 = @"http://schemas.microsoft.com/office/2009/07/customui";
        public const string NamespaceCustomUI2007 = @"http://schemas.microsoft.com/office/2006/01/customui";

        public virtual string GetCustomUI(string RibbonID)
        {
            if (RibbonID != "Microsoft.Excel.Workbook")
            {
                Debug.Print("ExcelRibbon.GetCustomUI - Invalid RibbonID for Excel. RibbonID: {0}", RibbonID);
                return null;
            }

            // Default behaviour for GetCustomUI is to look in DnaLibrary.
            // We return a CustomUI based on the version of Excel, as follows:
            // If Excel12, look for a CustomUI with namespace... If not found, return nothing
            // If Excel14 or bigger look for CustomUI with namespace ... If not found, look for ... else return nothing.
            // (not sure how to future-proof...)

            Dictionary<string, string> customUIs = new Dictionary<string, string>();
            foreach (XmlNode customUI in this.DnaLibrary.CustomUIs)
            {
                customUIs[customUI.NamespaceURI] = customUI.OuterXml;
            }

            if (ExcelDnaUtil.ExcelVersion >= 14.0)
            {
                if (customUIs.ContainsKey(NamespaceCustomUI2010))
                {
                    return customUIs[NamespaceCustomUI2010];
                }
                if (customUIs.ContainsKey(NamespaceCustomUI2007))
                {
                    return customUIs[NamespaceCustomUI2007];
                }
                return null;
            }
            if (ExcelDnaUtil.ExcelVersion >= 12.0)
            {
                if (customUIs.ContainsKey(NamespaceCustomUI2007))
                {
                    return customUIs[NamespaceCustomUI2007];
                }
                return null;
            }
            throw new InvalidOperationException("Not expected to provide CustomUI string for Excel version < 12.0");
        }

        // LoadImage helper - to use need to mark loadImage='LoadImage' in the xml.
        // 1. An IPictureDisp
        // 2. A System.Drawing.Bitmap
        // 3. A string containing an imageMso identifier
        // Our default implementation ...
        public virtual object LoadImage(string imageId)
        {
            // Default implementation ...
            return DnaLibrary.GetImage(imageId);
        }

        // RunTagMacro helper function
        public virtual void RunTagMacro(IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(control.Tag))
            {
                // CONSIDER: Is this a danger for shutting down - surely not...?
                object app = ExcelDnaUtil.Application;
                app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, new object[] { control.Tag }, new System.Globalization.CultureInfo(1033));
            }
        }
    }

    public static class ExcelComAddInHelper
    {
        // Com Add-ins loaded for Ribbons.
        static List<object> loadedComAddIns = new List<object>();

        public static void LoadComAddIn(ExcelComAddIn addIn)
        {
            // We pick a new Guid as ClassId for this add-in.
            CLSID clsId = Guid.NewGuid();
            // and make the ProgId from this Guid - max 39 chars....
            string progId = "AddIn." + clsId.ToString("N");
            addIn.SetProgId(progId);
            // Put together some nicer descriptions for the Add-ins dialog.
            string friendlyName;
            if (addIn is ExcelRibbon)
                friendlyName = addIn.DnaLibrary.Name + " (Ribbon Helper)";
            else if (addIn is ExcelCustomTaskPaneAddIn)
                friendlyName = addIn.DnaLibrary.Name + " (Custom Task Pane Helper)";
            else 
                friendlyName = addIn.DnaLibrary.Name + " (COM Add-in Helper)";
            string description = string.Format("Dynamically created COM Add-in to load custom UI for the Excel Add-in {0}, located at {1}.", addIn.DnaLibrary.Name, DnaLibrary.XllPath);

            try
            {
                Debug.Print("Getting Application object");
                object app = ExcelDnaUtil.Application;
                Type appType = app.GetType();
                Debug.Print("Got Application object: " + app.GetType().ToString());

                CultureInfo ci = new CultureInfo(1033);

                object excelComAddIns;
                object comAddIn;
                using (SingletonClassFactoryRegistration regClassFactory = new SingletonClassFactoryRegistration(addIn, clsId))
                using (ProgIdRegistration regProgId = new ProgIdRegistration(progId, clsId))
                using (ComAddInRegistration regComAddIn = new ComAddInRegistration(progId, friendlyName, description))
                {
                    excelComAddIns = appType.InvokeMember("COMAddIns", BindingFlags.GetProperty, null, app, null, ci);
//                            Debug.Print("Got COMAddins object: " + excelComAddIns.GetType().ToString());
                    appType.InvokeMember("Update", BindingFlags.InvokeMethod, null, excelComAddIns, null, ci);
//                            Debug.Print("Updated COMAddins object with AddIn registered");
                    comAddIn = excelComAddIns.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, excelComAddIns, new object[] { progId }, ci);
//                            Debug.Print("Got the COMAddin object: " + comAddIn.GetType().ToString());

                    // At this point Excel knows how to load our add-in by CLSID, so we could clean up the 
                    // registry aggressively, before the actual (dangerous?) loading starts.
                    // But this seems to lead to some distress - Excel has some assertion checked when 
                    // it updates the LoadBehavior after a successful load....
                    comAddIn.GetType().InvokeMember("Connect", BindingFlags.SetProperty, null, comAddIn, new object[] { true }, ci);
//                            Debug.Print("COMAddin is loaded.");
                    loadedComAddIns.Add(comAddIn);
                }
            }
            catch (Exception e)
            {
                Debug.Print("LoadComAddIn exception: " + e.ToString());
                // CONSIDER: How to deal with errors here...? For now I just re-raise the exception.
                throw;
            }
        }

        internal static void UnloadComAddIns()
        {
            CultureInfo ci = new CultureInfo(1033);
            foreach (object comAddIn in loadedComAddIns)
            {
                comAddIn.GetType().InvokeMember("Connect", System.Reflection.BindingFlags.SetProperty, null, comAddIn, new object[] { false }, ci);
                Debug.Print("COMAddin is unloaded.");
            }
        }
    }

}
