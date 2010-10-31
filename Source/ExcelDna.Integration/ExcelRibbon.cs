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

namespace ExcelDna.Integration.CustomUI
{
    [ComVisible(true)]
    //    [Guid("4B0C96EC-9740-4F78-8396-84B0FABB0E74")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ExcelRibbon : IDTExtensibility2, IRibbonExtensibility //, ICustomTaskPaneConsumer
    {
        public const string NamespaceCustomUI2010 = @"http://schemas.microsoft.com/office/2009/07/customui";
        public const string NamespaceCustomUI2007 = @"http://schemas.microsoft.com/office/2006/01/customui";

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
            Debug.Print("ExcelRibbon.OnConnection");
            // TODO: Grab an Application reference here and keep around...?
            //       (check that Excel shuts down if we do).
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Debug.Print("ExcelRibbon.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Debug.Print("ExcelRibbon.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Debug.Print("ExcelRibbon.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Debug.Print("ExcelRibbon.OnBeginShutdown");
        }
        #endregion

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
            foreach (XmlNode customUI in DnaLibrary.CustomUIs)
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
                object app = ExcelDnaUtil.Application;
                app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, new object[] { control.Tag });
                //object result;
                //XlCall.TryExcel(XlCall.xlcRun, out result, control.Tag);
            }
        }

        // For CustomTaskPane stuff
        // TODO: ActiveX -> WPF without Windows.Forms....?
        //public virtual void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.IUnknown)] object /*ICTPFactory*/ CTPFactoryInst)
        //{
        //    // throw new NotImplementedException();
        //}
    }

    internal static class ExcelComAddIn
    {
        // Com Add-ins loaded for Ribbons.
        static List<object> loadedComAddIns = new List<object>();

        public static void LoadComAddIn(ExcelRibbon addIn)
        {
            // We pick a new Guid as ClassId for this add-in.
            CLSID clsId = Guid.NewGuid();
            // and make the ProgId from this Guid - max 39 chars....
            string progId = "Ribbon." + clsId.ToString("N");
            addIn.SetProgId(progId);
            // Put together some nicer descriptions for the Add-ins dialog.
            string friendlyName = addIn.DnaLibrary.Name + " (Ribbon Helper)";
            string description = string.Format("Dynamically created COM Add-in to load Ribbon interface for the Excel Add-in {0}, located at {1}.", addIn.DnaLibrary.Name, DnaLibrary.XllPath);

            try
            {
                Debug.Print("Getting Application object");
                object app = ExcelDnaUtil.Application;
                Debug.Print("Got Application object: " + app.GetType().ToString());

                object excelComAddIns;
                object comAddIn;
                using (SingletonClassFactoryRegistration regClassFactory = new SingletonClassFactoryRegistration(clsId, addIn))
                {
                    using (ProgIdRegistration regProgId = new ProgIdRegistration(progId, clsId))
                    {
                        using (ComAddInRegistration regComAddIn = new ComAddInRegistration(progId, friendlyName, description))
                        {
                            excelComAddIns = app.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, app, null);
//                            Debug.Print("Got COMAddins object: " + excelComAddIns.GetType().ToString());
                            app.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, excelComAddIns, null);
//                            Debug.Print("Updated COMAddins object with AddIn registered");
                            comAddIn = excelComAddIns.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, excelComAddIns, new object[] { progId });
//                            Debug.Print("Got the COMAddin object: " + comAddIn.GetType().ToString());

                            // At this point Excel knows how to load our add-in by CLSID, so we could clean up the 
                            // registry aggressively, before the actual (dangerous?) loading starts.
                            // But this seems to lead to some distress - Excel has some assertion checked when 
                            // it updates the LoadBehavior after a successful load....
                            comAddIn.GetType().InvokeMember("Connect", BindingFlags.SetProperty, null, comAddIn, new object[] { true });
//                            Debug.Print("COMAddin is loaded.");
                            loadedComAddIns.Add(comAddIn);
                        }
                    }

                    // Here we have already removed the Excel AddIn entry, and the ProgId.
                    // Loading the add-in still works fine, but Excel seems to get a bit distressed, as it wants to write 
                    // a LoadBehavior entry to the registry, but will fail.

                    //comAddIn.GetType().InvokeMember("Connect", BindingFlags.SetProperty, null, comAddIn, new object[] { true });
                    //Debug.Print("COMAddin is loaded.");
                    //loadedComAddIns.Add(comAddIn);
                }

                // app.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, excelComAddIns, null);
                // Debug.Print("Updated COMAddins object with AddIn loaded and de-registered.");
            }
            catch (Exception e)
            {
                Debug.Print("LoadComAddIn exception: " + e.ToString());
            }
        }

        public static void UnloadComAddIns()
        {
            foreach (object comAddIn in loadedComAddIns)
            {
                comAddIn.GetType().InvokeMember("Connect", System.Reflection.BindingFlags.SetProperty, null, comAddIn, new object[] { false });
                Debug.Print("COMAddin is unloaded.");
            }
        }
    }
}
