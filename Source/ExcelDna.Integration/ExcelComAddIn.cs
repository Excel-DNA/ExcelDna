using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop.ComRegistration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;

namespace ExcelDna.Integration
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

        #region IDTExtensibility2 interface
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Debug.Print("ExcelComAddIn.OnConnection");
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

    public static class ExcelComAddInHelper
    {
        // Com Add-ins loaded for Ribbons.
        static List<object> loadedComAddIns = new List<object>();

        public static void LoadComAddIn(ExcelComAddIn addIn)
        {
            // If we are called without the addIn's DnaLibrary being set, default to the current library
            if (addIn.DnaLibrary == null)
                addIn.DnaLibrary = DnaLibrary.CurrentLibrary;

            Guid clsId;
            string progId;

            // If we have both an explicit Guid and an explicit ProgId, then use those, else create synthetic ProgId.
            object[] progIdAttrs = addIn.GetType().GetCustomAttributes(typeof(ProgIdAttribute), false);
            object[] guidAttrs = addIn.GetType().GetCustomAttributes(typeof(GuidAttribute), false);
            if (progIdAttrs.Length >= 1 && guidAttrs.Length >= 1)
            {
                // Use the attributes
                ProgIdAttribute progIdAtt = (ProgIdAttribute)progIdAttrs[0];
                progId = progIdAtt.Value;

                GuidAttribute guidAtt = (GuidAttribute)guidAttrs[0];
                clsId = new Guid(guidAtt.Value);
            }
            else
            {
                // Use a stable Guid derived from the Xll Path (since Excel stores load-times and other info for every COM add-in loaded in the registry)
                clsId = ExcelDnaUtil.XllGuid;
                // and make the ProgId from this Guid - max 39 chars....
                progId = "Dna." + clsId.ToString("N") + "." + loadedComAddIns.Count;
            }
            addIn.SetProgId(progId);

            // Put together some nicer descriptions for the Add-ins dialog.
            string friendlyName;
            if (addIn is ExcelRibbon)
                friendlyName = addIn.DnaLibrary.Name; // + " (Ribbon Helper)"; (No more - it is displayed in the Ribbon tooltip!)
            else if (addIn is ExcelCustomTaskPaneAddIn)
                friendlyName = addIn.DnaLibrary.Name + " (Custom Task Pane Helper)";
            else
                friendlyName = addIn.DnaLibrary.Name + " (COM Add-in Helper)";
            string description = string.Format("Dynamically created COM Add-in to load custom UI for the Excel Add-in {0}, located at {1}.", addIn.DnaLibrary.Name, DnaLibrary.XllPath);


            Debug.Print("Getting Application object");
            object app = ExcelDnaUtil.Application;
            Type appType = app.GetType();
            Debug.Print("Got Application object: " + app.GetType().ToString());

            CultureInfo ci = new CultureInfo(1033);
            object excelComAddIns;
            object comAddIn;

            try
            {
                using (new SingletonClassFactoryRegistration(addIn, clsId))
                using (new ProgIdRegistration(progId, clsId))
                using (new ComAddInRegistration(progId, friendlyName, description))
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
            catch (UnauthorizedAccessException secex)
            {
                Logging.LogDisplay.WriteLine("The Ribbon/COM Add-in helper required by add-in {0} could not be registered.\r\nThis may be due to restricted permissions on the user's HKCU\\Software\\Classes key.\r\nError message: {1}", DnaLibrary.CurrentLibrary.Name, secex.Message);
                Debug.Print("LoadComAddIn exception: " + secex.ToString());
            }
            catch (Exception ex)
            {
                Logging.LogDisplay.WriteLine("The Ribbon/COM Add-in helper required by add-in {0} could not be registered.\r\nThis is an unexpected error.\r\nError message: {1}", DnaLibrary.CurrentLibrary.Name, ex.Message);
                Debug.Print("LoadComAddIn exception: " + ex.ToString());
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
