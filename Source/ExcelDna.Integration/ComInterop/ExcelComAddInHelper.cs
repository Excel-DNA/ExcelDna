//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop.ComRegistration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
    public static class ExcelComAddInHelper
    {
        // Com Add-ins loaded for Ribbons.
        static List<object> loadedComAddIns = new List<object>();

        public static void OnUnloadComAddIn(ExcelComAddIn addIn, object addInInst)
        {
            loadedComAddIns.Remove(addInInst);
        }

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
                // Change from Dna.xxx.n to Dna_xxx_n to avoid McAfee bug that blocks registry writes with a "." anywhere
                progId = "Dna_" + clsId.ToString("N") + "_" + loadedComAddIns.Count;
            }
            addIn.SetProgId(progId);

            // Put together some nicer descriptions for the Add-ins dialog.
            string friendlyName;
            if (addIn.FriendlyName != null)
                friendlyName = addIn.FriendlyName;
            else if (addIn is ExcelRibbon
#if COM_GENERATED
                || addIn is ComInterop.Generator.ExcelRibbon
#endif
                )
                friendlyName = addIn.DnaLibrary.Name; // + " (Ribbon Helper)"; (No more - it is displayed in the Ribbon tooltip!)
            else if (addIn is ExcelCustomTaskPaneAddIn)
                friendlyName = addIn.DnaLibrary.Name + " (Custom Task Pane Helper)";
            else
                friendlyName = addIn.DnaLibrary.Name + " (COM Add-in Helper)";
            string description = addIn.Description ?? string.Format("Dynamically created COM Add-in to load custom UI for the Excel Add-in {0}, located at {1}.", addIn.DnaLibrary.Name, DnaLibrary.XllPath);


            Logger.ComAddIn.Verbose("Getting Application object");
            object app = ExcelDnaUtil.ApplicationObject;
            ComInterop.IType typeAdapter = ComInterop.Util.TypeAdapter;
            Logger.ComAddIn.Verbose("Got Application object: " + app.GetType().ToString());

            object excelComAddIns;
            object comAddIn;

            try
            {
                Logger.ComAddIn.Verbose("Loading Ribbon/COM Add-In {0} ({1}) {2} / {3}",
                    addIn.GetType().FullName, friendlyName, progId, clsId);

                using (new ProgIdRegistration(progId, clsId))
                using (new ClsIdRegistration(clsId, progId))
                using (new ComAddInRegistration(progId, friendlyName, description))
                using (new AutomationSecurityOverride(app))
                {
                    // NOTE: A bug was introduced in Excel around version 2310 (16.0.16921.20000) that broke some of the COM add-in load scenarios for ribbons and CTP
                    // This point in the code will (at least under these versions) load the add-in via COM DllGetClassObject
                    // But the add-in is not 'fully' loaded, e.g. the Ribbon is not loaded
                    // However, events on the add-in object will fire, specifically the OnDisconnect event when we do Connect = false later.
                    // So to avoid inconvenience to our 'real' add-in, we give Excel a dummy add-in for now
                    var dummyAddIn = new ComInterop.DummyComAddIn();
                    using (new SingletonClassFactoryRegistration(dummyAddIn, clsId))
                    {
                        excelComAddIns = typeAdapter.GetProperty("COMAddIns", app);

                        //                            Debug.Print("Got COMAddins object: " + excelComAddIns.GetType().ToString());
                        typeAdapter.Invoke("Update", null, excelComAddIns);
                        //                            Debug.Print("Updated COMAddins object with AddIn registered");
                        comAddIn = typeAdapter.Invoke("Item", new object[] { progId }, excelComAddIns);
                        //                            Debug.Print("Got the COMAddin object: " + comAddIn.GetType().ToString());

                        // At this point Excel knows how to load our add-in by CLSID, so we could clean up the 
                        // registry aggressively, before the actual (dangerous?) loading starts.
                        // But this seems to lead to some distress - Excel has some assertion checked when 
                        // it updates the LoadBehavior after a successful load....

                        object connectState = typeAdapter.GetProperty("Connect", comAddIn);
                        typeAdapter.SetProperty("Connect", false, comAddIn);
                    }
                    // Swap out the dummy add-in for the real one
                    using (new SingletonClassFactoryRegistration(addIn, clsId))
                    {
                        typeAdapter.SetProperty("Connect", true, comAddIn);
                    }
                    //                            Debug.Print("COMAddin is loaded.");
                    loadedComAddIns.Add(comAddIn);

                    Logger.ComAddIn.Verbose("Completed Loading Ribbon/COM Add-In");

                }
            }
            catch (UnauthorizedAccessException secex)
            {
                Logger.ComAddIn.Error(secex, "The Ribbon/COM Add-In helper required by add-in {0} could not be registered.\r\nThis may be due to restricted permissions on the HKCU\\Software\\Classes key", DnaLibrary.CurrentLibrary.Name);
            }
            catch (Exception ex)
            {
                // If Excel is running with the /K switch it seems to indicate we're running 
                // in a COM-unfriendly mode where (sometimes) the COM add-in for the ribbon won't load. 
                // We skip the log display in this case.
                // CONSIDER: How would an add-in know that its COM AddIn load failed in this case?
                if (Environment.CommandLine.Contains(" /K"))
                {
                    Logger.ComAddIn.Info("Load Ribbon/COM Add-In exception with /K in CommandLine \r\n{0}", ex.ToString());
                }
                else
                {
                    Logger.ComAddIn.Error(ex, "The Ribbon/COM add-in helper required by add-in {0} could not be registered.\r\nThis may be due to the helper add-in being disabled by Excel.\r\nTo repair, open Disabled items in the Options->Add-Ins page and re-enable target add-in, then restart Excel.\r\n\r\nError details: {1}", DnaLibrary.CurrentLibrary.Name, ex.ToString());
                }
            }
        }

        internal static void UnloadComAddIns()
        {
            CultureInfo ci = new CultureInfo(1033);
            foreach (object comAddIn in loadedComAddIns.ToArray()) // Disconnecting an add-in removes it from loadedComAddIns, so we operate on a copy of the collection.
            {
                comAddIn.GetType().InvokeMember("Connect", System.Reflection.BindingFlags.SetProperty, null, comAddIn, new object[] { false }, ci);
                Logger.ComAddIn.Info("Ribbon/COM Add-In Unloaded.");
            }
        }

    }

}
