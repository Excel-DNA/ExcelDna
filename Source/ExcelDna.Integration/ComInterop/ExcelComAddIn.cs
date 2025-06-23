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
    // NOTE: A COM add-in might be able to know phow Excel started by looking in the Custom() values passed to OnConnection
    // See: https://msdn.microsoft.com/en-us/library/aa189748%28office.10%29.aspx
    // Custom()	Variant	An array of Variant type values that provides additional data. 
    // The numeric value of the first element in this array indicates how the host application was started: 
    // from the user interface (1), 
    // by embedding a document created in the host application in another application (2), 
    // or through Automation (3).

    [ComVisible(true)]
#pragma warning disable CS0618 // Type or member is obsolete (but probably not forever)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618 // Type or member is obsolete
    public class ExcelComAddIn : IDTExtensibility2
    {
        internal DnaLibrary DnaLibrary { get; set; }
        private string _progId;
        private object _addInInst;
        protected string ProgId
        {
            get { return _progId; }
        }

        internal void SetProgId(string progId)
        {
            _progId = progId;
        }

        public string FriendlyName { get; protected set; }
        public string Description { get; protected set; }

        #region IDTExtensibility2 interface
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _addInInst = AddInInst;
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnConnection");
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            ExcelComAddInHelper.OnUnloadComAddIn(this, _addInInst);
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnBeginShutdown");
        }
        #endregion
    }
}
