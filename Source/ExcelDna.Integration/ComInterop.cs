using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExcelDna.ComInterop;
using Microsoft.Win32;

using HRESULT = System.Int32;
using IID     = System.Guid;
using CLSID   = System.Guid;
using DWORD   = System.UInt32;

namespace ExcelDna.ComInterop
{
    internal class ComAPI
    {
        public const HRESULT S_OK = 0;
        public const HRESULT CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const HRESULT E_NOINTERFACE = unchecked((int)0x80004002);
        public const string gstrIUnknown = "00000000-0000-0000-C000-000000000046";
        public static readonly Guid guidIUnknown = new Guid(gstrIUnknown);
        public const string gstrIClassFactory = "00000001-0000-0000-C000-000000000046";
        public const string gstrIDTExtensibility2 = "B65AD801-ABAF-11D0-BB8B-00A0C90F2744";
        public static readonly Guid guidIDTExtensibility2 = new Guid(gstrIDTExtensibility2);
        public const string gstrIRibbonExtensibility = "000C0396-0000-0000-C000-000000000046";
        public const string gstrIRtdServer = "EC0E6191-DB51-11D3-8F3E-00C04F3651B8";
        public static readonly Guid guidIRtdServer = new Guid(gstrIRtdServer);
        public const string gstrIRTDUpdateEvent = "A43788C1-D91B-11D3-8F39-00C04F3651B8";

        [ComImport]
        [Guid(gstrIClassFactory)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IClassFactory
        {
            // For HRESULTs use 
            [PreserveSig]
            HRESULT CreateInstance([In] IntPtr pUnkOuter,
                                [In] ref IID riid,
                                [Out] out IntPtr ppvObject);

            [PreserveSig]
            HRESULT LockServer([In, MarshalAs(UnmanagedType.VariantBool)] bool fLock);

            //        HRESULT STDMETHODCALLTYPE CreateInstance( 
            //        /* [unique][in] */ IUnknown *pUnkOuter,
            //        /* [in] */ REFIID riid,
            //        /* [iid_is][out] */ void **ppvObject) = 0;

            //    virtual /* [local] */ HRESULT STDMETHODCALLTYPE LockServer( 
            //        /* [in] */ BOOL fLock) = 0;
        }

        [DllImport("Ole32.dll")]
        public static extern HRESULT CoRegisterClassObject([In] ref CLSID rclsid, IntPtr pUnk,
                DWORD dwClsContext, DWORD flags, out DWORD lpdwRegister);

        [DllImport("Ole32.dll")]
        public static extern HRESULT CoRevokeClassObject(DWORD dwRegister);

        [DllImport("Ole32.dll")]
        public static extern HRESULT GetRunningObjectTable(DWORD reserved, out IRunningObjectTable pprot);

        [DllImport("Ole32.dll")]
        public static extern HRESULT CreateFileMoniker(string pathName, out IMoniker ppmk);
    }
}

#region COM Import declares for Office / Excel interfaces
// I'm trying to keep the imported parts quite limited.

#region Assembly Extensibility.dll, v1.0.3705
// c:\Program Files\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Common\Extensibility.dll
#endregion

namespace Extensibility
{
    [ComImport]
    [Guid(ComAPI.gstrIDTExtensibility2)]
    //[TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application, [In] ext_ConnectMode ConnectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5,
    }

    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3,
    }
}

#region Assembly Office.dll, v2.0.50727
// C:\WINDOWS\assembly\GAC_MSIL\Office\14.0.0.0__71e9bce111e9429c\Office.dll
#endregion
namespace Microsoft.Office.Core
{
    [ComImport]
    [Guid(ComAPI.gstrIRibbonExtensibility)]
    // [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    // [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FNonExtensible | TypeLibTypeFlags.FDual)] // 4288
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable(IntPtr /*ICTPFactory*/ CTPFactoryInst);
    }

    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    // [TypeLibType((short)0x1040)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { [return: MarshalAs(UnmanagedType.BStr)] [DispId(1)] get; }
        [DispId(2)]
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] [DispId(2)] get; }
        [DispId(3)]
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] [DispId(3)] get; }
    }

    [ComImport]
    [Guid("000C03A7-0000-0000-C000-000000000046")]
    // [TypeLibType((short)0x1040)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonUI
    {
        [DispId(1)]
        void Invalidate();
        [DispId(2)]
        void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
    }

    // Actually from System.Windows.Forms.UnsafeNativeMethods
    [ComImport]
    [Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IPictureDisp
    {
        IntPtr Handle { get; }
        IntPtr HPal { get; }
        short PictureType { get; }
        int Width { get; }
        int Height { get; }
        void Render(IntPtr hdc, int x, int y, int cx, int cy, int xSrc, int ySrc, int cxSrc, int cySrc);
    }
}

#region Assembly Microsoft.Office.Interop.Excel.dll, v1.1.4322
// C:\Program Files\Microsoft Visual Studio 9.0\Visual Studio Tools for Office\PIA\Office11\Microsoft.Office.Interop.Excel.dll
#endregion

// We want to export these names so that Rtd Servers can be implemented without referencing interop assemblies.
// So the name needs to change, to prevent conflicts.
// Since we don't have Type Equivalence in runtime versions before CLR 4, we need to be careful when using these
// types internally - the ExcelRtdWrapper implements the type re-unification.
namespace ExcelDna.Integration.Rtd
{
    // Summary:
    //     Represents an interface for a real-time data server.
    [ComImport]
    [Guid(ComAPI.gstrIRtdServer)]
    //[TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    //[TypeIdentity(ComAPI.guidIRtdServer, "Microsoft.Office.Interop.IRtdServer")]
    public interface IRtdServer
    {
        //
        // Summary:
        //     The ServerStart method is called immediately after a real-time data server
        //     is instantiated. Negative value or zero indicates failure to start the server;
        //     positive value indicates success.
        //
        // Parameters:
        //   CallbackObject:
        //     Required Microsoft.Office.Interop.Excel.IRTDUpdateEvent object. The callback
        //     object.
        [DispId(10)]
        int ServerStart(IRTDUpdateEvent CallbackObject);
        // Summary:
        //     Adds new topics from a real-time data server. The ConnectData method is called
        //     when a file is opened that contains real-time data functions or when a user
        //     types in a new formula which contains the RTD function.
        //
        // Parameters:
        //   TopicID:
        //     Required Integer. A unique value, assigned by Microsoft Excel, which identifies
        //     the topic.
        //
        //   GetNewValues:
        //     Required Boolean. True to determine if new values are to be acquired.
        //
        //   Strings:
        //     Required Object. A single-dimensional array of strings identifying the topic.
        [DispId(11)]
        object ConnectData(int topicId, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array strings, ref bool newValues);

        //
        // Summary:
        //     This method is called by Microsoft Excel to get new data.
        //
        // Parameters:
        //   TopicCount:
        //     Required Integer. The RTD server must change the value of the TopicCount
        //     to the number of elements in the array returned.
        [DispId(12)]
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        Array RefreshData(ref int topicCount);
        
        //
        // Summary:
        //     Notifies a real-time data (RTD) server application that a topic is no longer
        //     in use.
        //
        // Parameters:
        //   TopicID:
        //     Required Integer. A unique value assigned to the topic assigned by Microsoft
        //     Excel.
        [DispId(13)]
        void DisconnectData(int topicID);
        [DispId(14)]
        int Heartbeat();
        //
        // Summary:
        //     Terminates the connection to the real-time data server.
        [DispId(15)]
        void ServerTerminate();
    }

    // Summary:
    //     Represents real-time data update events.
    [ComImport]
    [Guid(ComAPI.gstrIRTDUpdateEvent)]
    // [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)] // 4160
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    // [TypeIdentity(ComAPI.guidIRTDUpdateEvent, "Microsoft.Office.Interop.IRTDUpdateEvent")]
    public interface IRTDUpdateEvent
    {
        // Summary:
        //     Instructs the real-time data server (RTD) to disconnect from the specified
        //     Microsoft.Office.Interop.Excel.IRTDUpdateEvent object.
        [DispId(10)]
        void UpdateNotify();
        [DispId(11)]
        int HeartbeatInterval { get; set; }

        [DispId(12)]
        void Disconnect();
    }
}

#region Assembly Office.dll, v1.1.4322
// C:\WINDOWS\assembly\GAC\Office\12.0.0.0__71e9bce111e9429c\Office.dll
#endregion

namespace Microsoft.Office.Core
{
    internal enum MsoButtonStyle
    {
        msoButtonAutomatic = 0,
        msoButtonIcon = 1,
        msoButtonCaption = 2,
        msoButtonIconAndCaption = 3,
        msoButtonIconAndWrapCaption = 7,
        msoButtonIconAndCaptionBelow = 11,
        msoButtonWrapCaption = 14,
        msoButtonIconAndWrapCaptionBelow = 15,
    }

    internal enum MsoControlType
    {
        msoControlCustom = 0,
        msoControlButton = 1,
        msoControlEdit = 2,
        msoControlDropdown = 3,
        msoControlComboBox = 4,
        msoControlButtonDropdown = 5,
        msoControlSplitDropdown = 6,
        msoControlOCXDropdown = 7,
        msoControlGenericDropdown = 8,
        msoControlGraphicDropdown = 9,
        msoControlPopup = 10,
        msoControlGraphicPopup = 11,
        msoControlButtonPopup = 12,
        msoControlSplitButtonPopup = 13,
        msoControlSplitButtonMRUPopup = 14,
        msoControlLabel = 15,
        msoControlExpandingGrid = 16,
        msoControlSplitExpandingGrid = 17,
        msoControlGrid = 18,
        msoControlGauge = 19,
        msoControlGraphicCombo = 20,
        msoControlPane = 21,
        msoControlActiveX = 22,
        msoControlSpinner = 23,
        msoControlLabelEx = 24,
        msoControlWorkPane = 25,
        msoControlAutoCompleteCombo = 26,
    }
}
#endregion

