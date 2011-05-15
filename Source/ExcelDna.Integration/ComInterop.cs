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
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections;
using System.Threading;

namespace ExcelDna.ComInterop
{
    internal class ComAPI
    {
        public const HRESULT S_OK = 0;
        public const HRESULT S_FALSE = 1;
        public const HRESULT CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const HRESULT CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
        public const HRESULT E_INVALIDARG = unchecked((int)0x80070057);
        public const HRESULT E_NOINTERFACE = unchecked((int)0x80004002);
        public const HRESULT E_UNEXPECTED = unchecked((int)0x8000FFFF);
        public const string gstrIUnknown = "00000000-0000-0000-C000-000000000046";
        public static readonly Guid guidIUnknown = new Guid(gstrIUnknown);
        public const string gstrIClassFactory = "00000001-0000-0000-C000-000000000046";
        public static readonly Guid guidIClassFactory = new Guid(gstrIClassFactory);
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

        [DllImport("ole32.dll")]
        public static extern HRESULT CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport("oleaut32.dll")]
        public static extern HRESULT LoadTypeLib([MarshalAs(UnmanagedType.LPWStr)] string fileName, out ITypeLib typeLib);
    }
}

#region COM Import declares for Office / Excel interfaces
// I'm trying to keep the imported parts quite limited.

#region Assembly Extensibility.dll, v1.0.3705
// c:\Program Files\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Common\Extensibility.dll
#endregion

// Changed namespace to remove type name clash
// The interfaces should still work fine due to the interface GUIDs
// Someday we might add .NET 4 Type exuivalence to properly sort the problem out.

// namespace Extensibility
namespace ExcelDna.Integration.Extensibility
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

// Changed namespace to remove type name clash
// The interfaces should still work fine due to the interface GUIDs
// Someday we might add .NET 4 Type exuivalence to properly sort the problem out.

//namespace Microsoft.Office.Core
namespace ExcelDna.Integration.CustomUI
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
        void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] ICTPFactory CTPFactoryInst);
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

    //// Actually from System.Windows.Forms.UnsafeNativeMethods
    //[ComImport]
    //[Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    //public interface IPictureDisp
    //{
    //    IntPtr Handle { get; }
    //    IntPtr HPal { get; }
    //    short PictureType { get; }
    //    int Width { get; }
    //    int Height { get; }
    //    void Render(IntPtr hdc, int x, int y, int cx, int cy, int xSrc, int ySrc, int cxSrc, int cySrc);
    //}

    // CONSIDER: Review all these CustomTaskPane declarations
    //           - horribly copied and hacked together from the office PIA.
    // Anyway - not late-bound like everything else - just for supporting the two events
    [ComImport, Guid("000C033D-0000-0000-C000-000000000046"), TypeLibType((short)0x10c0)]
    public interface ICTPFactory
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        CustomTaskPane CreateCTP([In, MarshalAs(UnmanagedType.BStr)] string CTPAxID, [In, MarshalAs(UnmanagedType.BStr)] string CTPTitle, [In, Optional, MarshalAs(UnmanagedType.Struct)] object CTPParentWindow);
    }

    public enum MsoCTPDockPosition
    {
        msoCTPDockPositionLeft,
        msoCTPDockPositionTop,
        msoCTPDockPositionRight,
        msoCTPDockPositionBottom,
        msoCTPDockPositionFloating
    }

    public enum MsoCTPDockPositionRestrict
    {
        msoCTPDockPositionRestrictNone,
        msoCTPDockPositionRestrictNoChange,
        msoCTPDockPositionRestrictNoHorizontal,
        msoCTPDockPositionRestrictNoVertical
    }

    // CONSIDER: SHould we be worried about set/get order? (see Adam Nathan page 991)
    [ComImport, Guid("000C033B-0000-0000-C000-000000000046"), TypeLibType((short)0x10c0), DefaultMember("Title")]
    public interface ICustomTaskPane
    {
        [DispId(0)]
        string Title { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)] get; }
        [DispId(1)]
        object Application { [return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)] get; }
        [DispId(2)]
        object Window { [return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)] get; }
        [DispId(3)]
        bool Visible { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)] get; [param: In] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)] set; }
        [DispId(4)]
        object ContentControl { [return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)] get; }
        [DispId(5)]
        int Height { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)] get; [param: In] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)] set; }
        [DispId(6)]
        int Width { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)] get; [param: In] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)] set; }
        [DispId(7)]
        MsoCTPDockPosition DockPosition { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)] get; [param: In] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)] set; }
        [DispId(8)]
        MsoCTPDockPositionRestrict DockPositionRestrict { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)] get; [param: In] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)] set; }
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
        void Delete();
    }

    [ComImport, TypeLibType((short)0x1010), InterfaceType((short)2), Guid("000C033C-0000-0000-C000-000000000046")]
    internal interface _CustomTaskPaneEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void VisibleStateChange([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void DockPositionStateChange([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
    }

    // Awesome write-up for COM events: http://blogs.msdn.com/b/varunsekhri/archive/2007/09/18/how-do-we-talk-with-com-the-language-of-events-and-delegates.aspx

    [ComVisible(false)]
    public delegate void CustomTaskPaneEvents_DockPositionStateChangeEventHandler([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
    [ComVisible(false)]
    public delegate void CustomTaskPaneEvents_VisibleStateChangeEventHandler([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);

    [ComEventInterface(typeof(_CustomTaskPaneEvents), typeof(_CustomTaskPaneEvents_EventProvider)), TypeLibType((short)0x10), ComVisible(false)]
    public interface ICustomTaskPaneEvents
    {
        // Events
        event CustomTaskPaneEvents_DockPositionStateChangeEventHandler DockPositionStateChange;
        event CustomTaskPaneEvents_VisibleStateChangeEventHandler VisibleStateChange;
    }

    [ComImport, TypeLibType((short)0x10d0), Guid("8A64A872-FC6B-4D4A-926E-3A3689562C1C")]
    internal interface CustomTaskPaneEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void VisibleStateChange([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void DockPositionStateChange([In, MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
    }

    internal sealed class _CustomTaskPaneEvents_EventProvider : ICustomTaskPaneEvents, IDisposable
    {
        private bool _disposed;

        // Fields
        private ArrayList m_aEventSinkHelpers;
        private IConnectionPoint m_ConnectionPoint;
        private IConnectionPointContainer m_ConnectionPointContainer;

        // Methods
        public _CustomTaskPaneEvents_EventProvider(object obj1)
        {
            _disposed = false;
            this.m_ConnectionPointContainer = (IConnectionPointContainer)obj1;
        }

        event CustomTaskPaneEvents_DockPositionStateChangeEventHandler ICustomTaskPaneEvents.DockPositionStateChange
        {
            add
            {
                Monitor.Enter(this);
                try
                {
                    if (this.m_ConnectionPoint == null)
                    {
                        this.Init();
                    }
                    _CustomTaskPaneEvents_SinkHelper customTaskPaneEvents_SinkHelper = new _CustomTaskPaneEvents_SinkHelper();
                    int dwCookie = 0;
                    this.m_ConnectionPoint.Advise((object)customTaskPaneEvents_SinkHelper, out dwCookie);
                    customTaskPaneEvents_SinkHelper.m_dwCookie = dwCookie;
                    customTaskPaneEvents_SinkHelper.m_DockPositionStateChangeDelegate = value;
                    this.m_aEventSinkHelpers.Add((object)customTaskPaneEvents_SinkHelper);
                }
                finally
                {
                    Monitor.Exit(this);
                }
            }
            remove
            {
                Monitor.Enter(this);
                try
                {
                    int count = this.m_aEventSinkHelpers.Count;
                    if (count > 0)
                    {
                        _CustomTaskPaneEvents_SinkHelper customTaskPaneEvents_SinkHelper;
                        for (int i = 0; i < count; i++)
                        {
                            customTaskPaneEvents_SinkHelper = (_CustomTaskPaneEvents_SinkHelper)m_aEventSinkHelpers[i];
                            if (customTaskPaneEvents_SinkHelper.m_DockPositionStateChangeDelegate != null && customTaskPaneEvents_SinkHelper.m_DockPositionStateChangeDelegate.Equals((object)value))
                            {
                                m_aEventSinkHelpers.RemoveAt(i);
                                m_ConnectionPoint.Unadvise(customTaskPaneEvents_SinkHelper.m_dwCookie);
                                if (m_aEventSinkHelpers.Count == 0)
                                {
                                    Marshal.ReleaseComObject(this.m_ConnectionPoint);
                                    m_ConnectionPoint = null;
                                    m_aEventSinkHelpers = null;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    Monitor.Exit(this);
                }
            }
        }

        event CustomTaskPaneEvents_VisibleStateChangeEventHandler ICustomTaskPaneEvents.VisibleStateChange
        {
            add
            {
                Monitor.Enter(this);
                try
                {
                    if (this.m_ConnectionPoint == null)
                    {
                        this.Init();
                    }
                    _CustomTaskPaneEvents_SinkHelper customTaskPaneEvents_SinkHelper = new _CustomTaskPaneEvents_SinkHelper();
                    int dwCookie = 0;
                    this.m_ConnectionPoint.Advise((object)customTaskPaneEvents_SinkHelper, out dwCookie);
                    customTaskPaneEvents_SinkHelper.m_dwCookie = dwCookie;
                    customTaskPaneEvents_SinkHelper.m_VisibleStateChangeDelegate = value;
                    this.m_aEventSinkHelpers.Add((object)customTaskPaneEvents_SinkHelper);
                }
                finally
                {
                    Monitor.Exit(this);
                }
            }
            remove
            {
                Monitor.Enter(this);
                try
                {
                    int count = this.m_aEventSinkHelpers.Count;
                    if (count > 0)
                    {
                        _CustomTaskPaneEvents_SinkHelper customTaskPaneEvents_SinkHelper;
                       for (int i = 0; i < count; i++)
                       {
                            customTaskPaneEvents_SinkHelper = (_CustomTaskPaneEvents_SinkHelper)m_aEventSinkHelpers[i];
                            if (customTaskPaneEvents_SinkHelper.m_VisibleStateChangeDelegate != null && customTaskPaneEvents_SinkHelper.m_VisibleStateChangeDelegate.Equals((object)value))
                            {
                                m_aEventSinkHelpers.RemoveAt(i);
                                m_ConnectionPoint.Unadvise(customTaskPaneEvents_SinkHelper.m_dwCookie);
                                if (m_aEventSinkHelpers.Count == 0)
                                {
                                    Marshal.ReleaseComObject(this.m_ConnectionPoint);
                                    m_ConnectionPoint = null;
                                    m_aEventSinkHelpers = null;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    Monitor.Exit(this);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Not thread-safe...
            if (!_disposed)
            {
                // if (disposing)
                // {
                //     // Here comes explicit free of other managed disposable objects.
                // }

                // Here comes clean-up
                Cleanup();
                _disposed = true;
            }
        }

        ~_CustomTaskPaneEvents_EventProvider()
        {
            Dispose(false);
        }

        void Cleanup()
        {
            Monitor.Enter(this);
            try
            {
                if (this.m_ConnectionPoint != null)
                {
                    int count = this.m_aEventSinkHelpers.Count;
                    int num2 = 0;
                    if (0 < count)
                    {
                        do
                        {
                            _CustomTaskPaneEvents_SinkHelper helper = (_CustomTaskPaneEvents_SinkHelper)this.m_aEventSinkHelpers[num2];
                            this.m_ConnectionPoint.Unadvise(helper.m_dwCookie);
                            num2++;
                        }
                        while (num2 < count);
                    }
                    Marshal.ReleaseComObject(this.m_ConnectionPoint);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                Monitor.Exit(this);
            }
        }

        private void Init()
        {
            IConnectionPoint ppCP = null;
            byte[] b = new byte[] { 60, 3, 12, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 70 };
            Guid riid = new Guid(b);
            this.m_ConnectionPointContainer.FindConnectionPoint(ref riid, out ppCP);
            this.m_ConnectionPoint = ppCP;
            this.m_aEventSinkHelpers = new ArrayList();
        }
    }

    [ClassInterface(ClassInterfaceType.None)]
    internal sealed class _CustomTaskPaneEvents_SinkHelper : _CustomTaskPaneEvents
    {
        // Fields
        public int m_dwCookie = 0;
        public CustomTaskPaneEvents_DockPositionStateChangeEventHandler m_DockPositionStateChangeDelegate = null;
        public CustomTaskPaneEvents_VisibleStateChangeEventHandler m_VisibleStateChangeDelegate = null;

        // Methods
        internal _CustomTaskPaneEvents_SinkHelper()
        {
        }

        public void DockPositionStateChange(CustomTaskPane pane1)
        {
            if (this.m_DockPositionStateChangeDelegate != null)
            {
                this.m_DockPositionStateChangeDelegate(pane1);
            }
        }

        public void VisibleStateChange(CustomTaskPane pane1)
        {
            if (this.m_VisibleStateChangeDelegate != null)
            {
                this.m_VisibleStateChangeDelegate(pane1);
            }
        }
    }

    [ComImport, Guid("000C033B-0000-0000-C000-000000000046")]
    public interface CustomTaskPane : ICustomTaskPane, ICustomTaskPaneEvents
    {
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
#endregion

