using ExcelDna.Loader;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelDna.ManagedHost
{
    public unsafe class AddInInitialize
    {
#if NETFRAMEWORK
        public static bool Initialize32(int xlAddInExportInfoAddress, int hModuleXll, string pathXll)
        {
            // NOTE: The sequence here is important - we install the AssemblyManage which can resolve packed assemblies
            //       before calling LoadIntegration, which will be the first time we try to resolve ExcelDna.Integration
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll);
            // TODO: Load up the DnaFile and Assembly names ?
            AppDomain.CurrentDomain.AssemblyResolve += (object sender, ResolveEventArgs args) => AssemblyManager.AssemblyResolve(new AssemblyName(args.Name));
            return XlAddInInitialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, AssemblyManager.GetResourceBytes, Logger.SetIntegrationTraceSource);
        }

        public static bool Initialize64(long xlAddInExportInfoAddress, long hModuleXll, string pathXll)
        {
            // NOTE: The sequence here is important - we install the AssemblyManage which can resolve packed assemblies
            //       before calling LoadIntegration, which will be the first time we try to resolve ExcelDna.Integration
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll);
            AppDomain.CurrentDomain.AssemblyResolve += (object sender, ResolveEventArgs args) => AssemblyManager.AssemblyResolve(new AssemblyName(args.Name));
            return XlAddInInitialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, AssemblyManager.GetResourceBytes, Logger.SetIntegrationTraceSource);
        }
#endif

#if NETCOREAPP
        static ExcelDnaAssemblyLoadContext _alc;

        [UnmanagedCallersOnly]
        public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL)
        {
            string pathXll = Marshal.PtrToStringUni((IntPtr)pPathXLL);
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll);
            _alc = new ExcelDnaAssemblyLoadContext(pathXll);

            var initOK = XlAddInInitialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, AssemblyManager.GetResourceBytes, Logger.SetIntegrationTraceSource);

            return initOK ? (short)1 : (short)0;
        }
#endif


        // NOTE: We need this code to be in a separate method, so that the assembly resolution for ExcelDna.Loader runs after the AssemblyManager is installed.
        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool XlAddInInitialize(IntPtr xlAddInExportInfoAddress, IntPtr hModuleXll, string pathXll, Func<string, int, byte[]> getResourceBytes, Action<TraceSource> setIntegrationTraceSource)
        {
            return XlAddIn.Initialize(xlAddInExportInfoAddress, hModuleXll, pathXll, getResourceBytes, setIntegrationTraceSource);
        }

    }
}
