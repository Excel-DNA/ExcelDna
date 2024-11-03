using ExcelDna.Integration;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelDna.Test
{
    public unsafe class AddInInitialize
    {
        [UnmanagedCallersOnly(EntryPoint = "Initialize", CallConvs = new[] { typeof(CallConvCdecl) })]
        public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
        {
            ExcelDna.Integration.DnaLibrary.MethodsForRegistration.Add(typeof(Functions).GetMethod(nameof(Functions.MyHello))!);

            return ExcelDna.ManagedHost.AddInInitialize.Initialize(xlAddInExportInfoAddress, hModuleXll, pPathXLL, disableAssemblyContextUnload, pTempDirPath);
        }
    }
}
