using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExcelDna.Integration;
using ExcelDna.Integration.ComInterop.Generator;
using ExcelDna.Integration.ComInterop.Generator.Interfaces;
using ExcelDna.Integration.CustomUI;

namespace ExcelDna.SourceGenerator.NativeAOT.Tests
{
    // Reproduction / regression tests for the NativeAOT ribbon callback dispatch
    // (reported broken at https://groups.google.com/g/exceldna/c/qwQLiufM5d4).
    //
    // These exercise the hand-rolled IDispatch implementation in
    // ExcelDna.Integration.ComInterop.Generator.ExcelRibbon directly, without Excel:
    //   - get* callbacks must marshal their return value back to Excel (pVarResult)
    //   - multi-argument callbacks (control + extra args) must be dispatched
    public class RibbonDispatch
    {
        public class TestRibbon : IExcelRibbon
        {
            public static bool LastPressed;
            public static string? LastText;

            public string GetCustomUI(string RibbonID) => "";

            // Value-returning callbacks - before the fix these returned nothing to Excel.
            public string GetLabel(RibbonControl control) => "the-label";
            public bool GetEnabled(RibbonControl control) => true;
            public int GetItemCount(RibbonControl control) => 3;

            // Multi-argument callbacks - before the fix (cArgs == 1 guard) these never ran.
            public void OnToggle(RibbonControl control, bool pressed) => LastPressed = pressed;
            public void OnChange(RibbonControl control, string text) => LastText = text;
        }

        private static int GetDispId(ExcelRibbon ribbon, string name)
        {
            int[] ids = new int[1];
            ribbon.GetIDsOfNames(Guid.Empty, new[] { name }, 1, 0, ids);
            return ids[0];
        }

        private static ExcelRibbon CreateRibbon()
        {
            return new ExcelRibbon(new TypeHelper<TestRibbon>(typeof(TestRibbon).GetMethods()));
        }

        // Invoke a callback and read back the variant the dispatcher wrote to pVarResult.
        private static object? InvokeWithResult(ExcelRibbon ribbon, int dispId, DispParams dispParams)
        {
            nint pVarResult = Marshal.AllocHGlobal(Marshal.SizeOf<VariantNative>());
            try
            {
                Marshal.StructureToPtr(default(VariantNative), pVarResult, false);
                ribbon.Invoke(dispId, Guid.Empty, 0, (ushort)INVOKEKIND.INVOKE_FUNC, in dispParams, pVarResult, 0, 0);
                return VariantMarshaller.ConvertToManaged(Marshal.PtrToStructure<VariantNative>(pVarResult)).Value;
            }
            finally
            {
                Marshal.FreeHGlobal(pVarResult);
            }
        }

        private static DispParams OneControlArg()
        {
            return new DispParams { cArgs = 1, rgvarg = new[] { new Variant(null) } };
        }

        [Fact]
        public void GetIDsOfNames_FindsCallback_AndLoadImage()
        {
            ExcelRibbon ribbon = CreateRibbon();

            Assert.True(GetDispId(ribbon, "GetLabel") >= 0);
            Assert.True(GetDispId(ribbon, "DoesNotExist") < 0);
            // LoadImage is the built-in callback dispatched past the end of the methods array.
            Assert.True(GetDispId(ribbon, "LoadImage") >= 0);
        }

        [Fact]
        public void Invoke_GetLabel_MarshalsStringReturnValueBackToExcel()
        {
            ExcelRibbon ribbon = CreateRibbon();
            object? result = InvokeWithResult(ribbon, GetDispId(ribbon, "GetLabel"), OneControlArg());
            Assert.Equal("the-label", result);
        }

        [Fact]
        public void Invoke_GetEnabled_MarshalsBoolReturnValueBackToExcel()
        {
            ExcelRibbon ribbon = CreateRibbon();
            object? result = InvokeWithResult(ribbon, GetDispId(ribbon, "GetEnabled"), OneControlArg());
            Assert.Equal(true, result);
        }

        [Fact]
        public void Invoke_GetItemCount_MarshalsIntReturnValueBackToExcel()
        {
            ExcelRibbon ribbon = CreateRibbon();
            object? result = InvokeWithResult(ribbon, GetDispId(ribbon, "GetItemCount"), OneControlArg());
            Assert.Equal(3, result);
        }

        [Fact]
        public void Invoke_OnToggle_BindsPressedBoolArgument()
        {
            ExcelRibbon ribbon = CreateRibbon();
            TestRibbon.LastPressed = false;
            DispParams dispParams = new DispParams { cArgs = 2, rgvarg = new[] { new Variant(null), new Variant(true) } };
            ribbon.Invoke(GetDispId(ribbon, "OnToggle"), Guid.Empty, 0, (ushort)INVOKEKIND.INVOKE_FUNC, in dispParams, 0, 0, 0);
            Assert.True(TestRibbon.LastPressed);
        }

        [Fact]
        public void Invoke_OnChange_BindsTextStringArgument()
        {
            ExcelRibbon ribbon = CreateRibbon();
            TestRibbon.LastText = null;
            DispParams dispParams = new DispParams { cArgs = 2, rgvarg = new[] { new Variant(null), new Variant("hello") } };
            ribbon.Invoke(GetDispId(ribbon, "OnChange"), Guid.Empty, 0, (ushort)INVOKEKIND.INVOKE_FUNC, in dispParams, 0, 0, 0);
            Assert.Equal("hello", TestRibbon.LastText);
        }
    }
}
