namespace ExcelDna.SourceGenerator.NativeAOT.Tests
{
    public class MarshalGenerator
    {
        [Fact]
        public void Empty()
        {
            Verify("",
                """
                class DirectMarshalTypeAdapter : ExcelDna.Registration.StaticRegistration.IDirectMarshalTypeAdapter
                {
                    private class XlDirectMarshalLazy
                    {
                        readonly Lazy<Delegate> _delegate;

                        public XlDirectMarshalLazy(Func<Delegate> delegateFactory)
                        {
                            _delegate = new Lazy<Delegate>(delegateFactory, LazyThreadSafetyMode.PublicationOnly);
                        }
                    



                    }

                    public nint GetActionPointerForDelegate(Delegate d, int parameters)
                    {


                        throw new NotImplementedException($"GetActionPointerForDelegate {parameters}");
                    }

                    public Type GetActionType(int parameters)
                    {


                        throw new NotImplementedException($"GetActionType {parameters}");
                    }

                    public nint GetFunctionPointerForDelegate(Delegate d, int parameters)
                    {


                        throw new NotImplementedException($"GetFunctionPointerForDelegate {parameters}");
                    }

                    public Type GetFunctionType(int parameters)
                    {


                        throw new NotImplementedException($"GetFunctionType {parameters}");
                    }

                    public Delegate CreateActionDelegate(Func<Delegate> delegateFactory, int parameters)
                    {
                        var lazyLambda = new XlDirectMarshalLazy(delegateFactory);
                        var delegateType = GetActionType(parameters);
                        var method = GetLazyAct(parameters);
                        return Delegate.CreateDelegate(delegateType, lazyLambda, method);
                    }

                    public Delegate CreateFunctionDelegate(Func<Delegate> delegateFactory, int parameters)
                    {
                        var lazyLambda = new XlDirectMarshalLazy(delegateFactory);
                        var delegateType = GetFunctionType(parameters);
                        var method = GetLazyFunc(parameters);
                        return Delegate.CreateDelegate(delegateType, lazyLambda, method);
                    }

                    private MethodInfo GetLazyAct(int parameters)
                    {


                        throw new NotImplementedException($"GetLazyAct {parameters}");
                    }

                    private MethodInfo GetLazyFunc(int parameters)
                    {


                        throw new NotImplementedException($"GetLazyFunc {parameters}");
                    }




                }
                """
                );
        }

        [Fact]
        public void Full()
        {
            Verify("""
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class Functions
                    {
                        [ExcelCommand(MenuText = "NativeCommandHello")]
                        public static void NativeCommandHello()
                        {
                        }

                        [ExcelFunction]
                        public static string NativeHello0()
                        {
                            return "";
                        }

                        [ExcelFunction]
                        public static string NativeHello1(string name)
                        {
                            return "";
                        }

                        [ExcelFunction]
                        public static string NativeHello2(string name1, string name2)
                        {
                            return "";
                        }

                        [ExcelFunction]
                        public static string NativeParams(object input, string QtherInpEt, params object[] args)
                        {
                            return "";
                        }
                    }
                }
                """,
                """
                class DirectMarshalTypeAdapter : ExcelDna.Registration.StaticRegistration.IDirectMarshalTypeAdapter
                {
                    private class XlDirectMarshalLazy
                    {
                        readonly Lazy<Delegate> _delegate;

                        public XlDirectMarshalLazy(Func<Delegate> delegateFactory)
                        {
                            _delegate = new Lazy<Delegate>(delegateFactory, LazyThreadSafetyMode.PublicationOnly);
                        }
                    
                public void Act0() => ((XlAct0)_delegate.Value)();

                public IntPtr Func0() => ((XlFunc0)_delegate.Value)();
                public IntPtr Func1(IntPtr p1) => ((XlFunc1)_delegate.Value)(p1);
                public IntPtr Func2(IntPtr p1, IntPtr p2) => ((XlFunc2)_delegate.Value)(p1, p2);
                public IntPtr Func16(IntPtr p1, IntPtr p2, IntPtr p3, IntPtr p4, IntPtr p5, IntPtr p6, IntPtr p7, IntPtr p8, IntPtr p9, IntPtr p10, IntPtr p11, IntPtr p12, IntPtr p13, IntPtr p14, IntPtr p15, IntPtr p16) => ((XlFunc16)_delegate.Value)(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16);
                    }

                    public nint GetActionPointerForDelegate(Delegate d, int parameters)
                    {
                switch (parameters)
                {
                case 0: return System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate<XlAct0>((XlAct0)d);
                }
                
                        throw new NotImplementedException($"GetActionPointerForDelegate {parameters}");
                    }

                    public Type GetActionType(int parameters)
                    {
                switch (parameters)
                {
                case 0: return typeof(XlAct0);
                }
                
                        throw new NotImplementedException($"GetActionType {parameters}");
                    }

                    public nint GetFunctionPointerForDelegate(Delegate d, int parameters)
                    {
                switch (parameters)
                {
                case 0: return System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate<XlFunc0>((XlFunc0)d);
                case 1: return System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate<XlFunc1>((XlFunc1)d);
                case 2: return System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate<XlFunc2>((XlFunc2)d);
                case 16: return System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate<XlFunc16>((XlFunc16)d);
                }

                        throw new NotImplementedException($"GetFunctionPointerForDelegate {parameters}");
                    }

                    public Type GetFunctionType(int parameters)
                    {
                switch (parameters)
                {
                case 0: return typeof(XlFunc0);
                case 1: return typeof(XlFunc1);
                case 2: return typeof(XlFunc2);
                case 16: return typeof(XlFunc16);
                }

                        throw new NotImplementedException($"GetFunctionType {parameters}");
                    }

                    public Delegate CreateActionDelegate(Func<Delegate> delegateFactory, int parameters)
                    {
                        var lazyLambda = new XlDirectMarshalLazy(delegateFactory);
                        var delegateType = GetActionType(parameters);
                        var method = GetLazyAct(parameters);
                        return Delegate.CreateDelegate(delegateType, lazyLambda, method);
                    }

                    public Delegate CreateFunctionDelegate(Func<Delegate> delegateFactory, int parameters)
                    {
                        var lazyLambda = new XlDirectMarshalLazy(delegateFactory);
                        var delegateType = GetFunctionType(parameters);
                        var method = GetLazyFunc(parameters);
                        return Delegate.CreateDelegate(delegateType, lazyLambda, method);
                    }

                    private MethodInfo GetLazyAct(int parameters)
                    {
                switch (parameters)
                {
                case 0: return typeof(XlDirectMarshalLazy).GetMethod("Act0")!;
                }
                
                        throw new NotImplementedException($"GetLazyAct {parameters}");
                    }

                    private MethodInfo GetLazyFunc(int parameters)
                    {
                switch (parameters)
                {
                case 0: return typeof(XlDirectMarshalLazy).GetMethod("Func0")!;
                case 1: return typeof(XlDirectMarshalLazy).GetMethod("Func1")!;
                case 2: return typeof(XlDirectMarshalLazy).GetMethod("Func2")!;
                case 16: return typeof(XlDirectMarshalLazy).GetMethod("Func16")!;
                }

                        throw new NotImplementedException($"GetLazyFunc {parameters}");
                    }

                private delegate void XlAct0();

                private delegate IntPtr XlFunc0();
                private delegate IntPtr XlFunc1(IntPtr p1);
                private delegate IntPtr XlFunc2(IntPtr p1, IntPtr p2);
                private delegate IntPtr XlFunc16(IntPtr p1, IntPtr p2, IntPtr p3, IntPtr p4, IntPtr p5, IntPtr p6, IntPtr p7, IntPtr p8, IntPtr p9, IntPtr p10, IntPtr p11, IntPtr p12, IntPtr p13, IntPtr p14, IntPtr p15, IntPtr p16);
                }
                """
                );
        }

        private static void Verify(string sourceCode, string body)
        {
            string template = """
// <auto-generated/>
using System;
using System.Reflection;
using System.Threading;

#nullable enable

namespace ExcelDna.SourceGenerator.NativeAOT
{
    [DIRECT-MARSHAL-TYPE-ADAPTER]
}
""";

            SourceGeneratorDriver.Verify(sourceCode, null, template.Replace("[DIRECT-MARSHAL-TYPE-ADAPTER]", body));
        }
    }
}
