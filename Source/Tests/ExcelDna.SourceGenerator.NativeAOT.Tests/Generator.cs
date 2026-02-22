namespace ExcelDna.SourceGenerator.NativeAOT.Tests
{
    public class Generator
    {
        [Fact]
        public void Empty()
        {
            Verify("");
        }

        [Fact]
        public void Params()
        {
            Verify("""
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class Functions
                    {
                        [ExcelFunction]
                        public static string NativeParamsJoinString(string separator, params string[] values)
                        {
                            return string.Join(separator, values);
                        }
                    }
                }
                """, functions: """
                List<Type> typeRefs = new List<Type>();
                ExcelDna.Registration.StaticRegistration.MethodsForRegistration.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Functions).GetMethod("NativeParamsJoinString")!);
                typeRefs.Add(typeof(Func<string, string[], string>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<string, string[], string>>));
                typeRefs.Add(typeof(Func<object, string>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<object, string>>));
                typeRefs.Add(typeof(Func<object, string[]>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<object, string[]>>));
                typeRefs.Add(typeof(Func<string,object,object,object,object,object,object,object,object,object,object,object,object,object,object,object,string>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<string,object,object,object,object,object,object,object,object,object,object,object,object,object,object,object,string>>));
                
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                methodRefs.Add(typeof(List<string>).GetMethod("ToArray")!);
                methodRefs.Add(typeof(List<string>).GetMethod("Add")!);
                """);
        }

        [Fact]
        public void ReturnsTask()
        {
            Verify("""
                using ExcelDna.Integration;
                using System.Threading.Tasks;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class Functions
                    {
                        [ExcelFunction]
                        public static Task<bool> NativeTaskBool()
                        {
                            return Task.FromResult(true);
                        }
                    }
                }
                """, functions: """
                List<Type> typeRefs = new List<Type>();
                ExcelDna.Registration.StaticRegistration.MethodsForRegistration.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Functions).GetMethod("NativeTaskBool")!);
                typeRefs.Add(typeof(Func<System.Threading.Tasks.Task<bool>>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<System.Threading.Tasks.Task<bool>>>));
                
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunTask")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunTaskObject")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunTaskWithCancellation")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunTaskObjectWithCancellation")!.MakeGenericMethod(typeof(bool)));
                """);
        }

        [Fact]
        public void ReturnsObservable()
        {
            Verify("""
                using System;
                using System.Collections.Generic;
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    internal class ObservableString : IObservable<string>
                    {
                        private string s;
                        private List<IObserver<string>> observers;

                        public ObservableString(string s)
                        {
                            this.s = s;
                            observers = new List<IObserver<string>>();
                        }

                        public IDisposable Subscribe(IObserver<string> observer)
                        {
                            observers.Add(observer);
                            observer.OnNext(s);
                            return new ActionDisposable(() => observers.Remove(observer));
                        }

                        private class ActionDisposable : IDisposable
                        {
                            private Action disposeAction;

                            public ActionDisposable(Action disposeAction)
                            {
                                this.disposeAction = disposeAction;
                            }

                            public void Dispose()
                            {
                                disposeAction();
                            }
                        }
                    }

                    public class Functions
                    {
                        [ExcelFunction]
                        public static IObservable<string> NativeStringObservable(string s)
                        {
                            return new ObservableString(s);
                        }
                    }
                }
                """, functions: """
                List<Type> typeRefs = new List<Type>();
                ExcelDna.Registration.StaticRegistration.MethodsForRegistration.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Functions).GetMethod("NativeStringObservable")!);
                typeRefs.Add(typeof(Func<string, System.IObservable<string>>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<string, System.IObservable<string>>>));
                typeRefs.Add(typeof(Func<object, string>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<object, string>>));
                
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("Observe3", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(string)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("ObserveObject", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(string)));
                """);
        }

        [Fact]
        public void AsyncFunction()
        {
            Verify("""
                using ExcelDna.Registration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class Functions
                    {
                        [ExcelAsyncFunction]
                        public static bool NativeAsyncBool()
                        {
                            return true;
                        }
                    }
                }
                """, functions: """
                List<Type> typeRefs = new List<Type>();
                ExcelDna.Registration.StaticRegistration.MethodsForRegistration.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Functions).GetMethod("NativeAsyncBool")!);
                typeRefs.Add(typeof(Func<bool>));
                typeRefs.Add(typeof(System.Linq.Expressions.Expression<Func<bool>>));
                
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunAsTask")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunAsTaskObject")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunAsTaskWithCancellation")!.MakeGenericMethod(typeof(bool)));
                methodRefs.Add(typeof(ExcelDna.Integration.ExcelAsyncUtil).GetMethod("RunAsTaskObjectWithCancellation")!.MakeGenericMethod(typeof(bool)));
                """);
        }

        [Fact]
        public void AssemblyAttributes()
        {
            Verify("""
                [assembly: ExcelDna.Integration.ExcelHandleExternal(typeof(System.Reflection.Assembly))]
                """, assemblyAttributes: """
                ExcelDna.Registration.StaticRegistration.AssemblyAttributes.Add(new ExcelDna.Integration.ExcelHandleExternalAttribute(typeof(System.Reflection.Assembly)));

                """);
        }

        [Fact]
        public void ExcelParameterConversions()
        {
            Verify("""
                using System;
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class Conversions
                    {
                        [ExcelParameterConversion]
                        public static Version ToVersion(string s)
                        {
                            return new Version(s);
                        }
                    }
                }
                """, parameterConversions: """
                ExcelDna.Registration.StaticRegistration.ExcelParameterConversions.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Conversions).GetMethod("ToVersion")!);

                """);
        }

        [Fact]
        public void ExcelReturnConversions()
        {
            Verify("""
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class TestType1
                    {
                        public string Value;

                        public TestType1(string value)
                        {
                            Value = value;
                        }
                    }

                    public class Conversions
                    {
                        [ExcelReturnConversion]
                        public static string FromTestType1(TestType1 value)
                        {
                            return value.Value;
                        }
                    }
                }
                """, returnConversions: """

                ExcelDna.Registration.StaticRegistration.ExcelReturnConversions.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Conversions).GetMethod("FromTestType1")!);
                """);
        }

        [Fact]
        public void ExcelFunctionExecutionHandlerSelectors()
        {
            Verify("""
                using ExcelDna.Integration;
                using ExcelDna.Registration;
                
                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class FunctionLoggingHandler : FunctionExecutionHandler
                    {
                        [ExcelFunctionExecutionHandlerSelector]
                        public static IFunctionExecutionHandler LoggingHandlerSelector(IExcelFunctionInfo functionInfo)
                        {
                            return new FunctionLoggingHandler();
                        }
                    }
                }
                """, executionHandlers: """

                ExcelDna.Registration.StaticRegistration.ExcelFunctionExecutionHandlerSelectors.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.FunctionLoggingHandler).GetMethod("LoggingHandlerSelector")!);
                """);
        }

        [Fact]
        public void PrivateFunctions()
        {
            Verify("""
                using ExcelDna.Integration;

                namespace ExcelDna.AddIn.RuntimeTestsAOT
                {
                    public class ParameterClass { }

                    public class PrivateFunctions
                    {
                        [ExcelFunction]
                        internal static string InternalFunction(ParameterClass c)
                        {
                            return "";
                        }

                        [ExcelFunction]
                        public string InstanceFunction(ParameterClass c)
                        {
                            return "";
                        }
                    }

                    internal class InternalClass2
                    {
                        [ExcelFunction]
                        public static string InternalClass(ParameterClass c)
                        {
                            return "";
                        }
                    }
                }
                """);
        }

        private static void Verify(string sourceCode, string? addins = null, string? functions = null, string? assemblyAttributes = null, string? parameterConversions = null, string? returnConversions = null, string? executionHandlers = null)
        {
            string template = """
        // <auto-generated/>
        using System;
        using System.Collections.Generic;
        using System.Reflection;
        using System.Runtime.CompilerServices;
        using System.Runtime.InteropServices;

        #nullable enable
        
        namespace ExcelDna.SourceGenerator.NativeAOT
        {
            public unsafe class AddInInitialize
            {
                [UnmanagedCallersOnly(EntryPoint = "Initialize", CallConvs = new[] { typeof(CallConvCdecl) })]
                [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL3050:RequiresDynamicCode", Justification = "SourceGenerator preserves types and methods")]
                public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
                {
                    
        
        [BODY]


                    ExcelDna.Registration.StaticRegistration.DirectMarshalTypeAdapter = new DirectMarshalTypeAdapter();
        
                    return ExcelDna.ManagedHost.AddInInitialize.InitializeNativeAOT(xlAddInExportInfoAddress, hModuleXll, pPathXLL, disableAssemblyContextUnload, pTempDirPath, typeof(AddInInitialize).Assembly);
                }
            }
        }
        """;
            functions = functions ?? """
                List<Type> typeRefs = new List<Type>();
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                """;
            string body = $"{functions}\r\n\r\n\r\n{assemblyAttributes}\r\n\r\n{parameterConversions}\r\n{returnConversions}\r\n\r\n{executionHandlers}";
            SourceGeneratorDriver.Verify(sourceCode, template.Replace("[BODY]", body), null);
        }
    }
}

