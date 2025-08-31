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
                ExcelDna.Integration.NativeAOT.MethodsForRegistration.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Functions).GetMethod("NativeParamsJoinString")!);
                typeRefs.Add(typeof(Func<string, string[], string>));
                typeRefs.Add(typeof(Func<object, string>));
                typeRefs.Add(typeof(Func<object, string[]>));
                typeRefs.Add(typeof(Func<string,object,object,object,object,object,object,object,object,object,object,object,object,object,object,object,string>));
                
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                methodRefs.Add(typeof(List<string>).GetMethod("ToArray")!);
                """);
        }

        [Fact]
        public void AssemblyAttributes()
        {
            Verify("""
                [assembly: ExcelDna.Integration.ExcelHandleExternal(typeof(System.Reflection.Assembly))]
                """, assemblyAttributes: """
                ExcelDna.Integration.NativeAOT.AssemblyAttributes.Add(new ExcelDna.Integration.ExcelHandleExternalAttribute(typeof(System.Reflection.Assembly)));

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
                ExcelDna.Integration.NativeAOT.ExcelParameterConversions.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Conversions).GetMethod("ToVersion")!);

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

                ExcelDna.Integration.NativeAOT.ExcelReturnConversions.Add(typeof(ExcelDna.AddIn.RuntimeTestsAOT.Conversions).GetMethod("FromTestType1")!);
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

        private static void Verify(string sourceCode, string? addins = null, string? functions = null, string? assemblyAttributes = null, string? parameterConversions = null, string? returnConversions = null)
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
                public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
                {
                    ExcelDna.Integration.NativeAOT.IsActive = true;
        
                    
        
        [BODY]


                    return ExcelDna.ManagedHost.AddInInitialize.InitializeNativeAOT(xlAddInExportInfoAddress, hModuleXll, pPathXLL, disableAssemblyContextUnload, pTempDirPath);
                }
            }
        }
        """;
            functions = functions ?? """
                List<Type> typeRefs = new List<Type>();
                List<MethodInfo> methodRefs = new List<MethodInfo>();
                """;
            string body = $"{functions}\r\n\r\n\r\n{assemblyAttributes}\r\n\r\n{parameterConversions}\r\n{returnConversions}";
            SourceGeneratorDriver.Verify(sourceCode, template.Replace("[BODY]", body));
        }
    }
}
