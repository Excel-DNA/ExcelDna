using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace ExcelDna.SourceGenerator.NativeAOT.Tests
{
    internal class SourceGeneratorDriver
    {
        public static void Verify(string sourceCode, string expected)
        {
            Compilation inputCompilation = CSharpCompilation.Create("compilation",
                [CSharpSyntaxTree.ParseText(sourceCode)],
                [
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(typeof(ExcelDna.Integration.NativeAOT).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(ExcelDna.ManagedHost.AddInInitialize).Assembly.Location),
                ],
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary).WithAllowUnsafe(true));
            NativeAOT.Generator generator = new NativeAOT.Generator();
            GeneratorDriver driver = CSharpGeneratorDriver.Create(generator).RunGeneratorsAndUpdateCompilation(inputCompilation, out var outputCompilation, out var diagnostics);

            Assert.Empty(diagnostics);
            Assert.True(outputCompilation.SyntaxTrees.Count() == 2);
            Assert.Empty(outputCompilation.GetDiagnostics());

            GeneratorDriverRunResult runResult = driver.GetRunResult();
            Assert.True(runResult.GeneratedTrees.Length == 1);
            Assert.Empty(runResult.Diagnostics);

            GeneratorRunResult generatorResult = runResult.Results[0];
            Assert.True(generatorResult.Generator == generator);
            Assert.Empty(generatorResult.Diagnostics);
            Assert.True(generatorResult.GeneratedSources.Length == 1);
            Assert.True(generatorResult.Exception is null);
            Assert.Equal(expected, generatorResult.GeneratedSources[0].SourceText.ToString());
        }
    }
}
