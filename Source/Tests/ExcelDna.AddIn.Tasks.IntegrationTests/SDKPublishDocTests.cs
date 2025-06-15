﻿using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKPublishDocTests : IntegrationTestBase
    {
        [Test]
        public void Packed()
        {
            const string projectBasePath = @"SDKPublishDoc\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish\";

            Clean(projectOutDir);

            // The first build on a clean project doesn't create .chm at all, but the second run creates .chm and publishes it: 
            for (int i = 0; i < 2; ++i)
                MsBuild(projectBasePath + "SDKPublishDoc.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertNotFound(Path.Combine(publishDir, "SDKPublishDoc-AddIn.chm"));
        }

        [Test]
        public void UnPacked()
        {
            const string projectBasePath = @"SDKPublishDoc\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish\";

            Clean(projectOutDir);

            // The first build on a clean project doesn't create .chm at all, but the second run creates .chm and publishes it: 
            for (int i = 0; i < 2; ++i)
                MsBuild(projectBasePath + "SDKPublishDoc.csproj /p:ExcelDnaUnpack=true /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.chm", "SDKPublishDoc-AddIn.chm");
        }
    }
}
