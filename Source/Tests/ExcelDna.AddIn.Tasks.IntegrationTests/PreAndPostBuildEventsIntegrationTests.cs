using System;
using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class PreAndPostBuildEventsIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void ExcelDna_build_targets_run_after_the_Pre_build_event_and_before_the_Post_build_event()
        {
            const string projectBasePath = @"PreAndPostBuildEvents\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "PreAndPostBuildEvents.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"),
                buildOutput =>
                {
                    var preBuildEventIndex = buildOutput.IndexOf("Running ExcelDna **Pre-build** event", StringComparison.InvariantCulture);
                    var excelDnaBuildEvent = buildOutput.IndexOf("ExcelDnaBuild:", StringComparison.InvariantCulture);
                    var excelDnaPackEvent = buildOutput.IndexOf("ExcelDnaPack:", StringComparison.InvariantCulture);
                    var postBuildEventIndex = buildOutput.IndexOf("Running ExcelDna **Post-build** event", StringComparison.InvariantCulture);

                    Assert.That(preBuildEventIndex, Is.GreaterThanOrEqualTo(0), "Pre-build event did not execute as expected");
                    Assert.That(excelDnaBuildEvent, Is.GreaterThanOrEqualTo(0), "ExcelDnaBuild target did not execute as expected");
                    Assert.That(excelDnaPackEvent, Is.GreaterThanOrEqualTo(0), "ExcelDnaPack target did not execute as expected");
                    Assert.That(postBuildEventIndex, Is.GreaterThanOrEqualTo(0), "Post-build event did not execute as expected");

                    Assert.That(preBuildEventIndex, Is.LessThan(excelDnaBuildEvent), "Pre-build event executed after ExcelDnaBuild target");
                    Assert.That(excelDnaBuildEvent, Is.LessThan(excelDnaPackEvent), "ExcelDnaBuild target executed after the ExcelDnaPack target");
                    Assert.That(excelDnaPackEvent, Is.LessThan(postBuildEventIndex), "ExcelDnaPack target executed after the Post-build event");
                });
        }
    }
}
