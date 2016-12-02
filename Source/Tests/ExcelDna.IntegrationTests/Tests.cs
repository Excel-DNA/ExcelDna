using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using NUnit.Framework;

namespace ExcelDna.Integration.Tests
{
    [TestFixture, Explicit] // the Explicit attribute causes the test to be skipped during aytomated runs. This is because XL is not suitable for non-interactive sessions
    public class ExcelBasedTests
    {
        #region Member variables
        private Process m_xlProcess;
        private Application m_xlApp;
        private Workbooks m_xlWorkbooks;
        private Workbook m_xlWorkbook;
        #endregion

        private string PrepareXlDna(string outDir)
        {
            bool is64Bit = (Marshal.SizeOf(m_xlApp.HinstancePtr) == 8);
            var xldna = Path.Combine(outDir, "ExcelDna") + (is64Bit ? "64" : "") + ".xll";
            var xldnaToLoad = Path.Combine(outDir, "ExcelDna.IntegrationTests-AddIn.xll");
            var mtx = new Mutex(false, "xldnaTests");
            lock (mtx)
                File.Copy(xldna, xldnaToLoad, true);
            Assert.IsTrue(File.Exists(xldnaToLoad));
            return xldnaToLoad;
        }

        string hashErrorToString(uint err)
        {
            string result;
            switch (err)
            {
                case 0x800a07d0:
                    result = "#NULL!";
                    break;
                case 0x800a07d7:
                    result = "#DIV/0!";
                    break;
                case 0x800a07df:
                    result = "#VALUE!";
                    break;
                case 0x800a07e7:
                    result = "#REF!";
                    break;
                case 0x800a07ed:
                    result = "#NAME?";
                    break;
                case 0x800a07f4:
                    result = "#NUM!";
                    break;
                case 0x800a07fa:
                    result = "#N/A";
                    break;
                default:
                    result = err.ToString();
                    break;
            }
            return result;
        }

        // Simply add a test workbook to the testWorkbooks directory and a corresponding TestCase attribute below.
        // The MSBuild project will copy the testWorkbooks directory to the output directory of the project.
        // Then this data-driven TestCase will do a CalculateFullRebuild on each workbook and read a single cell named RESULT
        // which must evaluate to a boolean TRUE value.
        // If everything works, then the copy of the workbook gets deleted. Otherwise it left there for post-mortem.
        [TestCase("basic.xlsm")]
        [Timeout(1800000)]
        public void ExcelTest(string workbook)
        {
            var outDir = Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);

            m_xlApp = new Application();
            Assert.IsNotNull(m_xlApp, "Could not create an Excel Application object");

            m_xlWorkbooks = m_xlApp.Workbooks;
            var testMacrosPath = Path.Combine(outDir, "testMacros.xlsm");
            var testMacros = m_xlWorkbooks.Open(testMacrosPath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var pidAsObject = m_xlApp.Run(testMacros.Name + "!getpid");
            m_xlProcess = Process.GetProcessById(pidAsObject);
            var xldnaToLoad = PrepareXlDna(outDir);

            workbook = Path.Combine(outDir, "testWorkbooks", workbook);
            var workbookName = Path.GetFileName(workbook);

            Assert.IsTrue(File.Exists(workbook));
            Assert.IsTrue(m_xlApp.RegisterXLL(xldnaToLoad));

            m_xlApp.DisplayAlerts = false;

            m_xlWorkbook = m_xlWorkbooks.Open(workbook, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var workDir = Path.Combine(outDir, "evidence");
            if (!Directory.Exists(workDir))
            {
                try { Directory.CreateDirectory(workDir); } catch (Exception) { }; // more than 1 thread might get here, but only 1 will succeed
            }
            Assert.IsTrue(Directory.Exists(workDir));
            m_xlApp.Run(testMacros.Name + "!setProcessCurrentDirectory", workDir);

            m_xlApp.CalculateBeforeSave = false;

            bool vbaResult = false;
            try
            {
                vbaResult = m_xlApp.Run(m_xlWorkbook.Name + "!VBATest");
            }
            catch (COMException x)
            {
                // This HR means "Programmatic access to Visual Basic is not trusted"
                // It is returned also in case the macro does not exist
                // If we get anything other than this type of error, we rethrow.
                // Otherwise there's nothing to test and we carry on.
                if ((uint)x.HResult != 0x800a03ec)
                    throw;
                vbaResult = true;
            }
            Assert.IsTrue(vbaResult);

            Range resultRange = m_xlApp.Range["RESULT"];
            Assert.AreEqual(1, resultRange.Count);
            resultRange.Dirty();
            m_xlApp.CalculateFullRebuild();
            dynamic testResult = resultRange.Value;
            if (typeof(int) == testResult.GetType())
                Assert.Fail("RESULT={0}", hashErrorToString((uint)testResult));
            Assert.AreEqual(typeof(bool), testResult.GetType(), "RESULT={0}", testResult.ToString());
            Assert.IsTrue((bool)testResult);

            // the test has passed. we don't need to leave the xlsm around
            m_xlWorkbook.Close(false);
            Marshal.FinalReleaseComObject(m_xlWorkbook);
            m_xlWorkbook = null;
            File.Delete(workbook);
        }

        [TearDown]
        public void TearDown()
        {
            if (m_xlWorkbook != null)
            {
                m_xlWorkbook.Save();
                m_xlWorkbook.Close(false);
                Marshal.FinalReleaseComObject(m_xlWorkbook);
            }
            if (m_xlWorkbooks != null)
            {
                Marshal.FinalReleaseComObject(m_xlWorkbooks);
            }
            if (m_xlApp != null)
            {
                m_xlApp.Quit();
                Marshal.FinalReleaseComObject(m_xlApp);
            }
            if (!m_xlProcess.WaitForExit(500))
                m_xlProcess.Kill();
        }
    }
}
