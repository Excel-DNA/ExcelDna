Imports System.IO
Imports System.Windows.Forms
Imports ExcelDna.Integration

Public Class ExampleAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Dim thisAddInName = Path.GetFileName(XlCall.Excel(XlCall.xlGetName))
        Dim message = String.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName)

        MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose

    End Sub

End Class
