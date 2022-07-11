Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

Public Module MyCommands
    ' We make a command macro that can be run by:
    ' * Pressing the quick menu under the Add-ins tab
    ' * Pressing the shortcut key Ctrl + Shift + D
    ' * Typing the name into the Alt+F8 Macro dialog (add-in macros won't we shown on this list, though)
    <ExcelCommand(MenuName:="VBfull", MenuText:="Dump Data", ShortCut:="^D")>
    Public Sub DumpData()
        ' We always get the root Application object with a call to ExcelDnaUtil.Application
        ' If we reference both Window Forms, and the Excel interop assemblies, 
        ' we need to be a bit careful about which 'Application' we mean.
        Dim app As Application = CType(ExcelDnaUtil.Application, Application)
        Dim newBook = app.Workbooks.Add()

        ' While .NET arrays and collections are 0-based,
        ' COM collections like Workbook.Sheets are 1-based.
        ' We could also say newBook.Sheets["Sheet1"].
        Dim worksheet As Worksheet = newBook.Sheets(1)
        Dim targetRange As Range = worksheet.Range("A1:C2")

        ' Set values to a range as a 2D object[,] array.
        Dim newValues As Object(,) = New Object(,) {
        {"One", 2, "Three"},
        {True, Date.Now, ""}}
        targetRange.Value = newValues

        ' Apply some formatting, so that the time is displayed correctly
        Dim dateCell As Range = targetRange.Cells(2, 2)
        dateCell.NumberFormat = "hh:mm:ss"
    End Sub
End Module
