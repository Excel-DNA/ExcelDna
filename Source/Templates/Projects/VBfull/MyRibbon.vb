Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports VBfull.My.Resources

<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Return RibbonResources.Ribbon
    End Function

    Public Overrides Function LoadImage(imageId As String) As Object
        ' This will return the image resource with the name specified in the image='xxxx' tag
        Return RibbonResources.ResourceManager.GetObject(imageId)
    End Function

    Public Sub OnButtonPressed(control As IRibbonControl)
        System.Windows.Forms.MessageBox.Show("Hello!")
    End Sub

End Class
