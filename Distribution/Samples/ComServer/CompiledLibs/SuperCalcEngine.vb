Imports System.Runtime.InteropServices

<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class SuperCalcEngine

    Public Function MultiplyThem(v1 As Double, v2 As Double) As Double
        Return v1 * v2
    End Function

    Public ReadOnly Property Version As String
        Get
            Return "SuperCalcEngine Version 0.1"
        End Get
    End Property
End Class
