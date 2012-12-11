Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Globalization
Imports ExcelDna.Integration

Public Module WebFunctions
    Dim myHttpClient As HttpClient

    Public Function webStringFromURL(url As String) As Object

        If myHttpClient Is Nothing Then
            ServicePointManager.DefaultConnectionLimit = 32
            Dim wrh As New WebRequestHandler
            wrh.AllowPipelining = True
            myHttpClient = New HttpClient(wrh)
        End If

        Return ExcelTaskUtil.RunTask("webStringFromURL", url,
            Function(cancellationToken As CancellationToken) As Task(Of String)
                Return myHttpClient.GetAsync(url, cancellationToken).ContinueWith(
                    Function(t As Task(Of HttpResponseMessage)) As Task(Of String)
                        Dim response = t.Result
                        response.EnsureSuccessStatusCode()
                        Return response.Content.ReadAsStringAsync()
                    End Function).Unwrap()
            End Function)
    End Function

    Public Function webSnippet(url As String, pattern As String) As Object

        Dim webResult = webStringFromURL(url)

        If Equals(webResult, ExcelError.ExcelErrorNA) Then
            Return webResult
        End If

        Dim match As Match = Regex.Match(webResult, pattern, RegexOptions.Singleline)
        If match.Success Then
            Dim result As String = match.Groups(1).Value
            result = result.Replace(Chr(10), " ")
            result = result.Replace(Chr(9), " ")
            Return result
        Else
            Return ExcelError.ExcelErrorValue
        End If
    End Function

    Public Function webYahooPrice(ByVal shareCode As String) As Object

        Dim webResult = webSnippet("http://finance.yahoo.com/q?s=" + shareCode, If(Left(shareCode, 1) = "^", "\", "") + LCase(shareCode) + """>([0-9.,]*)</span>")

        If TypeOf webResult Is ExcelError Then
            Return webResult
        Else
            Return Double.Parse(webResult, NumberStyles.Number, NumberFormatInfo.InvariantInfo)
        End If

    End Function
End Module


Public Class AddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelIntegration.RegisterUnhandledExceptionHandler(
                AddressOf HandleError)
        ExcelAsyncUtil.Initialize()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ExcelAsyncUtil.Uninitialize()
    End Sub

    Private Shared Function HandleError(ex As Object) As Object
        Return "!!! ERROR: " & ex.ToString()
    End Function
End Class
