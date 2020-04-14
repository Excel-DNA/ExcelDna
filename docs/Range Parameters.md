---
layout: page
title: Range Parameters
---

Parameters with the type of Excel's Range COM object are not directly supported by Excel-DNA.  There is a list of allowed parameter types here: [Reference](Reference)

If you want the function to also accept a sheet reference, your parameter should be of type 'object' and marked with an <ExcelArgument(AllowReference:=true)> attribute. In that case you'll get an object of type ExcelDna.Integration.ExcelReference if the function is called with a sheet reference. 

ExcelReference is not the same as the COM Range type, it is a small wrapper type for the Excel C API reference structure. From the ExcelReference it is possible to get a COM Range -

{% highlight vbnet %}
Imports ExcelDna.Integration.XlCall 
... 
Private Function ReferenceToRange(ByVal xlRef As ExcelReference) As Object 
    Dim cntRef As Long, strText As String, strAddress As String 
    strAddress = Excel(xlfReftext, xlRef.InnerReferences(0), True) 
    For cntRef = 1 To xlRef.InnerReferences.Count - 1 
        strText = Excel(xlfReftext, xlRef.InnerReferences(cntRef), True) 
        strAddress = strAddress & "," & Mid(strText, strText.LastIndexOf("!") + 2) ' +2 because IndexOf starts at 0 
    Next 
    ReferenceToRange = ExcelDnaUtil.Application.Range(strAddress) 
End Function 
{% endhighlight %}


The internal xlfReftext call in ReferenceToRange can only be made from functions that are registered as a macro-sheet functions. For this the exported function will need to be marked as IsMacroType:=True.

So a function that can accept a sheet reference, and process these as a COM Range object, might look like this:

{% highlight vbnet %}
<ExcelFunction(IsMacroType:=True)> _
Public Shared Function GetAddress(<ExcelArgument(AllowReference:=true)> ByVal arg As Object) As String
    Dim range As Object
    If TypeOf arg Is ExcelReference Then
        range = ReferenceToRange(arg)
        Return range.Address(False, False)
    Else
        Return "!!! Not a sheet reference"
    End If
End Function
{% endhighlight %}
