---
layout: page
title: Async macro example
---

We define a function that schedules a macro call to update the format of the calling range.

{% highlight csharp %}
    public static DateTime asyncFormatCaller()
    {
        object caller = XlCall.Excel(XlCall.xlfCaller);
        if (caller is ExcelReference)
        {
            ExcelAsyncUtil.QueueAsMacro(
                delegate
                {
                      // Set the desired selection, then apply formatting
                      using (new ExcelSelectionHelper((ExcelReference)caller))
                      {
                          XlCall.Excel(XlCall.xlcFormatNumber, "h:mm:ss");
                      }
                }
            });
        }
        return DateTime.Now;
    }
{% endhighlight %}

Above we use the following helper class to keep track of the current selection in a macro, so that the selection is correctly restored after the macro has completed.

{% highlight csharp %}
    // Helper class to deal with Excel selections in 'using' style
    public class ExcelSelectionHelper : XlCall, IDisposable
    {
        object oldScreenUpdating;
        object oldSelectionOnActiveSheet;
        object oldActiveCellOnActiveSheet;

        object oldSelectionOnRefSheet;
        object oldActiveCellOnRefSheet;

        public ExcelSelectionHelper(ExcelReference refToSelect)
        {
            oldScreenUpdating = Excel(xlfGetWorkspace, 40);
            Excel(xlcEcho, false);

            // Remember old selection state on the active sheet
            oldSelectionOnActiveSheet = Excel(xlfSelection);
            oldActiveCellOnActiveSheet = Excel(xlfActiveCell);

            // Switch to the sheet we want to select
            string refSheet = (string)Excel(xlSheetNm, refToSelect);
            Excel(xlcWorkbookSelect, new object[]() { refSheet });

            // record selection and active cell on the sheet we want to select
            oldSelectionOnRefSheet = Excel(xlfSelection);
            oldActiveCellOnRefSheet = Excel(xlfActiveCell);

            // make the selection
            Excel(xlcFormulaGoto, refToSelect);
        }

        public void Dispose()
        {
            Excel(xlcSelect, oldSelectionOnRefSheet, oldActiveCellOnRefSheet);

            string oldActiveSheet = (string)Excel(xlSheetNm, oldSelectionOnActiveSheet);
            Excel(xlcWorkbookSelect, new object[]() { oldActiveSheet });

            Excel(xlcSelect, oldSelectionOnActiveSheet, oldActiveCellOnActiveSheet);

            Excel(xlcEcho, oldScreenUpdating);
        }
    }
{% endhighlight %}

An improved function could first check whether the format of the caller needs to be updated before scheduling the macro call.