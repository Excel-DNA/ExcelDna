---
layout: post
title: "Resizing Excel UDF result arrays"
date: 2011-01-30 18:27:00 -0000
permalink: /2011/01/30/resizing-excel-udf-result-arrays/
categories: samples, .net, async, excel, exceldna, xll
---
**Update (21 June 2017): The most up-to-date version of the ArrayResizer utility can be found here**: [https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/Samples/ArrayResizer.dna][array-resizer]

**Update: To work correctly under Excel 2000/2002/2003, this sample requires at least version 0.29.0.12 of Excel-DNA**.

A common question on the [Excel-DNA group][excel-dna-group] is how to automatically resize the results of an array formula. The most well-know appearance of this trick is in the Bloomberg add-in.

**WARNING! This is a bad idea**. Excel does not allow you to modify the sheet from within a user-defined function. Doing this breaks Excel's calculation model.

Anyway, here is my attempt at an Excel-DNA add-in that implements this trick. My approach is to run a macro on a separate thread that will check and if required will expand the formula to an array formula of the right size. This way nothing ugly gets done if the array size is already correct - future recalculations will not run the formula array resizer if the size is still correct.

The code below will register a function call `Resize`. You can either call `Resize` from within your function, or enter something like `=Resize(MyFunction(…))` as the cell formula. The code also registers two sample functions, `MakeArray` and `MakeArrayAndResize` to play with, each take the number of rows and columns for the size of the returned array.

To test this:

1. [Get started with Excel-DNA][get-started].
2. Copy the code and xml wrapper into a text file called `Resizer.dna` (the xml wrapper is at the end of this post).
3. Copy the `ExcelDna.xll` in the Excel-DNA distribution to `Resizer.xll` (next to the `Resizer.dna`).
4. File->Open the `Resizer.xll` in Excel and enter something like `=MakeArrayAndResize(5,3)` into a cell.
See how it grows.

In the current version, the formula expansion is destructive, so anything in the way will be erased. One case I don't know how to deal with is when there is an array that would be partially overwritten by the expended function result. In the current version Excel will display an error that says "You cannot change part of an array.", and I replace the formula with a text version of it. I'd love to know how you think we should handle this case.

Any questions or comments (can if anyone can get it to work, or not?) can be directed to the [Excel-DNA Google group][excel-dna-group] or to me directly via e-mail. I'm pretty sure there are a few cases where it will break - please let me know if you run into any problems.

I'll try to gather the comments and suggestions for an improved implementation that might go into the next version of Excel-DNA.

Also, if you have any questions about how the implementation works, I'd be happy to write a follow up post that explains a bit more of what I'm doing. But first, let's try to get it working.

Here's the Resizer add-in code:

{% highlight csharp %}
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration;
 
public static class ResizeTest
{
    public static object MakeArray(int rows, int columns)
    {
        object[,] result = new string[rows, columns];
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                result[i,j] = string.Format("({0},{1})", i, j);
            }
        }
 
        return result;
    }
 
    public static object MakeArrayAndResize(int rows, int columns)
    {
        object result = MakeArray(rows, columns);
        // Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
        return XlCall.Excel(XlCall.xlUDF, "Resize", result);
    }
}
 
public class Resizer
{
    static Queue<ExcelReference> ResizeJobs = new Queue<ExcelReference>();
 
    // This function will run in the UDF context.
    // Needs extra protection to allow multithreaded use.
    public static object Resize(object[,] array)
    {
        ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        if (caller == null)
            return array;
 
        int rows = array.GetLength(0);
        int columns = array.GetLength(1);
 
        if ((caller.RowLast - caller.RowFirst + 1 != rows) ||
            (caller.ColumnLast - caller.ColumnFirst + 1 != columns))
        {
            // Size problem: enqueue job, call async update and return #N/A
            // TODO: Add guard for ever-changing result?
            EnqueueResize(caller, rows, columns);
            AsyncRunMacro("DoResizing");
            return ExcelError.ExcelErrorNA;
        }
 
        // Size is already OK - just return result
        return array;
    }
 
    static void EnqueueResize(ExcelReference caller, int rows, int columns)
    {
        ExcelReference target = new ExcelReference(caller.RowFirst, caller.RowFirst + rows - 1, caller.ColumnFirst, caller.ColumnFirst + columns - 1, caller.SheetId);
        ResizeJobs.Enqueue(target);
    }
 
    public static void DoResizing()
    {
        while (ResizeJobs.Count > 0)
        {
            DoResize(ResizeJobs.Dequeue());
        }
    }
 
    static void DoResize(ExcelReference target)
    {
        try
        {
            // Get the current state for reset later
 
            XlCall.Excel(XlCall.xlcEcho, false);
 
            // Get the formula in the first cell of the target
            string formula = (string)XlCall.Excel(XlCall.xlfGetCell, 41, target);
            ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);
 
            bool isFormulaArray = (bool)XlCall.Excel(XlCall.xlfGetCell, 49, target);
            if (isFormulaArray)
            {
                object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                object oldActiveCell = XlCall.Excel(XlCall.xlfActiveCell);
 
                // Remember old selection and select the first cell of the target
                string firstCellSheet = (string)XlCall.Excel(XlCall.xlSheetNm, firstCell);
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] {firstCellSheet});
                object oldSelectionOnArraySheet = XlCall.Excel(XlCall.xlfSelection);
                XlCall.Excel(XlCall.xlcFormulaGoto, firstCell);
 
                // Extend the selection to the whole array and clear
                XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                ExcelReference oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
 
                oldArray.SetValue(ExcelEmpty.Value);
                XlCall.Excel(XlCall.xlcSelect, oldSelectionOnArraySheet);
                XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
            }
            // Get the formula and convert to R1C1 mode
            bool isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
            string formulaR1C1 = formula;
            if (!isR1C1Mode)
            {
                // Set the formula into the whole target
                formulaR1C1 = (string)XlCall.Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, firstCell);
            }
            // Must be R1C1-style references
            object ignoredResult;
            XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);
            if (retval != XlCall.XlReturn.XlReturnSuccess)
            {
                // TODO: Consider what to do now!?
                // Might have failed due to array in the way.
                firstCell.SetValue("'" + formula);
            }
        }
        finally
        {
            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }
 
    // Most of this from the newsgroup: http://groups.google.com/group/exceldna/browse_thread/thread/a72c9b9f49523fc9/4577cd6840c7f195
    private static readonly TimeSpan BackoffTime = TimeSpan.FromSeconds(1);
    static void AsyncRunMacro(string macroName)
    {
        // Do this on a new thread....
        Thread newThread = new Thread( delegate ()
        {
            while(true)
            {
                try
                {
                    RunMacro(macroName);
                    break;
                }
                catch(COMException cex)
                {
                    if(IsRetry(cex))
                    {
                        Thread.Sleep(BackoffTime);
                        continue;
                    }
                    // TODO: Handle unexpected error
                    return;
                }
                catch(Exception ex)
                {
                    // TODO: Handle unexpected error
                    return;
                }
            }
        });
        newThread.Start();
    }
 
    static void RunMacro(string macroName)
    {
        object xlApp;
        try
        {
            object xlApp = ExcelDnaUtil.Application;
            xlApp.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, xlApp, new object[] {macroName});
        }
        catch (TargetInvocationException tie)
        {
            throw tie.InnerException;
        }
        finally
        {
            Marshal.ReleaseComObject(xlApp);
        }
    }
 
    const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
    const uint VBA_E_IGNORE = 0x800AC472;
    static bool IsRetry(COMException e)
    {
        uint errorCode = (uint)e.ErrorCode;
        switch(errorCode)
        {
            case RPC_E_SERVERCALL_RETRYLATER:
            case VBA_E_IGNORE:
                return true;
            default:
                return false;
        }
    }
}
{% endhighlight %}


You can easily make a test add-in for this by wrapping the code into a .dna file with this around:

{% highlight xml %}
<DnaLibrary Language="CS">
<![CDATA[

    <!--// Paste all of the above code here //-->

]]>
</DnaLibrary>
{% endhighlight %}

[array-resizer]: https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/Samples/ArrayResizer.dna
[excel-dna-group]: http://groups.google.com/group/exceldna
[get-started]: http://exceldna.codeplex.com/wikipage?title=Getting%20Started
