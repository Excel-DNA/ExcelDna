---
layout: page
title: Returning 1-D Arrays
---

By default, ExcelDna marshalls {"object[]()"} (a 1D array) back to Excel as a 1 row, many column result.

This will lead to unexpected results when someone tries to call an array function from a vertical range of cells. For example, if you call a function that returns the array {a,b,c} from three vertical cells, then the cells will show a, a, and a. If you call the function from a horizontal range, then you will see a, b, and c.

If you want an array function that is usable in both horizontal AND vertical mode, then you may want to apply a helper function like this to your 1D result (and return the result PackForCaller from your ExcelFunction):

{% highlight csharp %}
    public static object PackForCaller(object[]() vs) {
      var caller=(ExcelReference)XlCall.Excel(XlCall.xlfCaller);
      var rows=caller.RowLast-caller.RowFirst+1;
      var columns=caller.ColumnLast-caller.ColumnFirst+1;
      if(columns>=rows) {
        return vs;
      }
      var count=vs.Length;
      var vs2=new object[count,1](count,1);
      for(var i=0; i<count; i++) {
        vs2[i, 0](i,-0)=vs[i](i);
      }
      return vs2;
    }
{% endhighlight %}

If the caller is a vertical range, then this will return a 2D array with dimensions of {"[count, 1](count,-1)"}, which will be marshalled back properly to the vertical range of cells.