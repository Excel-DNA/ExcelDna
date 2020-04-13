---
layout: page
title: ExcelDna Performance
--- 

With ExcelDna you can create high-performance UDFs by using a restricted set of types, and taking responsibility for ensuring that exceptions are not leaked. Otherwise, ExcelDna is designed be flexible and make it easy to expose your functions safely - these extensions perform well but are not tuned for high-performance interop.
 
The best performing UDFs will have double or double[,](,) as the parameter and return types, and be marked as IsExceptionSafe=true (and IsThreadSafe=true for Excel 2007), in which case ExcelDna will not wrap your function call in an exception handler. In these cases there is no marshaling code invoked on the managed side, and Excel does the type conversion and error handling as needed. If your function throws an unhandled exception, Excel will crash fatally. For such functions, the only per call overhead is the unmanaged -> managed transition, which is less than 100 CPU instructions. On my fairly slow computer (Excel 2007), Excel is happy to make more than 300 000 calls per second to a simple ExcelDna function (I used CalcCircum from the Xll SDK Example, which multiplies the input number by a constant). With the example C add-in (the no-overhead 'native' case) Excel makes closer to 1 000 000 calls per second to the .xll.
 
Of course, as the code in your UDF becomes significant, this transition overhead of a few microseconds becomes less important,
and the performance of you JITted managed code dominates. (You should expect the performance of managed code to be excellent, but numerical and array-intensive routines might need to be profiled and tuned aggressively, ultimately being maybe 10% slower than C{"++"}.)
 
If an exported function is not marked with IsExceptionSafe=true, then an exception handler is created, and the return value is explicitly marshaled (the return type is effectively object, to allow an error value to be returned). With the additional overhead of the wrapper and marshaling the return value, the simple function recalculates at about 150 000 calls per second.
 
Any other data types force a managed code marshaling for each value. The overhead now includes a function call to the marshaling code, memory allocation and a copy of the data to or from the managed heap. For strings under Excel pre-2007 this includes a text encoding/decoding, in Excel 2007 at least a copy of the string.

## ExcelDna optimisation policy

I will maintain a fast, performance-sensitive path for functions (exception-safe or not)
where all parameter and return types are among the following :
* Double
* Double{"[,](,)"}
* String
* Object
(object return values should be one of
	* Double
	* String
	* Object[,](,)
	* Boolean
	* ExcelError
	* ExcelEmpty
	* ExcelMissing
	* ExcelReference)
 
For other data types (like the other types currently supported) and method extensions, I will focus on the ease with which extensions can be added and the availability and usability of more powerful features. In upcoming versions I want to add more type conversions, and include the ability to add your own type converters and function wrappers.
 
## Computing a million cells

double->double Not ExcelDna - native C .xll (+/- 1s == +/- 1 000 000 /s)
double->double ExcelDna fastest - marked ExceptionSafe (< 3s == 300 000 /s)
double->double ExcelDna - not marked ExceptionSafe (< 6s == 150 000 /s)
object->object  passing in doubles (< 8s == 120 000 /s)
object->string   passing in doubles, returning input.ToString() (< 9s == 110 000 /s)
 
Multi-threaded recalculation

Under Excel 2007 you can also add IsThreadSafe=true to your function. This can give your sheets a dramatic boost that I leave for you to discover ;-)
 
{% highlight csharp %}
{"[ExcelFunction(IsExceptionSafe=true, IsThreadSafe=true)](ExcelFunction(IsExceptionSafe=true,-IsThreadSafe=true))"}
public static double CalcCircumDna(double val)
{
    return val * 6.283185308;
}
{% endhighlight %}