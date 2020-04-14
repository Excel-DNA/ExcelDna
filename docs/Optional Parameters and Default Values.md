---
layout: page
title: Optional Parameters and Default Values
---

**Update:** There is now a project to develop rich parameter support using a custom registration processing pipeline. This includes support for Optional and Default values. See [https://github.com/Excel-DNA/CustomRegistration](https://github.com/Excel-DNA/CustomRegistration). I leave the documentation below as applying to the core Excel-DNA library. _



There is currently no special support built into Excel-DNA for optional parameters or default values. You can implement these in your add-in by changing your parameter type to be 'object' and then dealing with the different options explicitly. 

As an example, the code below shows how you could create a helper class to deal with the passed parameters, making the handling in you functions as easy as possible.

{% highlight csharp %}
    using System;
    using ExcelDna.Integration;

    // These are some functions that implement an optional parameter with some default value.
    public class MyFunctions
    {
        public static double TestDefault(double x, object yArg)
        {
            double y = Optional.Check(yArg, 17.0);
            
            return x + y;
        }
        
        public static string TestHello(object nameArg)
        {
            string name = Optional.Check(nameArg, " Unknown person!?");
            
            return "Hello " + name;
        }
        
        public static string TestSpecialDate(object dateArg)
        {
            DateTime date = Optional.Check(dateArg, DateTime.Now);
            
            return "The special date is: " + date.ToString("dd MMMM yyyy");
        }
    }
    
    // Here is the helper class - add to it or change as you require
    internal static class Optional
    {
        internal static string Check(object arg, string defaultValue)
        {
            if (arg is string)
                return (string)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                return arg.ToString();  // Or whatever you want to do here....

            // Perhaps check for other types and do whatever you think is right ....
            //else if (arg is double)
            //    return "Double: " + (double)arg;
            //else if (arg is bool)
            //    return "Boolean: " + (bool)arg;
            //else if (arg is ExcelError)
            //    return "ExcelError: " + arg.ToString();
            //else if (arg is object[,](,))
            //    // The object array returned here may contain a mixture of types,
            //    // reflecting the different cell contents.
            //    return string.Format("Array[{0},{1}]({0},{1})", 
            //      ((object[,](,)(,))arg).GetLength(0), ((object[,](,)(,))arg).GetLength(1));
            //else if (arg is ExcelEmpty)
            //    return "<<Empty>>"; // Would have been null
            //else if (arg is ExcelReference)
            //  // Calling xlfRefText here requires IsMacroType=true for this function.
			//				return "Reference: " + 
            //                     XlCall.Excel(XlCall.xlfReftext, arg, true);
			//			else
			//				return "!? Unheard Of ?!";
        }        

        internal static double Check(object arg, double defaultValue)
        {
            if (arg is double)
                return (double)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
                
        }
        
        // This one is more tricky - we have to do the double->Date conversions ourselves
        internal static DateTime Check(object arg, DateTime defaultValue)
        {
            if (arg is double)
                return DateTime.FromOADate((double)arg);    // Here is the conversion
            else if (arg is string)
                return DateTime.Parse((string)arg);
            else if (arg is ExcelMissing)
                return defaultValue;
                
            else 
                throw new ArgumentException();  // Or defaultValue or whatever
        }
    }
{% endhighlight %}
