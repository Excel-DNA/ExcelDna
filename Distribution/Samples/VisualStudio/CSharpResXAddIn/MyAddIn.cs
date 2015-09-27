using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace CSharpAddIn
{

     [ComVisible(true)]     
    public class MyRibbon : ExcelRibbon
    {
        public void OnShowCTP(IRibbonControl control)
        {
          
        }


        public void OnDeleteCTP(IRibbonControl control)
        {
           
        }
    }

    public static class MyAddIn
    {
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelFunction(Description="A bit more than your usual adding function.")]
        public static double AddThemCS(double val1, double val2)
        {
            return val1 + val2 + 25;
        }

        [ExcelFunction(Description = "Get Resources")]
        public static string GetWelcomeText()
        {
            return TestResource.WelcomeText;
        }

        [ExcelFunction(Description = "Get Resources")]
        public static string GetWelcomeTextUS()
        {
            return TestResource.ResourceManager.GetString("WelcomeText", System.Globalization.CultureInfo.GetCultureInfo("en-US"));
        }

    }


}
