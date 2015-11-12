using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.Resources;
using System.Globalization;

namespace CSharpAddIn
{

     [ComVisible(true)]     
    public class MyRibbon : ExcelRibbon
    {
        public MyRibbon()
        {
            ResourceManager rm = new ResourceManager(typeof(TestResource));

            var result = new List<CultureInfo>();

            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
            foreach (CultureInfo culture in cultures)
            {
                try
                {
                    ResourceSet rs = rm.GetResourceSet(culture, true, false);
                    // or ResourceSet rs = rm.GetResourceSet(new CultureInfo(culture.TwoLetterISOLanguageName), true, false);
                    string isSupported = (rs == null) ? " is not supported" : " is supported";
                    if (rs != null && culture.Name != "")
                    {
                        result.Add(culture);
                        //break;
                    }

                }
                catch
                {
                }
            }

        }

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
