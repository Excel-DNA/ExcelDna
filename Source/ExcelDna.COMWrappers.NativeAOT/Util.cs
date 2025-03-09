namespace ExcelDna.COMWrappers.NativeAOT
{
    public static class Util
    {
        public static object? GetExcelApplication()
        {
            foreach (var w in OpenWindowGetter.GetOpenWindows())
            {
                if (w.Value.EndsWith("Excel"))
                    return GetExcelObject(w.Key);
                //Console.WriteLine(w.Value);

            }

            return null;
        }

        static object? GetExcelObject(System.IntPtr w)
        {
            return Excel.GetApplicationFromWindow(w);
        }
    }
}
