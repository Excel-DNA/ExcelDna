using ExcelDna.Integration;

namespace SDKPackDoc
{
    public class Text
    {
        [ExcelFunction(Name = "Text.ConcatThem",
                        Description = "concatenates two strings",
                        HelpTopic = "SDKPackDoc-AddIn.chm!1002")]
        public static object ConcatThem(
            [ExcelArgument(Description = "the first string")] object a,
            [ExcelArgument(Description = "the second string")] object b)
        {
            return string.Concat(a.ToString(), b.ToString());
        }
    }
}
