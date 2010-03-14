using ExcelDna.Integration;

public class SomeClass
{
    [ExcelFunction(Description="This function says Hello", Category="PackMyLib")]
    public static string PackMyLib()
    {
        return "Hello from PackMyLib";
    }

	[ExcelFunction(Description = "This function also computes the sin of a number", Category = "PackMyLib")]
	public static double PackAlsoSin(double v1)
	{
		return (double)XlCall.Excel(XlCall.xlfSin, v1);
	}
}
