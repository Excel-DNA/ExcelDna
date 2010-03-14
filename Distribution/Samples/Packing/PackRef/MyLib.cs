using ExcelDna.Integration;

public class SomeClass
{
    [ExcelFunction(Description="This function says Hello", Category="PackRefMyLib")]
    public static string PackRefMyLib()
    {
        return "Hello from PackRefMyLib";
    }

	[ExcelFunction(Description = "This function also computes the sin of a number", Category = "PackRefMyLib")]
	public static double PackRefAlsoSin(double v1)
	{
		return (double)XlCall.Excel(XlCall.xlfSin, v1);
	}
}
