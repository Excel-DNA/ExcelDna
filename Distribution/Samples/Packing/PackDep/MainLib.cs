using ExcelDna.Integration;

public class SomeRefClass
{
    [ExcelFunction(Description="This function says Hello", Category="PackDepMyLib")]
    public static string PackDepLib()
    {
        return "Hello from MyLib";
    }

	[ExcelFunction(Description = "This function gets the other message", Category = "PackDepMyLib")]
	public static string PackDepTestOtherLib()
	{
		return SomeDepClass.OtherMessage();
	}

}
