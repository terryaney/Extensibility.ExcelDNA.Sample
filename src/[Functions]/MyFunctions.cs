using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public static class MyFunctions
{
	[ExcelFunction( Description = "My first .NET function" )]
	public static string SayHello( string name )
	{
		if ( name == "terry" )
		{
			throw new ApplicationException( "Use property case" );
		}
		return "Hello " + name;
	}
}