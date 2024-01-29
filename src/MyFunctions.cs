using ExcelDna.Integration;
using KAT.Camelot.RBLe.Lookups;

namespace KAT.Extensibility.Excel.AddIn;

public static class MyFunctions
{
	[ExcelFunction( Description = "My first .NET function" )]
	public static string SayHello( string name )
	{
		return "Hello " + name + ", there are " + LookupTables.MortalityTableNames().Length + " mortality tables available.";
	}
}