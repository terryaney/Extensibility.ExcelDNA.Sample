﻿using ExcelDna.Integration;

namespace KAT.Extensibility.Excel.AddIn;

public static class MyFunctions
{
	[ExcelFunction( Description = "My first .NET function" )]
	public static string SayHello( string name )
	{
		return "Hello " + name;
	}
}