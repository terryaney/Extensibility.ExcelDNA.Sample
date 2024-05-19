using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaBlockchain
{
	[ExcelFunction( Category = "Crypto", Description = "Hashes a value using the specified hash type." )]
	public static string BTRHashString(
		[ExcelArgument( "The value to hash." )]
		string value,
		[ExcelArgument( "Optional.  The hash implementation to use (i.e. Sha256)." )]
		object? hashType = null 
	)
	{
		var hashTypeArg = hashType.Check( nameof( hashType ), "Sha256" );
		return Blockchain.HashValue( value, hashTypeArg );
	}

	[ExcelFunction( Category = "Crypto", Description = "Hashes a hex string using the specified hash type." )]
	public static string BTRHashByteString(
		[ExcelArgument( "The value to hash." )]
		string byteString,
		[ExcelArgument( "Optional.  The hash implementation to use (i.e. Sha256)." )]
		object? hashType = null 
	)
	{
		var hashTypeArg = hashType.Check( nameof( hashType ), "Sha256" );
		return Blockchain.HashByteString( byteString, hashTypeArg );
	}

	[ExcelFunction( Category = "Crypto", Description = "Hashes a file using the specified hash type." )]
	public static string BTRHashFile(
		[ExcelArgument( "The file name to hash." )]
		string fileName,
		[ExcelArgument( "Optional.  The hash implementation to use (i.e. Sha256)." )]
		object? hashType = null 
	)
	{
		var hashTypeArg = hashType.Check( nameof( hashType ), "Sha256" );
		return Blockchain.HashFile( fileName, hashTypeArg );
	}

	[ExcelFunction( Category = "Crypto", Description = "Compute the little endian value for a hex string." )]
	public static string BTRLittleEndian(
		[ExcelArgument( "The value to compute." )]
		string hexValue 
	) => Blockchain.LittleEndian( hexValue );

	[ExcelFunction( Category = "Crypto", Description = "Convert a date to Unix time." )]
	public static long BTRToUnixTime(
		[ExcelArgument( "The date to convert." )]
		DateTime date 
	) => Blockchain.ToUnixTime( date );	
}