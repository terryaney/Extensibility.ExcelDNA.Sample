
using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;

// TODO: Could make ConfigurationFactory better by eliminating the need for End() to do a 'select' every time for the sheet...
// 1. Easiest would be to read entire 2 rows (table names/col names) and delete with object[,] data.
// 2. Going to need EndDown() eventually for data reading, so maybe figure out a way to activate a sheet in a 'ReadSheet/LoadSheet' type method before
//		doing all the End() calls and then I at least eliminate the 'read sheet name' and 'select sheet' api calls.
public class DnaCalcEngineConfigurationFactory : CalcEngineConfigurationFactory<DnaWorkbook, DnaWorksheet, ExcelReference>
{
	private readonly string fileName;
	public readonly List<string> ApiCalls = new();

	public DnaCalcEngineConfigurationFactory( string fileName ) => this.fileName = fileName;

	protected override string FileName => fileName;
	protected override string Version => GetTextOrNull( TryGetWorkbookReference( null, "Version" )! ) ?? "Unknown";

	protected override DnaWorksheet[] Worksheets
	{
		get
		{
			var sheets = (object[,])XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.SheetNames, fileName );
			var count = sheets.GetLength( 1 );

			return Enumerable.Range( 0, count )
				.Select( i => new DnaWorksheet( ( (string)sheets[ 0, i ] ).Split( ']' ).Last() ) )
				.ToArray();
		}
	}
	protected override DnaWorksheet GetSheet( ExcelReference range )
	{
		var sheetName = (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.SheetRef, range );
		return new( sheetName );
	}

	protected override string GetName( DnaWorksheet sheet ) => sheet.Name;
	protected override string? RangeTextOrNull( string name ) => GetTextOrNull( TryGetWorkbookReference( null, name ) );
	protected override string? RangeTextOrNull( DnaWorksheet sheet, string name ) => GetTextOrNull( TryGetWorkbookReference( null, $"'{sheet.Name}'!{name}" ) );
	protected override ExcelReference GetRange( string name ) => TryGetWorkbookReference( null, name )!;
	protected override ExcelReference GetRange( DnaWorksheet sheet, string name ) => TryGetWorkbookReference( null, $"'{sheet.Name}'!{name}" )!;
	protected override ExcelReference Offset( ExcelReference range, int rowOffset, int columnOffset ) =>
		new (
			range.RowFirst + rowOffset,
			range.RowLast + rowOffset,
			range.ColumnFirst + columnOffset,
			range.ColumnLast + columnOffset,
			range.SheetId
		);

	protected override ExcelReference EndRight( ExcelReference range ) => range.End( DirectionType.ToRight );
	protected override bool RangeExists( string name ) => TryGetWorkbookReference( null, name ) != null;
	protected override bool RangeExists( DnaWorksheet sheet, string name ) => TryGetWorkbookReference( null, $"'{sheet.Name}'!{name}" ) != null;
	protected override string GetAddress( ExcelReference range )
	{
		var address = (string)XlCall.Excel( XlCall.xlfReftext, range, true /* true - A1, false - R1C1 */ );
		return address.Split( '!' ).Last();
	}
	protected override string GetText( ExcelReference range ) => GetTextOrNull( range ) ?? "";

	private static string? GetTextOrNull( ExcelReference? range )
	{
		if ( range == null ) return null;

		var value = range.GetValue();
		return !value.Equals( ExcelEmpty.Value ) 
			? (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Text, range ) 
			: null;
	}
	protected override string GetFormula( ExcelReference range ) 
	{
		var f = XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Formula, range );
		var formula = f is ExcelError check && check == ExcelError.ExcelErrorValue ? null : (string)f;
		return !string.IsNullOrEmpty( formula ) ? formula : "";
	}

	ExcelReference? TryGetWorkbookReference( string? workbookName, string address )
	{
		var excelAddress = GetExcelAddress( address );

		var hasSheet = !string.IsNullOrEmpty( excelAddress.Sheet ) && excelAddress.Sheet != fileName; /* global range if so */
		if ( !hasSheet )
		{
			// Assuming if no sheet provided, they mean the 'active' sheet
			excelAddress = new()
			{
				Address = excelAddress.Address,
				Workbook = excelAddress.Workbook,
				Sheet = !string.IsNullOrEmpty( workbookName )
					? (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, workbookName )
					: ( (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveSheet ) ).Split( ']' ).Last()
			};
		}

		var addressWorkbook = excelAddress.Workbook ?? workbookName;
		var workbookNameSyntax = !string.IsNullOrEmpty( addressWorkbook ) ? $"[{addressWorkbook}]" : null;

		var fullAddress = $"='{workbookNameSyntax}{excelAddress.Sheet}'!{excelAddress.Address}";
		var reference = XlCall.Excel( XlCall.xlfEvaluate, fullAddress ) as ExcelReference;

		return reference;
	}

	private static ExcelAddress GetExcelAddress( string address )
	{
		// Start: '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'!$A$3

		var addressParts = address.Split( '!' );
		// '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'
		// $A$3

		if ( addressParts.Length == 1 )
		{
			return new ExcelAddress { Address = addressParts[ 0 ] };
		}

		var sheetName = addressParts[ 0 ].StartsWith( "'" )
			? addressParts[ 0 ][ 1..^1 ]
			: addressParts[ 0 ];
		// Current: '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'

		var rangeAddress = addressParts[ 1 ];
		// Current: $A$3

		addressParts = sheetName.Split( ']' );
		// [Buck_MurphyOil_SE debug macro.xls
		// RBLMacro

		sheetName = addressParts.Last();

		if ( sheetName.EndsWith( ".csv", StringComparison.InvariantCultureIgnoreCase ) )
		{
			// Seems a csv file with one tab (the normal format) comes through as
			// address -> 'sheetname.csv'!address
			sheetName = Path.GetFileNameWithoutExtension( sheetName );
		}
		return new ExcelAddress
		{
			Workbook = addressParts.Length == 2 ? addressParts[ 0 ][ 1.. ] : null,
			Sheet = sheetName,
			Address = rangeAddress
		};
	}
}