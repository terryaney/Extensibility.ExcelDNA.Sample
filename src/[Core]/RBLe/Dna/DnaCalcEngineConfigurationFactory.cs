
using ExcelDna.Integration;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;

// TODO: Could make ConfigurationFactory better by eliminating the need for End() to do a 'select' every time for the sheet...
// 1. Easiest would be to read entire 2 rows (table names/col names) and delete with object[,] data.
// 2. Going to need EndDown() eventually for data reading, so maybe figure out a way to activate a sheet in a 'ReadSheet/LoadSheet' type method before
//		doing all the End() calls and then I at least eliminate the 'read sheet name' and 'select sheet' api calls.
class DnaCalcEngineConfigurationFactory : CalcEngineConfigurationFactory<DnaWorkbook, DnaWorksheet, ExcelReference>
{
	private readonly string fileName;
	private readonly DnaWorkbook workbook;
	public DnaCalcEngineConfigurationFactory( string fileName )
	{
		this.fileName = fileName;
		workbook = new DnaWorkbook( fileName );
	}

	protected override string FileName => fileName;
	protected override string Version => workbook.Version ?? "Unknown";
	protected override DnaWorksheet[] Worksheets => new DnaWorkbook( fileName ).Worksheets;
	protected override DnaWorksheet GetSheet( ExcelReference range ) => new( workbook.Name, range.SheetName() );
	protected override string GetName( DnaWorksheet sheet ) => sheet.Name;
	protected override string? RangeTextOrNull( DnaWorksheet sheet, string name ) => sheet.RangeOrNull( name )?.GetValue<string>();
	protected override ExcelReference GetRange( string name ) => workbook.RangeOrNull( name )!;
	protected override ExcelReference GetRange( DnaWorksheet sheet, string nameOrAddress ) => sheet.RangeOrNull( nameOrAddress )!;
	protected override ExcelReference Offset( ExcelReference range, int rowOffset, int columnOffset ) => range.Offset( rowOffset, columnOffset );
	protected override ExcelReference EndRight( ExcelReference range ) => range.End( DirectionType.ToRight );
	protected override bool RangeExists( string name ) => workbook.RangeOrNull( name ) != null;
	protected override bool RangeExists( DnaWorksheet sheet, string name ) => sheet.RangeOrNull( name ) != null;
	protected override string GetAddress( ExcelReference range ) => range.GetAddress().Split( '!' ).Last();
	protected override string GetText( ExcelReference range ) => range.GetValue<string>() ?? "";
}