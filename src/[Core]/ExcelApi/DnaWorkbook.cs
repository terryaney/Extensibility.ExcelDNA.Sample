using ExcelDna.Integration;

#pragma warning disable CA1822 // Doesn't access instance data warning...

namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

class DnaWorkbook
{
	public readonly string Name;
	public readonly string? Version;

	public DnaWorkbook() : this( DnaApplication.ActiveWorkbookName() ) { }
	public DnaWorkbook( string name ) 
	{
		Name = name;
		Version = ReferenceOrNull( "Version" )?.GetValue<string>();
	} 
	public bool IsSaved => (bool)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.IsSaved, Name );

	public CalculationType CalculationType => (CalculationType)(int)(double)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.CalculationMode );
	
	public DnaWorksheet[] Worksheets
	{
		get
		{
			var sheets = (object[,])XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.SheetNames, Name );
			var count = sheets.GetLength( 1 );

			return Enumerable.Range( 0, count )
				.Select( i => new DnaWorksheet( Name, ( (string)sheets[ 0, i ] ).Split( ']' ).Last() ) )
				.ToArray();
		}
	}

	public ExcelReference? ReferenceOrNull( string address )
	{
		var sheetName = (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, Name );
		return XlCall.Excel( XlCall.xlfEvaluate, $"='[{Name}]{sheetName}'!{address}" ) as ExcelReference;
	}
}