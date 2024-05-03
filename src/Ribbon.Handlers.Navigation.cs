using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Navigation_NavigateToTable( IRibbonControl _ )
	{
		var tables = new List<NavigationTable>();

		var activeSheet = application.ActiveWorksheet();

		if ( activeSheet.Names.Cast<MSExcel.Name>().Any( n => n.Name == Constants.CalcEngines.RangeNames.StartTables || n.Name.EndsWith( $"!{Constants.CalcEngines.RangeNames.StartTables}" ) ) )
		{
			tables.AddRange( GetCalcEngineTables( activeSheet ) );
		}
		else if ( activeSheet.Name == Constants.SpecSheet.TabNames.HistoricalData )
		{
			tables.AddRange( GetSpecSheetHistoryTables( activeSheet ) );
		}
		else if ( activeSheet.Name == Constants.SpecSheet.TabNames.CodeTables )
		{
			tables.AddRange( 
				GetCodeTables( 
					activeSheet.Range[ "A5" ],
					start => start.Offset[ 0, 1 ].End[ MSExcel.XlDirection.xlToRight ]
				) 
			);
		}
		else if ( WorkbookState.SheetState.IsGlobalTableSheet )
		{
			tables.AddRange( 
				GetCodeTables( 
					activeSheet.Range[ activeSheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.TableStartAddress ) ],
					start => start.End[ MSExcel.XlDirection.xlToRight ]
				) 
			);
		}

		using var navigateToTable = new NavigateToTable( tables, GetWindowConfiguration( nameof( NavigateToTable ) ) );

		var info = navigateToTable.GetInfo();

		if ( info == null )
		{
			return;
		}
		activeSheet.Range[ info.Target ].Select();

		SaveWindowConfiguration( nameof( NavigateToTable ), info.WindowConfiguration );
	}

	private static IEnumerable<NavigationTable> GetCodeTables( MSExcel.Range start, Func<MSExcel.Range, MSExcel.Range> getNextTable )
	{
		while ( string.Compare( start.GetText(), "Table", true ) == 0 )
		{
			start = start.Offset[ 0, 1 ];

			var description = start.Offset[ -1, 0 ].GetText();

			yield return new()
			{
				Name = start.GetText(),
				Address = start.Address,
				Description = GetTableDescription( description ) ?? description,
				Columns = GetColumns( start.Offset[ 1, -1 ] )
			};

			start = getNextTable( start );
		}
	}

	private static IEnumerable<NavigationTable> GetSpecSheetHistoryTables( MSExcel.Worksheet activeSheet )
	{
		var start = activeSheet.Range[ "A5" ];

		while ( string.Compare( start.GetText(), "Data Type:", true ) == 0 )
		{
			var description =
				string.Join( " - ",
					new[] {
						start.Offset[ 2, 1 ].GetText(),
						start.Offset[ 2, 2 ].GetText()
					}.Where( s => !string.IsNullOrEmpty( s ) )
				);

			yield return new()
			{
				Name = start.Offset[ 0, 1 ].GetText(),
				Address = start.Offset[ 4, 0 ].Address,
				Description = GetTableDescription( description ) ?? description,
				Columns = GetColumns( start.Offset[ 4, 0 ], false )
			};

			start = start.End[ MSExcel.XlDirection.xlDown ].Offset[ 2, 0 ];
		}
	}

	private static IEnumerable<NavigationTable> GetCalcEngineTables( MSExcel.Worksheet activeSheet )
	{
		var start = activeSheet.Range[ "StartTables" ].Offset[ 1, 0 ];
		string? name;

		while ( !string.IsNullOrEmpty( name = start.GetText().Split( '/' )[ 0 ] ) )
		{
			string? suffix = null;
			if ( name.StartsWith( "<<" ) )
			{
				name = name[ 2..^2 ];
				suffix = " - Lookup Table";
			}
			else if ( name.StartsWith( "<" ) )
			{
				name = name[ 1..^1 ];
				suffix = " - Data Table";
			}

			var descriptionTest = start.Offset[ -1, 0 ].GetText();
			var description = descriptionTest.StartsWith( "//" ) ? descriptionTest : start.Offset[ -2, 0 ].GetText();

			yield return new()
			{
				Name = name.Split( '/' )[ 0 ] + suffix,
				Address = start.Address,
				Description = GetTableDescription( description ),
				Columns = GetColumns( start.Offset[ 1, 0 ] )
			};

			start = start.Offset[ 1, 0 ].End[ MSExcel.XlDirection.xlToRight ].Offset[ -1, 2 ];
		}
	}

	private static NavigationTarget[] GetColumns( MSExcel.Range start, bool isHorizontal = true )
	{
		var columns = new List<NavigationTarget>();
		var colStart = start;
		string? colName;

		while ( !string.IsNullOrEmpty( colName = colStart.GetText() ) )
		{
			columns.Add( new NavigationTarget
			{
				Name = colName,
				Address = colStart.Address
			} );

			colStart = isHorizontal
				? colStart.Offset[ 0, 1 ]
				: colStart.Offset[ 1, 0 ];
		}

		return columns.ToArray();
	}

	public void Navigation_GoToInputs( IRibbonControl _ ) => GotoInputNamedRange( "StartData" );
	public void Navigation_GoToInputData( IRibbonControl _ ) => GotoInputNamedRange( "xDSDataFields" );
	public void Navigation_GoToCalculationInputs( IRibbonControl _ ) => GotoInputNamedRange( "CalculationInputs" );
	public void Navigation_GoToFrameworkInputs( IRibbonControl _ ) => GotoInputNamedRange( "FrameworkInputs" );
	public void Navigation_GoToInputTables( IRibbonControl _ ) => GotoInputNamedRange( "StartTables" );
	public void Navigation_BackToRBLeMacro( IRibbonControl _ ) => GotoNamedRange( "RBLeMacro", false );

	public void Navigation_GoToBTRCellAddress( IRibbonControl _ )
	{
		var formula = "[Unavailable]";
		var address = "[Unavailable]";
		try
		{
			var selection = application.Selection as MSExcel.Range;
			formula = selection!.Formula as string;
			
			if ( formula!.Contains( "BTRCellAddress" ) )
			{
				address = selection.Text as string;
				var range = address!.GetReference().GetRange();
				range.Worksheet.Activate();
				range.Activate();
			}
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to go to BTRCellAddress selected.\r\n\r\nFormula: {formula}\r\nAddress: {address}", ex );
		}
	}

	private void GotoInputNamedRange( string name )
	{
		var inputTab =
			application
				.ActiveWorkbook
				.Sheets
				.Cast<MSExcel.Worksheet>()
				.FirstOrDefault( w => w.RangeOrNull<string>( "SheetType" ) == "Input" );

		if ( inputTab != null )
		{
			GotoNamedRange( $"{inputTab.Name}!{name}", true );
		}
	}

	private void GotoNamedRange( string name, bool activate )
	{
		var range =
			application.ActiveWorkbook.Names.Cast<MSExcel.Name>().FirstOrDefault( n => n.Name == name )?.RefersToRange ??
			application.ActiveWorkbook.Names.Cast<MSExcel.Name>().FirstOrDefault( n => n.Name == name.Split( '!' ).Last() )?.RefersToRange; // Incase they didn't scope sheet name properly, remove sheet and try

		if ( range != null )
		{
			range.Worksheet.Activate();
			if ( activate )
			{
				range.Select();
			}
		}
	}

	static string? GetTableDescription( string d ) =>
		d.StartsWith( "//" ) ? d[ 2.. ].Trim() :
		d.StartsWith( "#" ) ? d[ 1.. ].Trim() : null;
}