using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Navigation_NavigateToTable( IRibbonControl _ )
	{
		var tables = new List<NavigationTable>();

		var activeSheet = application.ActiveWorksheet();

		static string? getDescription( string d ) =>
			d.StartsWith( "//" ) ? d[ 2.. ].Trim() :
			d.StartsWith( "#" ) ? d[ 1.. ].Trim() : null;

		if ( activeSheet.Names.Cast<MSExcel.Name>().Any( n => n.Name == "StartTables" || n.Name.EndsWith( "!StartTables" ) ) )
		{
			// CalcEngine tables...
			var start = activeSheet.Range[ "StartTables" ].Offset[ 1, 0 ];
			string? name = null;

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

				tables.Add( new ()
				{
					Name = name.Split( '/' )[ 0 ] + suffix,
					Address = start.Address,
					Description = getDescription( description )
				} );

				start = start.Offset[ 1, 0 ].End[ MSExcel.XlDirection.xlToRight ].Offset[ -1, 2 ];
			}
		}
		else if ( activeSheet.Name == "Historical Data" )
		{
			// Spec Sheet History Table
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

				tables.Add( new ()
				{
					Name = start.Offset[ 0, 1 ].GetText(),
					Address = start.Offset[ 4, 0 ].Address,
					Description = getDescription( description ) ?? description
				} );

				start = start.End[ MSExcel.XlDirection.xlDown ].Offset[ 2, 0 ];
			}
		}
		else if ( activeSheet.Name == "Code Tables" )
		{
			// Spec sheet
			var start = activeSheet.Range[ "A5" ];

			while ( string.Compare( start.GetText(), "Table", true ) == 0 )
			{
				start = start.Offset[ 0, 1 ];

				var description = start.Offset[ -1, 0 ].GetText();

				tables.Add( new ()
				{
					Name = start.GetText(),
					Address = start.Address,
					Description = getDescription( description ) ?? description
				} );

				start = start.End[ MSExcel.XlDirection.xlToRight ].Offset[ 0, 2 ];
			}
		}

		using var navigateToTable = new NavigateToTable( tables, GetWindowConfiguration( nameof( NavigateToTable ) ) );

		var navigationResult = navigateToTable.GetTarget();

		if ( navigationResult == null )
		{
			return;
		}
		activeSheet.Range[ navigationResult.Target ].Select();

		SaveWindowConfiguration( nameof( NavigateToTable ), navigationResult.WindowConfiguration );
	}

	public void Navigation_GoToInputs( IRibbonControl _ ) => GotoInputNamedRange( "StartData" );
	public void Navigation_GoToInputData( IRibbonControl _ ) => GotoInputNamedRange( "xDSDataFields" );
	public void Navigation_GoToCalculationInputs( IRibbonControl _ ) => GotoInputNamedRange( "CalculationInputs" );
	public void Navigation_GoToFrameworkInputs( IRibbonControl _ ) => GotoInputNamedRange( "FrameworkInputs" );
	public void Navigation_GoToInputTables( IRibbonControl _ ) => GotoInputNamedRange( "StartTables" );

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

	public void Navigation_BackToRBLeMacro( IRibbonControl _ ) => GotoNamedRange( "RBLeMacro", false );

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
}