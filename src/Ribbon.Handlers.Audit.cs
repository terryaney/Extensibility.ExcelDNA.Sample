using System.Diagnostics;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;
using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;
using XLParser;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Audit_ShowDependencies( IRibbonControl _ )
	{
		foreach ( MSExcel.Range cell in ( application.Selection as MSExcel.Range )! )
		{
			cell.ShowDependents();
		}
	}

	public void Audit_HideDependencies( IRibbonControl _ )
	{
		foreach ( MSExcel.Range cell in ( application.Selection as MSExcel.Range )! )
		{
			cell.ShowDependents( true );
		}
	}

	public void Audit_ShowCellsWithEmptyDependencies( IRibbonControl _ )
	{
		var selection = ( application.Selection as MSExcel.Range )!;
		selection.Style = "Normal";

		var selectionRef = selection.GetReference();
		var firstCell = selectionRef.Corner( CornerType.UpperLeft );

		for ( var row = 0; row <= selectionRef.RowLast - selectionRef.RowFirst; row++ )
		{
			for ( var col = 0; col <= selectionRef.ColumnLast - selectionRef.ColumnFirst; col++ )
			{
				var cell = firstCell.Offset( row, col );
				var cellFormula = cell.GetFormula();

				if ( !string.IsNullOrEmpty( cellFormula ) && cellFormula.StartsWith( "=" ) )
				{
					var tree = ExcelFormulaParser.Parse( cellFormula );
					var references =
						tree.AllNodes()
							.Where( n => n.Term.Name == "NamedRange" || n.Term.Name == "Cell" )
							.Select( r => r.ChildNodes[ 0 ].Token.Text );

					foreach ( var r in references )
					{
						var reference = r.GetReference();
						var data = reference.GetValue();
						var dataValues = data as object[,];

						if ( dataValues?.Contains( ExcelEmpty.Value, false ) ?? Equals( data, ExcelEmpty.Value ) )
						{
							cell.GetRange().Style = "Bad";
							break;
						}
					}
				}
			}
		}
	}

	public void Audit_SearchLocalCalcEngines( IRibbonControl _ )
	{
		using var searchLocalCalcEngines = new SearchLocalCalcEngines( GetWindowConfiguration( nameof( SearchLocalCalcEngines ) ) );

		var info = searchLocalCalcEngines.GetInfo();

		if ( info == null )
		{
			return;
		}

		SaveWindowConfiguration( nameof( SearchLocalCalcEngines ), info.WindowConfiguration );

		// MessageBox.Show( $"The KAT Addin will search all Excel files in {searchLocalCalcEnginesInfo.Folder} and will display results when complete.", "Search Local CalcEngines", MessageBoxButtons.OK, MessageBoxIcon.Information );

		var currentFilePaths = application.Workbooks.Cast<MSExcel.Workbook>().Select( w => w.FullName ).ToArray();
		var csvFile = Path.Combine( AddIn.ResourcesPath, "SearchLocalCalcEngines.csv" );
		var csvWorkbook = application.Workbooks.Cast<MSExcel.Workbook>().FirstOrDefault( w => string.Compare( csvFile, w.FullName, true ) == 0 );
		csvWorkbook?.Close( false );

		RunRibbonTask( async () =>
		{
			if ( await EnsureSpreadsheetGearLicenseAsync() == false ) return;

			SetStatusBar( "Searching Local CalcEngines..." );

			var results = new List<(string CalcEngine, string Tab, string Address, string Formula)>();
			var ssg = SpreadsheetGear.Factory.GetWorkbookSet();

			var calcEngines = new DirectoryInfo( info.Folder ).GetFiles()
					.Where( f => new[] { ".xls", ".xlsm" }.Contains( f.Extension, StringComparer.InvariantCultureIgnoreCase ) )
					.Where( f => !f.Name.StartsWith( "~" ) )
					.ToArray(); // temp files

			var openCalcEngines =
				calcEngines
					.Where( c => currentFilePaths.Contains( c.FullName, StringComparer.InvariantCultureIgnoreCase ) )
					.Select( c => c.FullName )
					.ToArray();

			if ( openCalcEngines.Any() )
			{
				results.AddRange( openCalcEngines.Select( c => (c, "N/A", "N/A", "CalcEngines open in Excel can not be searched with the Audit feature.") ) );
			}

			foreach ( var file in calcEngines.Where( c => !openCalcEngines.Contains( c.FullName ) ) )
			{
				var wb = ssg.Workbooks.Open( file.FullName );

				try
				{
					SpreadsheetGear.IRange? lastOccur = null;
					foreach ( SpreadsheetGear.IWorksheet ws in wb.Worksheets )
					{
						foreach ( var token in info.Tokens )
						{
							var locationsFound = new HashSet<string>();

							while ( ( lastOccur = ws.Range.Find( token, lastOccur, SpreadsheetGear.FindLookIn.Formulas, SpreadsheetGear.LookAt.Part, SpreadsheetGear.SearchOrder.ByRows, SpreadsheetGear.SearchDirection.Next, false ) ) != null )
							{
								// Once the new instance loops back to the first instance, out of the loop.
								if ( locationsFound.Contains( lastOccur.Address ) )
									break;

								results.Add( (file.Name, ws.Name, lastOccur.Address, $"'{lastOccur.Formula}") );

								locationsFound.Add( lastOccur.Address );
							}
						}
					}
				}
				finally
				{
					wb.Close();
				}
			}

			ClearStatusBar();

			if ( results.Count > 0 )
			{
				await results.DumpCsvAsync( 
					csvFile,
					getHeader: m => m.Name switch
					{
						"Item1" => "CalcEngine",
						"Item2" => "Tab",
						"Item3" => "Address",
						"Item4" => "Formula",
						_ => throw new ArgumentException( "Unknown member", m.Name )
					}
				);
				ExcelAsyncUtil.QueueAsMacro( () => application.Workbooks.Open( csvFile ) );
			}
			else
			{
				ExcelAsyncUtil.QueueAsMacro( () => MessageBox.Show( $"No tokens were found in any CalcEngines.", "Search Local CalcEngines", MessageBoxButtons.OK, MessageBoxIcon.Information ) );
			}
		} );
	}

	public void Audit_CalcEngineTabs( IRibbonControl _ )
	{
		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			var name = application.ActiveWorkbook.Name;

			try
			{
				application.ScreenUpdating = false;

				var sw = Stopwatch.StartNew();
				var configuration = new ExcelCalcEngineConfigurationFactory( application.ActiveWorkbook ).Configuration;
				Console.WriteLine( $"ExcelCalcEngineConfigurationFactory configuration took {sw.ElapsedMilliseconds}ms." );
				sw.Restart();
				var dnaConfiguration = new DnaCalcEngineConfigurationFactory( application.ActiveWorkbook.Name ).Configuration;
				Console.WriteLine( $"DnaCalcEngineConfigurationFactory configuration took {sw.ElapsedMilliseconds}ms." );

				MessageBox.Show( $"All RBLe tabs in {name} are correctly configured ({configuration.InputTabs.Length} Input Tab(s) and {configuration.ResultTabs.Length} Result Tab(s)).", "CalcEngine Audit", MessageBoxButtons.OK, MessageBoxIcon.Information );
			}
			catch ( CalcEngineConfigurationException ex )
			{
				ExcelDna.Logging.LogDisplay.WriteLine( $"The RBLe tabs in {name} are incorrectly configured.  See the issues below for more details." + Environment.NewLine );

				foreach( var error in ex.Errors )
				{
					ExcelDna.Logging.LogDisplay.WriteLine( error );
				}

				ExcelDna.Logging.LogDisplay.Show();
			}
			finally
			{
				application.ScreenUpdating = true;
			}
		} );
	}
}