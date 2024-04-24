using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Domain.Extensions;
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

		var searchLocalCalcEnginesInfo = searchLocalCalcEngines.Search();

		if ( searchLocalCalcEnginesInfo == null )
		{
			return;
		}

		SaveWindowConfiguration( nameof( SearchLocalCalcEngines ), searchLocalCalcEnginesInfo.WindowConfiguration );

		MessageBox.Show( $"The KAT Addin will search all Excel files in {searchLocalCalcEnginesInfo.Folder} and will display results when complete.", "Search Local CalcEngines", MessageBoxButtons.OK, MessageBoxIcon.Information );

		application.StatusBar = "Searching CalcEngines...";
		
		ExcelAsyncUtil.QueueAsMacro( async () =>
		{
			// Search...
			var sb = new StringBuilder();
			var ssg = SpreadsheetGear.Factory.GetWorkbookSet();

			var currentFilePaths = application.Workbooks.Cast<MSExcel.Workbook>().Select( w => w.FullName ).ToArray();
			var calcEngines = new DirectoryInfo( searchLocalCalcEnginesInfo.Folder ).GetFiles()
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
				sb.AppendLine(
					$"{Environment.NewLine}***WARNING***{Environment.NewLine}Current CalcEngine will *not* be searched.  Any CalcEngines open in Excel can not be searched with the Audit feature." +
					$"{Environment.NewLine}{string.Join( Environment.NewLine, openCalcEngines.Select( c => $"- {c}" ) )}{Environment.NewLine}{Environment.NewLine}" );
			}

			foreach ( var file in calcEngines.Where( c => !openCalcEngines.Contains( c.FullName ) ) )
			{
				/*
				var wb = ssg.Workbooks.Open( file.FullName );

				SpreadsheetGear.IRange? lastOccur = null;
				foreach ( SpreadsheetGear.IWorksheet ws in wb.Worksheets )
				{
					foreach ( var token in searchLocalCalcEnginesInfo.Tokens )
					{
						var locationsFound = new HashSet<string>();

						while ( ( lastOccur = ws.Range.Find( token, lastOccur, SpreadsheetGear.FindLookIn.Formulas, SpreadsheetGear.LookAt.Part, SpreadsheetGear.SearchOrder.ByRows, SpreadsheetGear.SearchDirection.Next, false ) ) != null )
						{
							// Once the new instance loops back to the first instance, out of the loop.
							if ( locationsFound.Contains( lastOccur.Address ) )
								break;

							sb.AppendLine( $"{file.Name}: {ws.Name}!{lastOccur.Address} - {lastOccur.Formula ?? lastOccur.Text}" );

							locationsFound.Add( lastOccur.Address );
						}
					}
				}

				wb.Close();
				*/
			}

			if ( sb.Length > 0 )
			{
				ExcelDna.Logging.LogDisplay.Clear();
				ExcelDna.Logging.LogDisplay.WriteLine( "*****Local CalcEngines Search Results*****" );
				ExcelDna.Logging.LogDisplay.WriteLine( sb.ToString() );
				ExcelDna.Logging.LogDisplay.Show();
			}
			else
			{
				MessageBox.Show( $"No tokens were found in any CalcEngines.", "Search Local CalcEngines", MessageBoxButtons.OK, MessageBoxIcon.Information );
			}
		} );
	}

	public void Audit_CalcEngineTabs( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}