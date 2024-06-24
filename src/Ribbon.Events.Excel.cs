using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	private bool skipProcessSaveHistory;
	private bool skipWorkbookActivateEvents;

	private void Application_WorkbookOpen( MSExcel.Workbook wb )
	{
		try
		{
			UpdateWorkbookLinks( wb );
		}
		catch ( Exception ex )
		{
			ShowException( ex, "Application_WorkbookOpen" );
		}
	}

	private void Application_WorkbookDeactivate( MSExcel.Workbook Wb ) 
	{
		if ( skipWorkbookActivateEvents )
		{
			return;
		}

		// Used to simply trigger a SheetDeactivate if ActiveSheet != null
		/*
		if ( Wb.ActiveSheet != null )
		{
			Application_SheetDeactivate( Wb.ActiveSheet );
		}
		*/

		// Clear error info whenever a new workbook is opened.  Currenly, only show any 
		// errors after a cell is calculated.  Could call application.Calculate() to force everything
		// to re-evaluate, but that could be expensive, so for now, not doing it, the function log display
		// is just helpful information for CalcEngine developer to 'clean' up their formulas.
		auditShowLogBadgeCount = 0;
		cellsInError.Clear();
		ExcelDna.Logging.LogDisplay.Clear();

		if ( application.Workbooks.Count == 1 )
		{
			WorkbookState.ClearState();
			// Don't need invalidate if > 1 because WorkbookActivate will be called and trigger invalidate
			ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
		}
	}
	
	private void Application_WorkbookActivate( MSExcel.Workbook wb )
	{
		if ( skipWorkbookActivateEvents )
		{
			return;
		}

		try
		{
			application.Cursor = MSExcel.XlMousePointer.xlWait;

			var wbName = application.ActiveWorkbook.Name;
			WorkbookState.UpdateWorkbook( wb );
			
			RunRibbonTask( async () => 
			{
				await EnsureAddInCredentialsAsync();
				var validations = await WorkbookState.UpdateCalcEngineInfoAsync( wbName );
				ExcelAsyncUtil.QueueAsMacro( () =>
				{
					ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
					if ( validations != null )
					{
						ShowValidations( validations );
					}
				} );
			} );
		}
		catch ( Exception ex )
		{
			ShowException( ex, "Application_WorkbookActivate" );
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	}

	private void Application_WorkbookBeforeSave( MSExcel.Workbook wb, bool SaveAsUI, ref bool Cancel )
	{
		if ( skipProcessSaveHistory ) return;

		try
		{
			if ( auditShowLogBadgeCount > 0 )
			{
				Kat_ShowLog( null );
				Cancel = true;
				return;
			}

			// Used to call Application_SheetActivate if wb.ActiveSheet is not null to process Tahiti spec sheet lists

			var auditResult = AuditCalcEngineTabs( wb );

			if ( auditResult == DialogResult.Ignore )
			{
				return;
			}

			if ( auditResult == DialogResult.No )
			{
				Cancel = true;
				return;
			}
		}
		catch ( Exception ex )
		{
			ShowException( ex, "Application_WorkbookBeforeSave" );
		}
	}

	private void Application_WorkbookAfterSave( MSExcel.Workbook wb, bool Success )
	{
		if ( skipProcessSaveHistory ) return;

		RunRibbonTask( async () =>
		{
			var password = await AddIn.Settings.GetClearPasswordAsync();

			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				var info = ProcessSaveHistory( wb, password );

				if ( info != null )
				{
					skipProcessSaveHistory = true;
					wb.Save();
					RunRibbonTask( () => UploadCalcEngineToManagementSiteAsync( info ) );
				}
			} );
		} );
	}

	private void Application_SheetActivate( object sheet )
	{
		if ( skipWorkbookActivateEvents )
		{
			return;
		}

		WorkbookState.UpdateSheet( ( application.ActiveWorkbook.ActiveSheet as MSExcel.Worksheet )! );
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnSheetChange );

		// Used to add event handlers to all charts that helped with old 'Excel' chart export 
		// functionality, but SSG does not support that so only use Highcharts/Apex now.

		// Used to update 'validation lists' in Tahiti spec sheets, but no longer use Tahiti.
	}

	private DialogResult AuditCalcEngineTabs( MSExcel.Workbook workbook )
	{
		if ( !WorkbookState.IsCalcEngine )
		{
			return DialogResult.Yes;
		}
		
		var rblSheets = 
			workbook.Worksheets.Cast<MSExcel.Worksheet>()
				.Where( s => s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) && Constants.CalcEngines.IsRBLeSheet( (string)s.Range[ "SheetType" ].Text ) ) );

		var sheetsWithHiddenColumns = new List<string>();

		foreach ( var s in rblSheets )
		{
			if ( !s.Cells.SpecialCells( MSExcel.XlCellType.xlCellTypeVisible ).CountLarge.Equals( s.Cells.CountLarge ) )
			{
				sheetsWithHiddenColumns.Add( s.Name );
			}
		}

		if ( sheetsWithHiddenColumns.Count > 0 && MessageBox.Show( string.Format( "The following RBL sheets have hidden columns.  Hidden columns can adversely affect RBL processing.  Do you want to continue?\r\n\r\n{0}", string.Join( ", ", sheetsWithHiddenColumns ) ), "Continue with Hidden Columns?", MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.No )
		{
			return DialogResult.No;
		}

		if ( rblSheets.Any() )
		{
			try
			{
				// TODO: Create a mechanism to audit CalcEngine tabs and display message if invalid
				// var test = new RBLe.RBLeWorkbook( application.ActiveWorkbook.Name, application.ActiveWorkbook );
			}
			catch ( Exception ex )
			{
				MessageBox.Show( $"{application.ActiveWorkbook.Name} audit has failed.{Environment.NewLine + Environment.NewLine}The error is '{ex.Message}'.{Environment.NewLine + Environment.NewLine}The file will be saved automatic RBLInfo documentation and Management Site processing will NOT occur.", "Invalid Configuration", MessageBoxButtons.OK, MessageBoxIcon.Stop );
				return DialogResult.Ignore;
			}
		}

		return DialogResult.Yes;
	}

	private CalcEngineUploadInfo? ProcessSaveHistory( MSExcel.Workbook workbook, string? password )
	{
		if ( !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !skipProcessSaveHistory )
		{
			using var saveHistory = new SaveHistory( workbook, WorkbookState, GetWindowConfiguration( nameof( SaveHistory ) ) );

			var info = saveHistory.GetInfo( 
				AddIn.Settings.SaveHistoryName, 
				AddIn.Settings.KatUserName, 
				password 
			);

			if ( info.Result == DialogResult.Ignore )
			{
				return null;
			}

			SaveWindowConfiguration( nameof( SaveHistory ), info.WindowConfiguration );

			var currentVersion = (string?)info.VersionRange.Text;

			if ( info.Result != DialogResult.Retry )
			{
				// Update history log
				var descriptions = 
					info.Description?
						.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None )
						.Select( d => d.Replace( "\t", "  " ) )
						.Reverse()
						.ToArray() ?? Array.Empty<string>();

				var historyRange = info.HistoryRange.Offset[ 2, 0 ];
				var historySheet = historyRange.Worksheet;

				for (var i = 0; i < descriptions.Length; i++)
				{
					historyRange.Offset[ 1, 0 ].EntireRow.Insert( MSExcel.XlInsertShiftDirection.xlShiftDown );
					historySheet.Range[ historyRange, historyRange.Offset[ 0, 3 ] ].Copy( historyRange.Offset[ 1, 0 ] );
					historySheet.Range[ historyRange, historyRange.Offset[ 0, 3 ] ].Value = null;

					historyRange.Offset[ 0, 3 ].Value = descriptions[ i ];

					// If last row...
					if ( i == descriptions.Length - 1 )
					{
						historyRange.Value = info.Version;
						historyRange.Offset[ 0, 1 ].Value = string.Format( "{0:MM/dd/yyyy hh:mm tt}", DateTime.Now );
						historyRange.Offset[ 0, 2 ].Value = info.Author;
					}
				}

				info.VersionRange.Value = double.Parse( info.Version );
			}

			return info.Result != DialogResult.OK 
				? new()
				{
					UserName = info.UserName,
					Password = info.Password,
					ForceUpload = info.ForceUpload,
					ExpectedVersion = currentVersion,
					WindowConfiguration = info.WindowConfiguration, // not needed but my class derived from requirement
					FullName = application.ActiveWorkbook.FullName
				} 
				: null;
		}

		skipProcessSaveHistory = false;
		return null;
	}

	private async Task UploadCalcEngineToManagementSiteAsync( CalcEngineUploadInfo info )
	{
		try
		{
			await UpdateAddInCredentialsAsync( info.UserName, info.Password );

			var ceContent = await File.ReadAllBytesAsync( info.FullName );

			SetStatusBar( "Uploading CalcEngine to Management Site..." );

			// TODO: Upload ceContent, need to pass userName, password, expectedVersion, forceUpload (boolean) and confirm it can be done

			SetStatusBar( "CalcEngine successfully uploaded to Management Site." );

			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				WorkbookState.UpdateVersion( application.ActiveWorkbook );
				ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
			} );
		}
		catch ( Exception ex )
		{
			ClearStatusBar();
			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				MessageBox.Show( "Uploading CalcEngine to Management Site FAILED. " + ex.Message, "Upload Failed", MessageBoxButtons.OK, MessageBoxIcon.Error );
			} );
			throw;
		}
	}
}