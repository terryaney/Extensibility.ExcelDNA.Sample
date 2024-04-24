using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	private bool skipHistoryUpdateOnMoveSpecFromDownloads;

	private void Application_WorkbookOpen( MSExcel.Workbook wb )
	{
		try
		{
			UpdateWorkbookLinks( wb );
		}
		catch ( Exception ex )
		{
			LogError( $"Application_WorkbookOpen", ex );
		}
	}

	private void Application_WorkbookDeactivate( MSExcel.Workbook Wb ) 
	{
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
		try
		{
			application.Cursor = MSExcel.XlMousePointer.xlWait;

			var wbName = application.ActiveWorkbook.Name;
			WorkbookState.UpdateWorkbook( wb );
			
			RunRibbonTask( async () => 
			{
				await EnsureAddInCredentialsAsync();
				await WorkbookState.UpdateCalcEngineInfoAsync( wbName );
				ExcelAsyncUtil.QueueAsMacro( () => ribbon.Invalidate() ); // .InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
			} );
		}
		catch ( Exception ex )
		{
			LogError( $"Application_WorkbookActivate", ex );
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	}

	private void Application_WorkbookBeforeSave( MSExcel.Workbook wb, bool SaveAsUI, ref bool Cancel )
	{
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
			LogError( $"Application_WorkbookBeforeSave", ex );
		}
	}

	private void Application_WorkbookAfterSave( MSExcel.Workbook wb, bool Success ) =>
		RunRibbonTask( () => UploadCalcEngineToManagementSiteAsync( wb ) );

	private void Application_SheetActivate( object sheet )
	{
		WorkbookState.UpdateSheet( ( application.ActiveWorkbook.ActiveSheet as MSExcel.Worksheet )! );
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnSheetChange );

		// Used to add event handlers to all charts that helped with old 'Excel' chart export 
		// functionality, but SSG does not support that so only use Highcharts/Apex now.

		// Used to update 'validation lists' in Tahiti spec sheets, but no longer use Tahiti.
	}

	private async Task EnsureAddInCredentialsAsync()
	{
		if ( WorkbookState.ShowCalcEngineManagement && ( string.IsNullOrEmpty( AddIn.Settings.KatUserName ) || string.IsNullOrEmpty( AddIn.Settings.KatPassword ) ) )
		{
			using var credentials = new Credentials( GetWindowConfiguration( nameof( Credentials ) ) );
			var credentialInfo = credentials.GetCredentials(  
				AddIn.Settings.KatUserName, 
				await AddIn.Settings.GetClearPasswordAsync() 
			);

			if ( credentialInfo != null )
			{
				await UpdateAddInCredentialsAsync( credentialInfo.UserName, credentialInfo.Password );
				SaveWindowConfiguration( nameof( Credentials ), credentialInfo.WindowConfiguration );
			}
		}
	}

	private DialogResult AuditCalcEngineTabs( MSExcel.Workbook workbook )
	{
		if ( !WorkbookState.IsCalcEngine )
		{
			return DialogResult.Yes;
		}
		
		var rblSheets = 
			workbook.Worksheets.Cast<MSExcel.Worksheet>()
				.Where( s => s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) && Constants.CalcEngines.SheetTypes.Contains( (string)s.Range[ "SheetType" ].Text ) ) );

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

	private async Task<CalcEngineUploadInfo?> ProcessSaveHistoryAsync( MSExcel.Workbook workbook )
	{
		if ( !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !skipHistoryUpdateOnMoveSpecFromDownloads )
		{
			using var saveHistory = new SaveHistory( workbook, WorkbookState, GetWindowConfiguration( nameof( SaveHistory ) ) );

			var saveHistoryInfo = saveHistory.GetHistoryInformation( 
				AddIn.Settings.SaveHistoryName, 
				AddIn.Settings.KatUserName, 
				await AddIn.Settings.GetClearPasswordAsync() 
			);

			if ( saveHistoryInfo.Result == DialogResult.Ignore )
			{
				return null;
			}

			SaveWindowConfiguration( nameof( SaveHistory ), saveHistoryInfo.WindowConfiguration );

			var currentVersion = (string?)saveHistoryInfo.VersionRange.Text;

			if ( saveHistoryInfo.Result != DialogResult.Retry )
			{
				// Update history log
				var descriptions = 
					saveHistoryInfo.Description?
						.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None )
						.Select( d => d.Replace( "\t", "  " ) )
						.Reverse()
						.ToArray() ?? Array.Empty<string>();

				var historyRange = saveHistoryInfo.HistoryRange.Offset[ 2, 0 ];
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
						historyRange.Value = saveHistoryInfo.Version;
						historyRange.Offset[ 0, 1 ].Value = string.Format( "{0:MM/dd/yyyy hh:mm tt}", DateTime.Now );
						historyRange.Offset[ 0, 2 ].Value = saveHistoryInfo.Author;
					}
				}

				saveHistoryInfo.VersionRange.Value = double.Parse( saveHistoryInfo.Version );
			}

			return saveHistoryInfo.Result != DialogResult.OK 
				? new()
				{
					UserName = saveHistoryInfo.UserName,
					Password = saveHistoryInfo.Password,
					ForceUpload = saveHistoryInfo.ForceUpload,
					ExpectedVersion = currentVersion,
					WindowConfiguration = saveHistoryInfo.WindowConfiguration
				} 
				: null;
		}

		return null;
	}

	private async Task UploadCalcEngineToManagementSiteAsync( MSExcel.Workbook wb )
	{
		var calcEngineUploadInfo = await ProcessSaveHistoryAsync( wb );

		if ( calcEngineUploadInfo != null )
		{
			try
			{
				if ( calcEngineUploadInfo.UserName != AddIn.Settings.KatUserName || calcEngineUploadInfo.Password != await AddIn.Settings.GetClearPasswordAsync() )
				{
					await UpdateAddInCredentialsAsync( calcEngineUploadInfo.UserName, calcEngineUploadInfo.Password );
				}

				SetStatusBar( "Uploading CalcEngine to Management Site..." );

				var ceContent = await File.ReadAllBytesAsync( application.ActiveWorkbook.FullName );

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
				ExcelAsyncUtil.QueueAsMacro( () => {
					MessageBox.Show( "Uploading CalcEngine to Management Site FAILED. " + ex.Message, "Upload Failed", MessageBoxButtons.OK, MessageBoxIcon.Error );
				} );
				throw;
			}
		}
	}
}