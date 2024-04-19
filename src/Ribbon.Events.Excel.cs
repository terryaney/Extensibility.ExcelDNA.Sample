using ExcelDna.Integration;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	private CalcEngineUploadInfo? calcEngineUploadInfo;
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
		
		if ( application.Workbooks.Count == 1 )
		{
			WorkbookState = new();
			ribbon.InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
		}
	}
	
	private async void Application_WorkbookActivate( MSExcel.Workbook wb )
	{
		// Clear error info whenever a new workbook is opened.  Currenly, only show any 
		// errors after a cell is calculated.  Could call application.Calculate() to force everything
		// to re-evaluate, but that could be expensive, so for now, not doing it, the function log display
		// is just helpful information for CalcEngine developer to 'clean' up their formulas.
		auditShowLogBadgeCount = 0;
		cellsInError.Clear();
		ExcelDna.Logging.LogDisplay.Clear();

		WorkbookState = await WorkbookState.GetCurrentAsync( application );
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
	}

	private void Application_WorkbookBeforeSave( MSExcel.Workbook wb, bool SaveAsUI, ref bool Cancel )
	{
		try
		{
			calcEngineUploadInfo = null;

			if ( auditShowLogBadgeCount > 0 )
			{
				RBLe_ShowLog( null );
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

			// TODO: Is this acceptable?  Excel seems to close successfully even without QueueAsMacro but can't make task async due to ref.
			calcEngineUploadInfo = ProcessSaveHistoryAsync( wb ).GetAwaiter().GetResult();
		}
		catch ( Exception ex )
		{
			LogError( $"Application_WorkbookBeforeSave", ex );
		}
	}

	private async void Application_WorkbookAfterSave( MSExcel.Workbook wb, bool Success )
	{
		try
		{
			ribbon.InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );

			Console.WriteLine( $"AfterSave Start: {DateTime.Now}" );
			await UploadCalcEngineToManagementSiteAsync();
		}
		catch ( Exception ex )
		{
			LogError( $"Application_WorkbookAfterSave", ex );
		}
		finally
		{
			calcEngineUploadInfo = null;
		}
	}

	private async void Application_SheetActivate( object sheet )
	{
		WorkbookState = await WorkbookState.GetCurrentAsync( application );
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnSheetChange );

		// Used to add event handlers to all charts that helped with old 'Excel' chart export 
		// functionality, but SSG does not support that so only use Highcharts/Apex now.

		// Used to update 'validation lists' in Tahiti spec sheets, but no longer use Tahiti.
	}

	private async Task EnsureAddInCredentialsAsync()
	{
		await Task.Delay( 2000 ); // TODO: Remove
		if ( WorkbookState.ShowCalcEngineManagement && ( string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Email ) || string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Password ) ) )
		{
			// TODO: Implement Credentials dialog
			// var credentials = new Credentials( AddIn.Settings.CalcEngineManagement.Email, AddIn.Settings.CalcEngineManagement.Password );
			// credentials.ShowDialog();

			// If updated...
			// await UpdateAddInCredentialsAsync( "terry.aney@conduent.com", "password" );
		}
	}

	private async Task UpdateAddInCredentialsAsync( string userName, string password )
	{
		application.StatusBar = "Saving CalcEngine Management credentials...";
		// TODO: Save new settings in appsettings.json and appsettings.secrets.json
		// Disable edit notifications...
		// calcEngineUploadInfo.UserName;
		var encryptedPassword = await CalcEngineManagement.EncryptPasswordAsync( password );
		// trigger Change() of appsettings.json file
		WorkbookState = await WorkbookState.GetCurrentAsync( application );
	}

	private DialogResult AuditCalcEngineTabs( MSExcel.Workbook workbook )
	{
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
			using var saveHistory = new SaveHistory( workbook, WorkbookState );

			var saveHistoryInfo = await saveHistory.GetHistoryInformationAsync();

			if ( saveHistoryInfo.Result == DialogResult.Ignore )
			{
				return null;
			}

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
					ExpectedVersion = currentVersion
				} 
				: null;
		}

		return null;
	}

	private async Task UploadCalcEngineToManagementSiteAsync()
	{
		if ( calcEngineUploadInfo != null )
		{
			try
			{
				if ( calcEngineUploadInfo.UserName != AddIn.Settings.CalcEngineManagement.Email || calcEngineUploadInfo.Password != await AddIn.Settings.CalcEngineManagement.GetClearPasswordAsync() )
				{
					await UpdateAddInCredentialsAsync( calcEngineUploadInfo.UserName, calcEngineUploadInfo.Password );
				}

				application.StatusBar = "Uploading CalcEngine to Management Site...";

				var ceContent = await File.ReadAllBytesAsync( application.ActiveWorkbook.FullName );

				// TODO: Upload ceContent, need to pass userName, password, expectedVersion, forceUpload (boolean) and confirm it can be done

				application.StatusBar = "CalcEngine successfully uploaded to Management Site.";

				// TODO: Can probably just update info of state instead of calling api to get updated info (latest version, etc.)
				WorkbookState = await WorkbookState.GetCurrentAsync( application );
				ribbon.InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
			}
			catch ( Exception ex )
			{
				application.StatusBar = "";
				MessageBox.Show( "Uploading CalcEngine to Management Site FAILED. " + ex.Message, "Upload Failed", MessageBoxButtons.OK, MessageBoxIcon.Error );
				throw;
			}
		}
	}
}