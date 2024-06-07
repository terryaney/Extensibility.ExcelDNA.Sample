using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.RBLe.Calculations;
using KAT.Camelot.Data.Repositories;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Domain.IO;
using KAT.Camelot.Domain.Telemetry;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;
using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	string? lastPopulateInputTabAuthId = null;

	public void CalcEngineUtilities_PopulateInputTab( IRibbonControl _ )
	{
		RunRibbonTask( async () =>
		{
			var password = await AddIn.Settings.GetClearPasswordAsync();

			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				var owner = new NativeWindow();
				owner.AssignHandle( new IntPtr( application.Hwnd ) );

				var config = GetWindowConfiguration( nameof( LoadInputTab ) );

				using var loadInputTab = new LoadInputTab(
					application.ActiveWorkbook.RangeOrNull<string>( "ProfileLoadGroup" ),
					lastPopulateInputTabAuthId,
					WorkbookState.ManagementName,
					AddIn.Settings.DataServices,
					config
				);

				var info = loadInputTab.GetInfo(
					AddIn.Settings.KatUserName,
					password,
					owner
				);

				if ( info == null ) return;

				var globalTablesIsOpen = info.DownloadGlobalTables && DownloadLatestCalcEngineCheck( Constants.FileNames.GlobalTables, AddIn.ResourcesPath ) == null;

				application.ScreenUpdating = false;
				application.Cursor = MSExcel.XlMousePointer.xlWait;

				try
				{
					SaveWindowConfiguration( nameof( LoadInputTab ), info.WindowConfiguration );

					var cts = new CancellationTokenSource();

					RunRibbonTask( async () =>
					{
						(string AuthId, Dictionary<string, string> Profile, Dictionary<string, JsonArray> History)? data = null;
						ApiValidation[]? validations = null;

						if ( string.IsNullOrEmpty( info.DataSourceFile ) )
						{
							await UpdateAddInCredentialsAsync( info.UserName!, info.Password! );

							SetStatusBar( $"Requesting data from {info.DataSource}..." );
							var response = await apiService.GetxDSDataAsync( info.ClientName, info.AuthId!, info.DataSource, info.UserName, info.Password, cts.Token );

							validations = response.Validations;

							if ( response.Response != null )
							{
								var history =
									response.Response!.History.ToDictionary(
										h => h.Key,
										h => h.Value.Select( r =>
											new JsonObject()
												.AddProperties(
													r.Select( p => new JsonKeyProperty( p.Key, p.Value ) )
												)
										).ToJsonArray()
									);

								data = (response.Response!.AuthId, response.Response!.Profile, history);
							}
						}
						else
						{
							data = await GetInputDataAsync( info, cts.Token );
						}

						if ( validations != null || data == null )
						{
							ExcelAsyncUtil.QueueAsMacro( () =>
							{
								application.ScreenUpdating = true;
								if ( validations != null )
								{
									ShowValidations( validations );
								}
								else
								{
									MessageBox.Show( "No data found for the specified Client and AuthId.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning );
								}
							} );
							return;
						}							

						lastPopulateInputTabAuthId = data.Value.AuthId;

						await DownloadLookupFilesAsync( info, globalTablesIsOpen, cts.Token );

						ExcelAsyncUtil.QueueAsMacro( () =>
						{
							application.Cursor = MSExcel.XlMousePointer.xlWait;

							try
							{
								var fileName = application.ActiveWorkbook.Name;
								var tabName = application.ActiveWorksheet().Name;

								if ( !globalTablesIsOpen && info.LoadLookupTables && File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) ) )
								{
									application.Workbooks.Open( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) );
								}

								LookupsConfiguration? globalLookupsConfiguration;
								try
								{
									globalLookupsConfiguration = info.LoadLookupTables
										? new DnaCalcEngineConfigurationFactory( Constants.FileNames.GlobalTables ).LookupsConfiguration
										: null;
								}
								finally
								{
									if ( info.LoadLookupTables && !globalTablesIsOpen )
									{
										application.GetWorkbook( Constants.FileNames.GlobalTables )!.Close( false );
									}
								}

								var configuration = new DnaCalcEngineConfigurationFactory( fileName ).Configuration;
								using var calcEngine = new DnaCalcEngine( fileName, configuration );

								var parameters = new CalculationParameters
								{
									CalculationId = Guid.NewGuid(),
									RequestInfo = new SingleRequest()
									{
										AuthId = "LOCAL.DEBUG",
										CalcEngines = Array.Empty<RequestCalcEngine>(),
										TraceEnabled = true
									},
									Payload = new()
									{
										Profile = data.Value.Profile,
										History = data.Value.History
									},
									LookupTables = GetLookupTables( info, globalLookupsConfiguration )
								};

								var diagnosticTraceLogger = new DiagnosticTraceLogger();
								diagnosticTraceLogger.Start();

								var calculationMode = application.Calculation;
								application.Calculation = MSExcel.XlCalculation.xlCalculationManual;

								try
								{
									dnaCalculationService.LoadParticipantData( calcEngine, parameters, tabName, diagnosticTraceLogger );
								}
								finally
								{
									application.Calculation = calculationMode;
								}
							}
							finally
							{
								application.Cursor = MSExcel.XlMousePointer.xlDefault;
								application.ScreenUpdating = true;
							}
						} );
					} );
				}
				catch
				{
					application.Cursor = MSExcel.XlMousePointer.xlDefault;
					application.ScreenUpdating = true;
					throw;
				}
			} );
		} );
	}

	private async Task DownloadLookupFilesAsync( LoadInputTabInfo info, bool globalTablesIsOpen, CancellationToken cancellationToken )
	{
		if ( !info.LoadLookupTables ) return;

		if ( !string.IsNullOrEmpty( info.ConfigLookupsUrl ) )
		{
			// Make http request to get ConfigLookupsUrl (which returns xml file) and trap any errors
			var client = httpClientFactory.CreateClient();
			var response = await client.GetAsync( info.ConfigLookupsUrl, cancellationToken );
			if ( response.IsSuccessStatusCode )
			{
				var lookups = await XElement.LoadAsync( await response.Content.ReadAsStreamAsync( cancellationToken ), LoadOptions.None, cancellationToken );
				var path = Path.Combine( AddIn.ResourcesPath, "ConfigLookups", $"{info.ClientName}.xml" );
				Directory.CreateDirectory( Path.GetDirectoryName( path )! );
				lookups.Save( path );
			}
		}

		if ( info.DownloadGlobalTables && !globalTablesIsOpen )
		{
			await DownloadLatestCalcEngineAsync( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ), false );
		}
	}

	private static LookupTable[] GetLookupTables( LoadInputTabInfo info, LookupsConfiguration? globalLookupsConfiguration )
	{
		if ( !info.LoadLookupTables ) return Array.Empty<LookupTable>();

		var clientLookupPath = new []
		{
			info.ConfigLookupsPath,
			Path.Combine( AddIn.ResourcesPath, "ConfigLookups", $"{info.ClientName}.xml" )
		}.FirstOrDefault( File.Exists );

		var clientLookups = !string.IsNullOrEmpty( clientLookupPath )
			? XElement.Load( clientLookupPath )
			: new XElement( "DataTableDefs" );

		var globalTables =
			new XElement( "Tables",
				globalLookupsConfiguration?.Tables.Select( t =>
					new XElement( "Table",
						new XElement( "TableName", t.Name ),
						t.Rows.Cast<JsonObject>().Select( r =>
							new XElement( "TableRow",
								r.Select( p =>
									new XElement( "ItemDef",
										new XAttribute( "Name", p.Key ),
										p.Value
									)
								)
							)
						)
					)
				)
			);

		xDSRepository.MergeGlobalTables( clientLookups, globalTables );

		var rblLookups = xDSRepository.MapGlobalTablesToRble( clientLookups.Elements( "DataTable" ) );

		var lookups = 
			rblLookups.Elements( "TableDef" )
				.Select( t => new LookupTable
				{
					Name = (string)t.Element( "TableName" )!,
					Rows = new JsonArray
					(
						t.Elements( "TableRow" )
							.Select( r => new JsonObject().AddProperties(
								r.Elements( "ItemDef" )
									.Select( c => new JsonKeyProperty( (string)c.Attribute( "Name" )!, (string)c ) )
							) )
							.ToArray()
					)
				} ).ToArray();

		return lookups;
	}

	private static async Task<(string AuthId, Dictionary<string, string> Profile, Dictionary<string, JsonArray> History)?> GetInputDataAsync( LoadInputTabInfo info, CancellationToken cancellationToken )
	{
		var dateAttribute = XmlConvert.ToString( DateTime.Now, XmlDateTimeSerializationMode.Local );
		var authorAttribute = AddIn.Settings.KatUserName?.Split( '@' ).First() ?? "LOCAL.DEBUG";

		if ( string.Compare( Path.GetExtension( info.DataSourceFile ), ".json", true ) == 0 )
		{
			static (string AuthId, Dictionary<string, string> Profile, Dictionary<string, JsonArray> History) getInputData( JsonObject json, string dateAttribute, string authorAttribute )
			{
				var profile =
					( json[ "profile" ] as JsonObject )!.ToDictionary( p => p.Key, p => p.Value?.ToString() );
				var history =
					( json[ "history" ] as JsonObject )!
						.ToDictionary(
							p => p.Key,
							p => ( p.Value as JsonArray )!
						);

				return ((string)json[ "@authId" ]!, profile, history)!;
			}

			if ( info.DataSource == LoadInputTab.LocalProfileFile )
			{
				var json = ( JsonNode.Parse( await File.ReadAllTextAsync( info.DataSourceFile!, cancellationToken: cancellationToken ) ) as JsonObject )!;
				return getInputData( json, dateAttribute, authorAttribute );
			}
			else
			{
				using var fileStream = new FileStream( info.DataSourceFile!, FileMode.Open, FileAccess.Read );

				await foreach ( var json in JsonSerializer.DeserializeAsyncEnumerable<JsonObject>( fileStream, cancellationToken: cancellationToken ) )
				{
					if ( (string?)json![ "@authId" ] == info.AuthId )
					{
						return getInputData( json, dateAttribute, authorAttribute );
					}
				}
			}
		}
		else
		{
			static (string AuthId, Dictionary<string, string> Profile, Dictionary<string, JsonArray> History)? getInputData( XElement? xml, string dateAttribute, string authorAttribute )
			{
				if ( xml == null ) return null;

				var profile = xml.Profile().Elements().ToDictionary( e => e.Name.LocalName, e => (string)e );
				var history = 
					xml.HistoryItems().GroupBy( h => h.hisType() )
						.ToDictionary(
							g => g.Key,
							g => g.Select( h => 
								new JsonObject()
									.AddProperties( 
										h.Elements()
											.Select( e => new JsonKeyProperty( e.Name.LocalName, (string)e ) ) 
									)
							).ToJsonArray()
						);

				return ((string)xml.xDataDef().Attribute( "id-auth" )!, profile, history)!;
			}

			if ( info.DataSource == LoadInputTab.LocalProfileFile )
			{
				var xml = XElement.Load( info.DataSourceFile! );
				return getInputData( xml, dateAttribute, authorAttribute );
			}
			else
			{
				using var reader = new xDataDefReader( info.DataSourceFile! );
				return getInputData( reader.FindAuthId( info.AuthId! ), dateAttribute, authorAttribute );
			}
		}

		return null;
	}

	public void CalcEngineUtilities_RunMacros( IRibbonControl _ )
	{
		var helpersOpen = application.GetWorkbook( Constants.FileNames.Helpers ) != null;

		if ( !helpersOpen && !File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.Helpers ) ) )
		{
			MessageBox.Show( "The Helpers workbook is missing.  Please download it before processing the workbook.", "Missing Helpers", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
		}

		ExcelAsyncUtil.QueueAsMacro( async () =>
		{
			var diagnosticTraceLogger = new DiagnosticTraceLogger();
			diagnosticTraceLogger.Start();

			// Haven't found C API equivalent to setting Saved property...
			var isSaved = application.ActiveWorkbook.Saved;
			var fileName = application.ActiveWorkbook.Name;

			application.ScreenUpdating = false;

			// TODO: See why .RestoreSelection doesn't work here.
			var selection = DnaApplication.Selection;

			try
			{
				skipWorkbookActivateEvents = true;

				var configuration = new DnaCalcEngineConfigurationFactory( fileName ).Configuration;
				using var calcEngine = new DnaCalcEngine( fileName, configuration );

				var helpersWb =
					application.GetWorkbook( Constants.FileNames.Helpers ) ??
					application.Workbooks.Open( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.Helpers ) );
				var helpersFilename = helpersWb.Name;

				try
				{
					using var helpers = new DnaCalcEngine( helpersFilename );

					calcEngine.Activate();
					var cts = new CancellationTokenSource();

					var parameters = new CalculationParameters
					{
						CalculationId = Guid.NewGuid(),
						RequestInfo = new SingleRequest() 
						{
							AuthId = "LOCAL.DEBUG",
							CalcEngines = Array.Empty<RequestCalcEngine>(),
							TraceEnabled = true
						},
						// TODO: Would need to build this from input tab and all 'data' elements passed in...
						Payload = new()
						{
							Profile = new(),
							History = new()
						}
					};

					var appliedDataUpdates = await dnaCalculationService.ProcessMacrosAsync( calcEngine, helpers, diagnosticTraceLogger, cts.Token );

					MessageBox.Show( "The RBLe Macros ran with no errors.", "RBLe Macros Succeeded", MessageBoxButtons.OK, MessageBoxIcon.Information );

					if ( diagnosticTraceLogger.HasTrace )
					{
						ExcelDna.Logging.LogDisplay.Clear();
						foreach ( var t in diagnosticTraceLogger.Trace )
						{
							ExcelDna.Logging.LogDisplay.WriteLine( t.Replace( "\t", "    " ) );
						}
						ExcelDna.Logging.LogDisplay.Show();
					}
				}
				finally
				{
					if ( !helpersOpen )
					{
						helpersWb.Close( false );
					}
				}
			}
			catch ( Exception ex )
			{
				MessageBox.Show( "The RBLe Macros failed.  See log for details.", "RBLe Macros Failed", MessageBoxButtons.OK, MessageBoxIcon.Error );

				ExcelDna.Logging.LogDisplay.Clear();

				ShowException( ex, null, diagnosticTraceLogger.HasTrace ? new [] { "", "RBLe Macro Trace" }.Concat( diagnosticTraceLogger.Trace.Select( t => t.Replace( "\t", "    " ) ) ) : null );
			}
			finally
			{
				selection.Select();
				skipWorkbookActivateEvents = false;
				application.ScreenUpdating = true;
				application.ActiveWorkbook.Saved = isSaved;
			}
		} );
	}

	public void CalcEngineUtilities_PreviewResults( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_LocalBatchCalc( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_DownloadGlobalTables( IRibbonControl _ )
	{
		var fullName = DownloadLatestCalcEngineCheck( Constants.FileNames.GlobalTables, AddIn.ResourcesPath );
		RunRibbonTask( () => DownloadLatestCalcEngineAsync( fullName ) );
	}

	public void CalcEngineUtilities_DownloadHelpersCalcEngine( IRibbonControl _ )
	{
		var fullName = DownloadLatestCalcEngineCheck( Constants.FileNames.Helpers, AddIn.ResourcesPath );
		RunRibbonTask( () => DownloadLatestCalcEngineAsync( fullName ) );
	}

	public void CalcEngineUtilities_LinkToLoadedAddIns( IRibbonControl _ ) => UpdateWorkbookLinks( application.ActiveWorkbook );

	private void UpdateWorkbookLinks( MSExcel.Workbook wb )
	{
		if ( wb == null )
		{
			ExcelDna.Logging.LogDisplay.RecordLine( $"LinkToLoadedAddIns: ActiveWorkbook is null." );
			return;
		}

		if ( Path.GetFileName( wb.Name ) != "RBL.Template.xlsx" || !WorkbookState.HasLinks ) return;

		var linkSources = ( wb.LinkSources( MSExcel.XlLink.xlExcelLinks ) as Array )!;

		var protectedInfo = wb.ProtectStructure
			? new[] { "Entire Workbook" }
			: wb.Worksheets.Cast<MSExcel.Worksheet>().Where( w => w.ProtectContents ).Select( w => string.Format( "Worksheet: {0}", w.Name ) ).ToArray();

		if ( protectedInfo.Length > 0 )
		{
			MessageBox.Show( "Unable to update links due to protection.  The following items are protected:\r\n\r\n" + string.Join( "\r\n", protectedInfo ), "Unable to Update", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
		}

		var saved = wb.Saved;

		foreach ( var addin in application.AddIns.Cast<MSExcel.AddIn>().Where( a => a.Installed ) )
		{
			var fullName = addin.FullName;
			var name = Path.GetFileName( fullName );

			foreach ( var o in linkSources )
			{
				var link = (string)o;
				var linkName = Path.GetFileName( link );

				if ( string.Compare( name, linkName, true ) == 0 )
				{
					try
					{
						application.ActiveWorkbook.ChangeLink( link, fullName );
					}
					catch ( Exception ex )
					{
						ExcelDna.Logging.LogDisplay.RecordLine( $"LinkToLoadedAddIns Exception:\r\n\tAddIn Name:{addin.Name}\r\n\tapplication Is Null:{application == null}\r\n\tapplication.ActiveWorkbook Is Null:{application?.ActiveWorkbook == null}\r\n\tName: {name}\r\n\tLink: {link}\r\n\tFullName: {fullName}\r\n\tMessage: {ex.Message}" );
						throw;
					}
				}
			}
		}

		wb.Saved = saved;
	}
}