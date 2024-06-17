using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.RBLe;
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

	public void CalcEngineUtilities_LoadInputTab( IRibbonControl _ )
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
						(string AuthId, RBLePayload PayLoad)? data = null;
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

								data = (response.Response!.AuthId, new () { Profile = response.Response!.Profile, History = history });
							}
						}
						else
						{
							data = await GetRBLePayloadAsync( info, cts.Token );
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
									CalcEngineInfo = new()
									{
										InputTab = tabName
									},
									RequestInfo = new SingleRequest()
									{
										AuthId = "LOCAL.DEBUG",
										CalcEngines = Array.Empty<RequestCalcEngine>(),
										TraceEnabled = true
									},
									Payload = data.Value.PayLoad,
									LookupTables = GetLookupTables( info, globalLookupsConfiguration )
								};

								var diagnosticTraceLogger = new DiagnosticTraceLogger();
								diagnosticTraceLogger.Start();

								var calculationMode = application.Calculation;
								application.Calculation = MSExcel.XlCalculation.xlCalculationManual;

								try
								{
									dnaCalculationService.LoadParticipantData( calcEngine, parameters, diagnosticTraceLogger );
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

	private static async Task<(string AuthId, RBLePayload Payload)?> GetRBLePayloadAsync( LoadInputTabInfo info, CancellationToken cancellationToken )
	{
		var dateAttribute = XmlConvert.ToString( DateTime.Now, XmlDateTimeSerializationMode.Local );
		var authorAttribute = AddIn.Settings.KatUserName?.Split( '@' ).First() ?? "LOCAL.DEBUG";

		if ( string.Compare( Path.GetExtension( info.DataSourceFile ), ".json", true ) == 0 )
		{
			if ( info.DataSource == LoadInputTab.LocalProfileFile )
			{
				var json = ( JsonNode.Parse( await File.ReadAllTextAsync( info.DataSourceFile!, cancellationToken: cancellationToken ) ) as JsonObject )!;
				return GetRBLePayload( json /*, dateAttribute, authorAttribute */ );
			}
			else
			{
				using var fileStream = new FileStream( info.DataSourceFile!, FileMode.Open, FileAccess.Read );

				await foreach ( var json in JsonSerializer.DeserializeAsyncEnumerable<JsonObject>( fileStream, cancellationToken: cancellationToken ) )
				{
					if ( (string?)json![ "@authId" ] == info.AuthId )
					{
						return GetRBLePayload( json /*, dateAttribute, authorAttribute */ );
					}
				}
			}
		}
		else
		{
			if ( info.DataSource == LoadInputTab.LocalProfileFile )
			{
				var xml = XElement.Load( info.DataSourceFile! );
				return GetRBLePayload( xml /*, dateAttribute, authorAttribute */ );
			}
			else
			{
				using var reader = new xDataDefReader( info.DataSourceFile! );
				return GetRBLePayload( reader.FindAuthId( info.AuthId! ) /*, dateAttribute, authorAttribute */ );
			}
		}

		return null;
	}

	public static (string AuthId, RBLePayload Payload) GetRBLePayload( JsonObject json /*, string dateAttribute, string authorAttribute */ )
	{
		var profile =
			( json[ "profile" ] as JsonObject )!.ToDictionary( p => p.Key, p => p.Value?.ToString() );
		var history =
			( json[ "history" ] as JsonObject )!
				.ToDictionary(
					p => p.Key,
					p => ( p.Value as JsonArray )!
				);

		return ((string)json[ "@authId" ]!, new() { Profile = profile!, History = history });
	}

	public static (string AuthId, RBLePayload Payload)? GetRBLePayload( XElement? xml /*, string dateAttribute, string authorAttribute */ )
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

		return ((string)xml.xDataDef().Attribute( "id-auth" )!, new() { Profile = profile!, History = history })!;
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

	public async Task CalcEngineUtilities_LocalBatchCalc( IRibbonControl _ )
	{
		try
		{
			var helpers = application.GetWorkbook( Constants.FileNames.Helpers );
			if ( helpers == null && !File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.Helpers ) ) )
			{
				MessageBox.Show( "The Helpers workbook is missing.  Please download it before processing the workbook.", "Missing Helpers", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}

			var fileName = application.ActiveWorkbook.Name;
			var tabName = application.ActiveWorksheet().Name;
			var selection = DnaApplication.Selection;

			CalcEngineConfiguration ceConfig;
			application.ScreenUpdating = false;
			try
			{
				ceConfig = new DnaCalcEngineConfigurationFactory( fileName ).Configuration;
			}
			finally
			{
				selection.Select();
				application.ScreenUpdating = true;
			}

			var owner = new NativeWindow();
			owner.AssignHandle( new IntPtr( application.Hwnd ) );

			var windowConfig = GetWindowConfiguration( nameof( LocalBatch ) );

			using var loadInputTab = new LocalBatch(
				ceConfig,
				tabName,
				windowConfig
			);

			var info = loadInputTab.GetInfo( owner );

			if ( info == null ) return;

			SaveWindowConfiguration( nameof( LocalBatch ), info.WindowConfiguration );

			var ceBatchPath = Path.Combine( AddIn.ResourcesPath, "LocalBatch", $"{Path.GetFileNameWithoutExtension( WorkbookState.ManagementName )}.localbatch{Path.GetExtension( WorkbookState.ManagementName )}" );
			var helpersBatchPath = Path.Combine( AddIn.ResourcesPath, "LocalBatch", $"{Path.GetFileNameWithoutExtension( Constants.FileNames.Helpers )}.localbatch{Path.GetExtension( Constants.FileNames.Helpers )}" );

			Directory.CreateDirectory( Path.GetDirectoryName( ceBatchPath )! );
			foreach ( var f in new[] { ceBatchPath, helpersBatchPath } )
			{
				if ( File.Exists( f ) )
				{
					File.Delete( f );
				}
			}

			application.ActiveWorkbook.SaveCopyAs( ceBatchPath );
			helpers?.SaveCopyAs( helpersBatchPath );

			if ( helpers == null )
			{
				File.Copy( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.Helpers ), helpersBatchPath );
			}

			// Note, only way this works is if the following is true (however, I've asked for advice in Excel-DNA group https://groups.google.com/g/exceldna/c/iUBUHNyYg3g):
			// 1. Function needs to be 'async Task' instead of 'void' -> public async Task CalcEngineUtilities_LocalBatchCalc
			// 2. Need 'ExcelAsyncUtil.QueueAsMacro' call otherwise 'DnaApplication.Selection' fails.  Since Govert says (https://groups.google.com/g/exceldna/c/XN6cwgYGtIs) 
			//		that 'ExcelAsyncUtil.QueueAsMacro( async () => { } )' is not a valid pattern, I moved QueueAsMacro up into 'Ribbon_OnAction' like it was in
			//		Evolution version of addin.  'Ribbon_OnAction' calls button handlers via 'MemberInfo.Invoke', so not sure how that handles 'async Task' 
			//		handlers, but it seems to be working.
			// 3. Using this, means lose benefit of built in exception handling of RunRibbonTask :(


			// Task.Run( async () => 
			// {
			/* Sample for Excel-DNA group...
				try
				{
					var processingConfig = GetWindowConfiguration( nameof( Processing ) );
					using var processing = new Processing( "Local Batch Calculation", 1, 10, processingConfig );

					var result = await processing.ProcessAsync( async ( currentProgress, currentStatus, cancellationToken ) =>
					{
						currentStatus.Report( "Starting batch calculations..." );

						for ( var i = 0; i < 10; i++ )
						{
							await Task.Delay( 1000, cancellationToken );
							currentProgress.Report( -1 );
						}

						currentStatus.Report( "Finished batch calculations" );
					} );
				}
				catch ( Exception ex )
				{
					ExcelAsyncUtil.QueueAsMacro( () => MessageBox.Show( ex.Message ) );
				}
			*/
			// } );

			// RunRibbonTask( async () =>
			// {
			var processingConfig = GetWindowConfiguration( nameof( Processing ) );
			var batchManager = new BatchCalculations.BatchManager( info, ceBatchPath, helpersBatchPath, serviceProvider, updatesJwtInfo );
			
			var batchFile = application.GetWorkbook( Path.GetFileName( batchManager.ErrorFile ) );
			batchFile?.Close( false );
			batchFile = application.GetWorkbook( Path.GetFileName( info.OutputFile ) );
			batchFile?.Close( false );

			using var processing = new Processing( "Local Batch Calculation", batchManager.ProgressMax, processingConfig );

			var result = await processing.ProcessAsync( batchManager.RunBatchAsync );

			SaveWindowConfiguration( nameof( Processing ), result.WindowConfiguration );

			if ( result.Result == DialogResult.None )
			{
				MessageBox.Show( "Process failed. Please see Log Display for more information", "Local Batch Calculation", MessageBoxButtons.OK, MessageBoxIcon.Error );
				ExcelDna.Logging.LogDisplay.Show();
			}
			else
			{
				MessageBox.Show(
					result.Result == DialogResult.OK
						? $"Process complete. Ran {batchManager.TotalCalculations} calculation(s) with {batchManager.TotalErrors} error(s) in {batchManager.Elapsed.TotalMinutes:0.00} minutes."
						: $"Process cancelled after {batchManager.Elapsed.TotalMinutes:0.00} minutes.  Completed {batchManager.TotalCalculations} calculation(s) with {batchManager.TotalErrors} error(s).",
					"Local Batch Calculation",
					MessageBoxButtons.OK,
					result.Result == DialogResult.OK ? MessageBoxIcon.Information : MessageBoxIcon.Warning
				);
			}
			// } );
		}
		catch ( Exception ex )
		{
			ShowException( ex, $"Ribbon_OnAction {nameof( CalcEngineUtilities_LocalBatchCalc )}" );
		}
	}

	public void CalcEngineUtilities_ExportResultDocGenXml( IRibbonControl _ )
	{
		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			var (results, jsonPath) = ExportResultJsonData();

			var fileName = application.ActiveWorkbook.Name;
			var version = application.ActiveWorkbook.RangeOrNull<string>( "Version" )!;
			var xml = new XElement( "RBL",
				new XElement( "Profile",
					new XAttribute( "id-auth", "N/A" ),
					new XAttribute( "CalcEngineName", fileName ),
					new XAttribute( "CalcEngineVersion", version ),
					new XAttribute( "job-token", Guid.NewGuid().ToString() ),
					new XElement( "Data",
						results.ToEvolutionXml( version, fileName, "N/A", "N/A", "N/A", 0, DateTime.Now )
					)
				)
			);

			var profileXml = new XElement( "xDataDefs", new XElement( "xDataDef", new XElement( "FolderItems" ) ) );
			var docgenXml = xml.ToDocGenXml( results, calculationChartBuilder, profileXml );	

			var docgenPath = Path.ChangeExtension( jsonPath, ".xml" );

			docgenXml.Save( docgenPath );

			Process.Start( AddIn.Settings.TextEditor, $"\"{docgenPath}\"" );
		} );
	}

	public void CalcEngineUtilities_ExportResultJsonData( IRibbonControl _ )
	{
		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			var (_, path) = ExportResultJsonData();
			Process.Start( AddIn.Settings.TextEditor, $"\"{path}\"" );
		} );
	}

	public void CalcEngineUtilities_ConfigureHighCharts( IRibbonControl _ )
	{
		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			var (results, _) = ExportResultJsonData();

			var chartConfigurations = calculationChartBuilder.GetConfigurations( results ).ToArray();

			var hasCulture = !( (string?)results.Tables.FirstOrDefault( t => t.Name == "variable" )?.Rows.FirstOrDefault( r => (string?)r![ "@id" ] == "culture" )?[ "value" ] ?? "en-" ).StartsWith( "en-" );

			var script = new StringBuilder();

			script.AppendLine( @$"
const defaultTooltip = function(tooltipFormat, seriesFormats) {{
	return {{
		formatter: function() {{
			var s = '';
			var t = 0;

			$.each(this.points, function (i, point) {{
				if (point.y > 0 ) {{
					s += '<br/>' + point.series.name + ' : ' + String.localeFormat( '{{0:' + seriesFormats[ i ] + '}}', point.y );
					t += point.y;
				}}
			}});

			return tooltipFormat.replace('{{x}}', this.x)
						.replace('{{stackTotal}}', String.localeFormat( '{{0:' + seriesFormats[ 0 ] + '}}', t))
						.replace('{{seriesDetail}}', s);			
		}},
		shared: true
	}};
}};

function GetChartConfiguration(e,i) {{
	var t = '';
	for (var n in e) {{
		if (['data','categories','AsLinq'].indexOf(n) < 0 && (void 0!=i || 'series' != n) ) {{
			var r = void 0 != i ? i + '.' + n : n;
			var a = e[n];
			var o = typeof a;
			
			if( 'object' == o ) {{
				t += GetChartConfiguration(a,r);
			}} 
			else {{
				var g = 'string' == o
					? BTR.RBLe.Debugging.GetCsvValue(a)
					: 'function' == o
						? BTR.RBLe.Debugging.GetCsvValue(a.toString())
						:a;
				t += r + ', ' + g + '\r\n';
			}}
		}}
	}}
	return t;
}}
					" );

			if ( hasCulture )
			{
				script.AppendLine( @"
Highcharts.setOptions({
	yAxis: {
		labels: {
			formatter: function() {
				return String.localeFormat( ""{0:c0}"", this.value );
			}
		},
		stackLabels: {}
			formatter: function() {
				return String.localeFormat( ""{0:c0}"", this.total );
			}
		}
	}
});
				" );
			}

			foreach ( var chart in chartConfigurations )
			{
				script.AppendLine( $"var {chart.HtmlId}_options = {chart.Options.ToJsonString()};" );
				script.AppendLine( $"var {chart.HtmlId}_seriesFormats = {chart.SeriesFormats.ToJsonString()};" );

				if ( !string.IsNullOrEmpty( chart.ToolTipFormat ) )
				{
					script.AppendLine( $"{chart.HtmlId}_options.tooltip = defaultTooltip(\"{chart.ToolTipFormat}\", {chart.HtmlId}_seriesFormats)" );
				}
			}

			script.AppendLine( "function renderHighCharts() {" );

			foreach ( var chart in chartConfigurations )
			{
				script.AppendLine( $"	$('#{chart.HtmlId}').highcharts({chart.HtmlId}_options);" );
			}

			script.AppendLine( "}" );

			var chartDOMItems = string.Join( "",
				chartConfigurations.Select( c => $"<div class='panel panel-default'><div class='panel-heading'>{c.Name} Chart</div><div class='panel-body'><div id='{c.HtmlId}'></div><p><a href='#' class='config' data-chart='{c.HtmlId}'>Copy Configuration</a></p></div></div>" )
			);

			script.AppendLine( $@"
$(document).ready(function() {{
	$( '.col-xs-12' ).append( ""{chartDOMItems}"" );

	$('a.config').click(function() {{ 
		var chartId = $(this).data('chart'); 
		console.log('\r\n' + GetChartConfiguration(eval(chartId + '_options'))); 
		alert( 'View console log to copy key/value properties.' );
	}});

	renderHighCharts();
}});
			" );

			Clipboard.SetText( script.ToString() );
			MessageBox.Show( "Click OK to launch a JSFiddle session in your default browser.  The HighCharts configuration has been copied to you clipbard.  Please paste the content into the 'Javascript' window (lower left) and click on the Run button.", "Opening JSFiddle...", MessageBoxButtons.OK, MessageBoxIcon.Information );

			Process.Start( new ProcessStartInfo
			{
				FileName = "https://jsfiddle.net/p3en0wu6/",
				UseShellExecute = true // Required for .NET Core and .NET 5+
			} );
		} );
	}

	private (ResponseTab Results, string Path) ExportResultJsonData() 
	{
		var fileName = application.ActiveWorkbook.Name;
		var tabName = application.ActiveWorksheet().Name;
		// Haven't found C API equivalent to setting Saved property...
		var isSaved = application.ActiveWorkbook.Saved;
		// TODO: See why .RestoreSelection doesn't work here.
		var selection = DnaApplication.Selection;
		application.ScreenUpdating = false;

		try
		{
			var configuration = new DnaCalcEngineConfigurationFactory( fileName ).Configuration;
			using var calcEngine = new DnaCalcEngine( fileName, configuration );

			var parameters = new CalculationParameters
			{
				CalculationId = Guid.NewGuid(),
				CalcEngineInfo = new()
				{
					ResultTabs = new [] { tabName }
				},
				RequestInfo = new SingleRequest() 
				{
					AuthId = "LOCAL.DEBUG",
					CalcEngines = Array.Empty<RequestCalcEngine>(),
					TraceEnabled = true
				},
				Payload = new()
				{
					Profile = new(),
					History = new()
				}
			};

			var diagnosticTraceLogger = new DiagnosticTraceLogger();
			diagnosticTraceLogger.Start();

			var results = dnaCalculationService.ReadResults( calcEngine, parameters, diagnosticTraceLogger );

			var path = Path.Combine( AddIn.ResourcesPath, "ResultExports", $"{Path.GetFileNameWithoutExtension( fileName )}.{tabName}.json" );
			Directory.CreateDirectory( Path.GetDirectoryName( path )! );
			File.WriteAllText( path, results.ToJsonString( writeIndented: true, ignoreNulls: true, camelCase: true ) );

			if ( diagnosticTraceLogger.HasTrace )
			{
				ExcelDna.Logging.LogDisplay.Clear();
				foreach ( var t in diagnosticTraceLogger.Trace )
				{
					ExcelDna.Logging.LogDisplay.WriteLine( t.Replace( "\t", "    " ) );
				}
				ExcelDna.Logging.LogDisplay.Show();
			}

			return (results, path);
		}
		finally
		{
			selection.Select();
			application.ScreenUpdating = true;
			application.ActiveWorkbook.Saved = isSaved;
		}
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