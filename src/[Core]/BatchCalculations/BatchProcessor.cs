
using System.Xml;
using System.Xml.Linq;
using KAT.Camelot.Abstractions.RBLe;
using KAT.Camelot.Abstractions.RBLe.Calculations;
using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Domain.Services;
using KAT.Camelot.Domain.Telemetry;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.RBLe.Core.Calculations.SpreadsheetGear;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace KAT.Camelot.Extensibility.Excel.AddIn.BatchCalculations;

class BatchProcessor
{
	private readonly LocalBatchInfo localBatchInfo;

	public int TotalCalculations { get; private set; }
	public int TotalErrors { get; private set; }
	public bool IsUpdate { get; private set; }
	public XElement? UpdateConfiguration { get; private set; }

	private readonly XmlWriter? xmlWriter;
	private readonly TextWriter? csvWriter;
	private readonly TextWriter errors;
	private readonly SpreadsheetGearCalcEngine calcEngine;
	private readonly SpreadsheetGearCalcEngine helpers;
	private readonly SpreadsheetGearCalculationService calculationService;
	private readonly string outputFile;
	private readonly string errorOutputFile;

	public string[]? CsvHeaders { get; private set; }
	private string? csvTableName;
	private string[]? csvTableIds;
	private string[]? csvTableColumns;

	public BatchProcessor( 
		ServiceProvider serviceProvider, 
		JwtInfo updatesJwtInfo, 
		LocalBatchInfo localBatchInfo, 
		string ceBatchPath, 
		string helpersBatchPath, 
		byte[] ceContent, 
		byte[] helpersContent 
	)
	{
		this.localBatchInfo = localBatchInfo;

		var isXml = localBatchInfo.ExportType == ExportFormatType.Xml;
		outputFile = Path.Combine( Path.GetTempPath(), $"{Guid.NewGuid()}.{( isXml ? "xml" : "csv" )}" );
		errorOutputFile = outputFile + ".errors.csv";

		if ( isXml )
		{
			xmlWriter = XmlWriter.Create( outputFile, new XmlWriterSettings { Async = true } );
			xmlWriter.WriteStartElement( "Results" );
		}
		else
		{
			csvWriter = new StreamWriter( outputFile );
		}
		errors = new StreamWriter( errorOutputFile );

		var configuration = new SpreadsheetGearCalcEngineConfigurationFactory( ceBatchPath, "1.0", ceContent ).Configuration;
		calcEngine = new SpreadsheetGearCalcEngine( ceBatchPath, "1.0", ceContent, configuration );
		helpers = new SpreadsheetGearCalcEngine( helpersBatchPath, "1.0", helpersContent );

		calculationService = new SpreadsheetGearCalculationService(
			serviceProvider.GetRequiredService<IHttpClientFactory>(), 
			new FakeEmailService(), 
			new FakeTextService(), 
			updatesJwtInfo, 
			serviceProvider.GetRequiredService<ILogger<CalculationSourceContext>>()
		);
	}

	public async Task RunAsync( string authId, RBLePayload payload, CancellationToken cancellationToken )
	{
		var diagnosticTraceLogger = new DiagnosticTraceLogger();
		diagnosticTraceLogger.Start();

		var parameters = new CalculationParameters
		{
			CalculationId = Guid.NewGuid(),

			ClearGlobalTables = false,
			ClearInputs = false,

			CalcEngineInfo = new()
			{
				InputTab = localBatchInfo.InputTab,
				ResultTabs = new [] { localBatchInfo.ResultTab }
			},

			RequestInfo = new SingleRequest() 
			{
				AuthId = authId,
				CalcEngines = Array.Empty<RequestCalcEngine>(),
				TraceEnabled = false
			},

			Payload = payload
		};

		try
		{
			var responseTab = ( await calculationService.CalculateAsync( calcEngine, helpers, parameters, diagnosticTraceLogger, cancellationToken ) ).First();
			TotalCalculations++;			

			await WriteResultsAsync( authId, responseTab, cancellationToken );
		}
		catch ( Exception ex ) when ( ex is not OperationCanceledException )
		{
			TotalErrors++;
			errors.WriteLine( new[] { authId, ex.Message, ex.StackTrace }.GetCsvLine() );
			throw;
		}
	}

	private async Task WriteResultsAsync( string authId, ResponseTab responseTab, CancellationToken cancellationToken ) 
	{
		if ( localBatchInfo.ExportType == ExportFormatType.Xml )
		{
			var isUpdate = string.Compare( responseTab.Type, "Update", true ) == 0;

			var profileUpdates = responseTab.DataUpdates?.Fields!;
			var historyUpdates = responseTab.DataUpdates?.HistoryUpdates!;

			if ( isUpdate && UpdateConfiguration == null )
			{
				var historyTypes = historyUpdates.Select( h => h.Type ).Distinct();

				UpdateConfiguration =
					new XElement( "Configuration",
						new XElement( "Profile", profileUpdates.Select( f => new XElement( f.Name ) ) ),
						historyTypes.Select( h =>
							new XElement( "HistoryData",
								new XAttribute( "type", h ),
								new XElement( "index" ),
								historyUpdates.First( r => r.Type == h ).Fields.Where( f => f.Name != "index" ).Select( f => new XElement( f.Name ) )
							)
						)
					);
			}

			if ( isUpdate )
			{
				xmlWriter!.WriteStartElement( "xDataDef" );

				xmlWriter.WriteAttributeString( "id-auth", authId );

				if ( profileUpdates.Any() )
				{
					xmlWriter.WriteStartElement( "Profile" );
					profileUpdates
						.Where( f => !f.Delete )
						.Select( f => new XElement( f.Name, f.Value ) )
						.ForAll( f => f.WriteTo( xmlWriter ) );
					xmlWriter.WriteEndElement();
				}

				if ( historyUpdates.Any() )
				{
					xmlWriter.WriteStartElement( "HistoryData" );

					historyUpdates
						.Select( h =>
							new XElement( "HistoryItem",
								new XAttribute( "hisType", h.Type ),
								new XAttribute( "index", h.Index ),
								h.Fields.Where( f => !f.Delete ).Select( f => new XElement( f.Name, f.Value ) )
							)
						).ForAll( h => h.WriteTo( xmlWriter ) );

					xmlWriter.WriteEndElement();
				}

				xmlWriter.WriteEndElement();
			}
			else
			{
				var result =
					new XElement( "Profile",
						new XAttribute( "id-auth", authId ),
						new XAttribute( "CalcEngineName", Path.GetFileNameWithoutExtension( calcEngine.FileName ) ),
						new XAttribute( "CalcEngineVersion", calcEngine.Version ),
						new XAttribute( "job-token", Guid.NewGuid() ),
						new XElement( "Data",
							responseTab.ToEvolutionXml( calcEngine.Version, Path.GetFileName( calcEngine.FileName ), helpers.Version, "1.0", Constants.FileNames.GlobalTables, 0, DateTime.Now )
						)
					);

				await result.WriteToAsync( xmlWriter!, cancellationToken );
			}
		}
		else
		{
			if ( string.IsNullOrEmpty( csvTableName ) )
			{
				var csvConfigTable = responseTab.Tables.First();
				
				csvTableName = csvConfigTable.Name;
				csvTableColumns = csvConfigTable.Specification!.Columns.Select( c => c.Name ).ToArray();
				csvTableIds = localBatchInfo.ExportType == ExportFormatType.CsvTransposed
					? csvConfigTable.Rows.Select( r => (string)r![ "id" ]! ).ToArray()
					: null;
				CsvHeaders = localBatchInfo.ExportType == ExportFormatType.CsvTransposed
					? csvTableIds!.SelectMany( i => csvTableColumns.Select( c => $"{i}_{c}" ) ).ToArray()
					: csvTableColumns;
			}

			var exportTable = responseTab.Tables.FirstOrDefault( t => t.Name == csvTableName );

			if ( exportTable == null )
			{
				// TODO: Should maybe have user specify which table to export in case first calc doesn't return it, should I write 'empty' line in csv?
				return;
			}
		
			if ( localBatchInfo.ExportType == ExportFormatType.Csv )
			{
				foreach ( var row in exportTable.Rows )
				{
					await csvWriter!.WriteLineAsync( new [] { authId }.Concat( csvTableColumns!.Select( c => (string?)row![ c ] ) ).GetCsvLine() );
				}
			}
			else
			{
				var values =
					csvTableIds!.SelectMany( i => {
						var row = exportTable.Rows.FirstOrDefault( r => (string?)r![ "id" ] == i );
						var vals = csvTableColumns!.Select( c =>
							(string?)row?[ c ]
						);
						return vals;
					} );
				
				await csvWriter!.WriteLineAsync( new [] { authId }.Concat( values ).GetCsvLine() );
			}
		}
	}

	public void Close()
	{
		calcEngine.Dispose();
		helpers.Dispose();
	}

	public async Task MoveResultsToAsync( XmlWriter xmlResults, CancellationToken cancellationToken )
	{
		if ( xmlWriter == null ) return;

		xmlWriter.WriteEndElement();
		xmlWriter.Close();
		
		using ( var reader = XmlReader.Create( outputFile ) )
		{
			while ( reader.Read() )
			{
				if ( reader.LocalName == ( IsUpdate ? "xDataDef" : "Profile" ) )
				{
					await xmlResults.WriteNodeAsync( reader.ReadSubtree(), true );
					cancellationToken.ThrowIfCancellationRequested();
				}
			}
		}

		File.Delete( outputFile );
	}

	public async Task MoveResultsToAsync( StreamWriter csvResults, CancellationToken cancellationToken )
	{
		if ( csvWriter == null ) return;

		csvWriter.Close();

		using ( var fs = new StreamReader( outputFile ) )
		{
			while ( !fs.EndOfStream )
			{
				await csvResults.WriteLineAsync( fs.ReadLine() );
				cancellationToken.ThrowIfCancellationRequested();
			}
		}

		File.Delete( outputFile );
	}

	public async Task MoveErrorsToAsync( StreamWriter errorResults, CancellationToken cancellationToken )
	{
		errors.Close();

		if ( TotalErrors != 0 )
		{
			using var fs = new StreamReader( errorOutputFile );
			
			string? line;
			while ( !string.IsNullOrEmpty( line = fs.ReadLine() ) )
			{
				await errorResults.WriteLineAsync( line );
				cancellationToken.ThrowIfCancellationRequested();
			}
		}

		File.Delete( errorOutputFile );
	}

	public void SaveCalcEngine( string errorCalcEngineFile )
	{
		using var fs = new FileStream( errorCalcEngineFile, FileMode.Create );
		calcEngine.SaveToStream( fs );
	}
}