using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core;
using Microsoft.Extensions.DependencyInjection;

namespace KAT.Camelot.Extensibility.Excel.AddIn.BatchCalculations;

class BatchManager
{
	private readonly LocalBatchInfo localBatchInfo;
	private readonly string ceBatchPath;
	private readonly string helpersBatchPath;
	private readonly long fileSize;
	private readonly int progressSizeReductionFactor;
	private readonly ServiceProvider serviceProvider;
	private readonly JwtInfo updatesJwtInfo;

	public string ErrorFile { get; }

	private int errorCalcEnginesSaved;

	public int TotalCalculations { get; private set; }
	public int TotalErrors { get; private set; }
	public TimeSpan Elapsed { get; private set; }
	public int ProgressMax => localBatchInfo.InputRows ?? (int)( fileSize / progressSizeReductionFactor );

	public BatchManager( 
		LocalBatchInfo localBatchInfo, 
		string ceBatchPath, 
		string helpersBatchPath, 
		ServiceProvider serviceProvider,
		JwtInfo updatesJwtInfo
	)
	{
		this.localBatchInfo = localBatchInfo;
		this.ceBatchPath = ceBatchPath;
		this.helpersBatchPath = helpersBatchPath;

		fileSize = new FileInfo( localBatchInfo.InputFile ).Length;
		progressSizeReductionFactor = 1;
		while ( ( fileSize / progressSizeReductionFactor ) > int.MaxValue )
		{
			progressSizeReductionFactor *= 10;
		}

		this.serviceProvider = serviceProvider;
		this.updatesJwtInfo = updatesJwtInfo;

		ErrorFile = Path.Combine( Path.GetDirectoryName( localBatchInfo.OutputFile )!, $"{Path.GetFileNameWithoutExtension( localBatchInfo.OutputFile )}.errors.csv" );
	}

	public async Task RunBatchAsync( IProgress<int> currentProgress, IProgress<string> currentStatus, CancellationToken cancellationToken )
	{
		currentStatus.Report( "Preparing batch calculation..." );

		var ceContent = await File.ReadAllBytesAsync( ceBatchPath, cancellationToken );
		var helpersContent = await File.ReadAllBytesAsync( helpersBatchPath, cancellationToken );

		var stopwatch = Stopwatch.StartNew();
		var maxConcurrent = Math.Min( 2, Environment.ProcessorCount );
		var processors = new ConcurrentBag<BatchProcessor>();

		if ( File.Exists( ErrorFile ) )
		{
			File.Delete( ErrorFile );
		}

		try
		{
			currentStatus.Report( $"Running calculations {ceContent.Length}, {helpersContent.Length}..." );

			var allowedErrorCalcEngines = Math.Min( localBatchInfo.ErrorCalcEngines ?? 0, 10 );

			await Parallel.ForEachAsync(
				ParseProfilesAsync( cancellationToken ),
				new ParallelOptions { MaxDegreeOfParallelism = Math.Min( localBatchInfo.InputRows ?? maxConcurrent, maxConcurrent ) },
				async ( profile, token ) =>
				{
					if ( !processors.TryTake( out var processor ) )
					{
						processor = new BatchProcessor( 
							serviceProvider, 
							updatesJwtInfo, 
							localBatchInfo, 
							ceBatchPath,
							helpersBatchPath,
							ceContent, 
							helpersContent 
						);
					}

					try
					{
						await processor.RunAsync( profile.AuthId, profile.Payload, cancellationToken );
						currentProgress.Report( localBatchInfo.InputRows != null ? -1 : (int)( profile.StreamPosition / progressSizeReductionFactor ) );
					}
					catch ( Exception ex ) when ( ex is not OperationCanceledException )
					{						
						if ( ex is CalcEngineProcessingException && Interlocked.Increment( ref errorCalcEnginesSaved ) < allowedErrorCalcEngines )
						{
							var errorCalcEngineFile = Path.Combine( Path.GetDirectoryName( ceBatchPath )!, $"{Path.GetFileNameWithoutExtension( ceBatchPath )}.Error.{profile.AuthId}{Path.GetExtension( ceBatchPath )}" );
							processor.SaveCalcEngine( errorCalcEngineFile );
						}
					}
					finally
					{
						processors.Add( processor );
					}
				}
			);

			stopwatch.Stop();
			currentStatus.Report( "Finished batch calculations, saving output..." );

			processors.ForAll( p => p.Close() );

			await WriteResultsAsync( processors, cancellationToken );
		}
		finally
		{
			stopwatch.Stop();

			TotalCalculations = processors.Sum( p => p.TotalCalculations );
			TotalErrors = processors.Sum( p => p.TotalErrors );

			Elapsed = stopwatch.Elapsed;
		}
	}

	private async Task WriteResultsAsync( ConcurrentBag<BatchProcessor> processors, CancellationToken cancellationToken )
	{
		if ( processors.Any( p => p.TotalErrors > 0 ) )
		{
			using var sw = new StreamWriter( ErrorFile );

			await sw.WriteLineAsync( "AuthID,Exception,Trace" );

			foreach( var p in processors )
			{
				await p.MoveErrorsToAsync( sw, cancellationToken );
			}
		}

		if ( localBatchInfo.ExportType == ExportFormatType.Xml )
		{
			var settings = new XmlWriterSettings
			{
				Encoding = Encoding.UTF8,
				Indent = Debugger.IsAttached
			};

			using var xw = XmlWriter.Create( localBatchInfo.OutputFile, settings );

			if ( processors.Any( p => p.IsUpdate ) )
			{
				xw.WriteStartElement( "xDataDefs" );
				xw.WriteAttributeString( "TotalRows", processors.Sum( p => p.TotalCalculations - p.TotalErrors ).ToString() );

				processors.First( p => p.IsUpdate ).UpdateConfiguration!.WriteTo( xw );
				
				foreach( var p in processors )
				{
					await p.MoveResultsToAsync( xw, cancellationToken );
				}

				xw.WriteEndElement();
			}
			else
			{
				xw.WriteStartElement( "Results" );

				foreach( var p in processors )
				{
					await p.MoveResultsToAsync( xw, cancellationToken );
				}

				xw.WriteEndElement();
			}
		}
		else
		{
			using var sw = new StreamWriter( localBatchInfo.OutputFile );

			if ( processors.Any( p => p.IsUpdate ) )
			{
				await sw.WriteLineAsync(
					( localBatchInfo.ExportType == ExportFormatType.Csv ? new[] { "AuthID", "id" } : new[] { "AuthID" } )
					.Concat( processors.FirstOrDefault( p => p.CsvHeaders != null )?.CsvHeaders ?? Enumerable.Empty<string>() )
					.GetCsvLine()
				);

				foreach( var p in processors )
				{
					await p.MoveResultsToAsync( sw, cancellationToken );
				}
			}
			else
			{
				await sw.WriteLineAsync( "AuthID,Response" );

				foreach( var p in processors )
				{
					await p.MoveResultsToAsync( sw, cancellationToken );
				}
			}
		}
	}

	private async IAsyncEnumerable<BatchProfile> ParseProfilesAsync( [EnumeratorCancellation] CancellationToken cancellationToken )
	{
		var processed = 0;
		var filterByXPath = localBatchInfo.Filter;
		var hasFilter = !string.IsNullOrEmpty( filterByXPath );
		var limitBatchSize = localBatchInfo.InputRows;

		using var fileStream = new FileStream( localBatchInfo.InputFile, FileMode.Open, FileAccess.Read );

		if ( string.Compare( Path.GetExtension( localBatchInfo.InputFile ), ".xml", true ) == 0 )
		{
			using var reader = XmlReader.Create( fileStream );
			while( reader.Read() )
			{
				if ( reader.LocalName == "xDataDef" )
				{
					if ( cancellationToken.IsCancellationRequested || ( limitBatchSize != null && processed >= limitBatchSize.Value ) )
					{
						yield break;
					}

					var data = XElement.Load( reader.ReadSubtree() );

					if ( !hasFilter || data.XPathSelectElement( filterByXPath! ) != null )
					{
						processed++;
						var inputData = Ribbon.GetRBLePayload( data )!;
						yield return new BatchProfile 
						{ 
							AuthId = inputData.Value.AuthId,
							Payload = inputData.Value.Payload,
							StreamPosition = fileStream.Position 
						};
					}
				}
			}
		}
		else
		{
			await foreach ( var json in JsonSerializer.DeserializeAsyncEnumerable<JsonObject>( fileStream, cancellationToken: cancellationToken ) )
			{
				if ( cancellationToken.IsCancellationRequested || ( limitBatchSize != null && processed >= limitBatchSize.Value ) )
				{
					yield break;
				}

				var (authId, payload) = Ribbon.GetRBLePayload( json! );

				if ( !hasFilter || payload.ToxDSXml( authId ).XPathSelectElement( filterByXPath! ) != null )
				{
					processed++;
					yield return new BatchProfile 
					{ 
						AuthId = authId,
						Payload = payload,
						StreamPosition = fileStream.Position 
					};
				}
			}
		}
	}
}