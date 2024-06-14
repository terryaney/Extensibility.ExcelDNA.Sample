using System.Diagnostics;

namespace KAT.Camelot.Extensibility.Excel.AddIn.BatchCalculations;

class BatchManager
{
	private readonly LocalBatchInfo localBatchInfo;
	private readonly string ceBatchPath;
	private readonly string helpersBatchPath;
	private readonly int progressSizeReductionFactor;

	public int TotalCalculations { get; private set; }
	public int TotalErrors { get; private set; }
	public TimeSpan Elapsed { get; private set; }

	public BatchManager( LocalBatchInfo localBatchInfo, string ceBatchPath, string helpersBatchPath, int progressSizeReductionFactor )
	{
		this.localBatchInfo = localBatchInfo;
		this.ceBatchPath = ceBatchPath;
		this.helpersBatchPath = helpersBatchPath;
		this.progressSizeReductionFactor = progressSizeReductionFactor;
	}

	public async Task RunBatchAsync( IProgress<int> currentProgress, IProgress<string> currentStatus, CancellationToken cancellationToken )
	{
		var stopwatch = Stopwatch.StartNew();

		currentStatus.Report( "Preparing batch calculation..." );

		var ceContent = await File.ReadAllBytesAsync( ceBatchPath, cancellationToken );
		var helpersContent = await File.ReadAllBytesAsync( helpersBatchPath, cancellationToken );

		currentStatus.Report( $"Running calculations {ceContent.Length}, {helpersContent.Length}..." );

		for ( var i = 0; i < 10; i++ )
		{
			await Task.Delay( 1000, cancellationToken );
			currentProgress.Report( localBatchInfo.InputRows != null ? -1 : i );
			TotalCalculations++;
		}

		currentStatus.Report( "Finished batch calculations" );

		stopwatch.Stop();

		Elapsed = stopwatch.Elapsed;
	}
}