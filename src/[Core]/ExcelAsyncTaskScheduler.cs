
using ExcelDna.Integration;

namespace KAT.Extensibility.Excel.AddIn;

// https://groups.google.com/g/exceldna/c/9OkHWILuFMo/m/RpilXElgAQAJ
// https://github.com/Excel-DNA/Samples/blob/master/AsyncAwaitMacro/ExcelAsyncTask.cs
public class ExcelAsyncTaskScheduler : TaskScheduler
{
	public static Task Run( Func<Task> function )
	{
		return Task<Task>.Factory.StartNew(
			function,
			CancellationToken.None,
			TaskCreationOptions.DenyChildAttach,
			Instance
		).Unwrap();
	}

	public static Task<T> Run<T>( Func<Task<T>> function )
	{
		return Task<Task<T>>.Factory.StartNew(
			function,
			CancellationToken.None,
			TaskCreationOptions.DenyChildAttach,
			Instance
		).Unwrap();
	}

	void PostCallback( object? obj )
	{
		var task = (Task)obj!;
		TryExecuteTask( task );
	}

	protected override void QueueTask( Task task ) => ExcelAsyncUtil.QueueAsMacro( PostCallback, task );
	protected override bool TryExecuteTaskInline( Task task, bool taskWasPreviouslyQueued ) => false;
	protected override IEnumerable<Task> GetScheduledTasks() => Enumerable.Empty<Task>();
	internal static ExcelAsyncTaskScheduler Instance = new ExcelAsyncTaskScheduler();
}