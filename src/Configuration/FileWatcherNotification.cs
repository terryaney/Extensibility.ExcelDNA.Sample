using System.Collections.Concurrent;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

/// <summary>
/// FileSystemWatcher notifications are capable of happening 'multiple' times for a single 'action'.  For example, if Notepad saves a file,
/// you might not get a single 'Changed' event when everything is 'done', you might get multiple 'Changed' events.  Similarily, "when a file is 
/// moved from one directory to another, several OnChanged and some OnCreated and OnDeleted events might be raised." (from MS Docs) So this class 
/// mitigates that by having an internal timer that starts/restarts on each event.  So once the timer is created (first event), if no other events
/// occur for notificationDelay milliseconds, then, and only then, is the original event raised to the caller along with the associated 
/// FolderConfiguration element.
/// </summary>
/// <remarks>
/// See https://asp-blogs.azurewebsites.net/ashben/31773 - see 'Events being raised multiple times'
/// </remarks>
public class FileWatcherNotification
{
	private readonly ConcurrentDictionary<string, System.Timers.Timer> notificationTimers = new();
	private readonly FileSystemWatcher watcher;
	private readonly int notificationDelay;
	private readonly string path;
	private readonly Action<FileSystemEventArgs> action;

	public FileWatcherNotification( int notificationDelay, string path, string filter, Action<FileSystemEventArgs> action )
	{
		this.notificationDelay = notificationDelay;
		this.path = path;
		this.action = action;

		watcher = new FileSystemWatcher( path, filter ) { EnableRaisingEvents = true };
		watcher.Changed += Watcher_Changed;
	}

	private void Watcher_Changed( object sender, FileSystemEventArgs e )
	{
		var timer = notificationTimers.GetOrAdd(
			e.FullPath,
			p =>
			{
				var t = new System.Timers.Timer( notificationDelay );

				t.Elapsed += ( sender, args ) =>
				{
					t.Enabled = false;
					action( e );
				};

				return t;
			}
		);

		timer.Stop();
		timer.Start();
	}

	public void Disable() => watcher.EnableRaisingEvents = false;
	public void Enable() => watcher.EnableRaisingEvents = true;

	public void Changed( string name )
	{
		Enable();
		Watcher_Changed( this, new FileSystemEventArgs( WatcherChangeTypes.Changed, path, name ) );
	}
}
