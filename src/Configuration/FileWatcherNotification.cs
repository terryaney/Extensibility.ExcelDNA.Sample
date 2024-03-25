namespace KAT.Extensibility.Excel.AddIn;

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
	private readonly System.Timers.Timer timer;
	private readonly FileSystemWatcher watcher;

	private readonly string path;
	private readonly string name;
	private FileSystemEventArgs fileSystemEventArgs = null!;

	public FileWatcherNotification( int notificationDelay, string path, string name, Action<FileSystemEventArgs> action )
	{
		watcher = new FileSystemWatcher( path, name ) { EnableRaisingEvents = true };
		watcher.Changed += watcher_Changed;

		timer = new( notificationDelay );
		timer.Elapsed += ( sender, args ) =>
		{
			timer.Enabled = false;
			action( fileSystemEventArgs );
		};
		this.path = path;
		this.name = name;
	}

	private void watcher_Changed( object sender, FileSystemEventArgs e )
	{
		timer.Stop();
		fileSystemEventArgs = e;
		timer.Start();
	}

	public void Start() => timer.Start();
	public void Stop() => timer.Stop();

	public void Changed() => watcher_Changed( this, new FileSystemEventArgs( WatcherChangeTypes.Changed, path, name ) );	
}
