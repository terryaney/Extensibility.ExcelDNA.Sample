# Excel-DNA Add-In

An Excel-DNA add-in for the KAT teams that uses many of the features provided by Excel-DNA along with how I overcame many, not so obvious, obstacles.

## TODO

1. Implement ribbon 'state'
	1. Get 'State Coded'
	1. Ribbon_GetContent - is hacked to get it working
1. Implement Ribbon handlers
	1. exportMappedxDSData - Rename this better after you figure out what it is doing
	1. [Using Windows Forms](https://groups.google.com/g/exceldna/c/84IIhcdAPRk/m/8cRtFOvvAAAJ)	
1. [Custom Intellisense/Help Info](https://github.com/Excel-DNA/IntelliSense/issues/21) - read this and linked topic to see what's possible
	1. https://github.com/Excel-DNA/Tutorials/blob/master/SpecialTopics/IntelliSenseForVBAFunctions/README.md
1. [Path of Xll](https://groups.google.com/g/exceldna/c/1rScvDdeVOk) - `XlCall.Excel( XlCall.xlGetName )` to get the name of the add-in
1. [Possible Async Information](https://github.com/Excel-DNA/Samples/blob/master/Registration.Sample/AsyncFunctionExamples.cs)
	1. https://excel-dna.net/docs/guides-advanced/performing-asynchronous-work
	1. https://excel-dna.net/docs/tips-n-tricks/creating-a-threaded-modal-dialog
	1. Excel-DNA/Samples/Archive/Async/NativeAsyncCancellation.dna
1. BTR* functions...
	1. https://excel-dna.net/docs/guides-advanced/dynamic-delegate-registration - possible dynamic creation instead of having to create functions for each item and passing through?  Reference the SSG assembly detect custom functions.
1. Look for email/thread about excel not shutting down properly

## Development Environment

All the Excel-DNA samples seem to make the assumption that Visual Studio will be the IDE of choice.  I prefer to use Visual Studio Code.  This section will describe how I set up my development environment to work with Excel-DNA.


### launch.json/tasks.json Configuration To Enable Debugging

**launch.json**

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Excel.AddIn",
            "type": "coreclr",
            "request": "launch",
            "preLaunchTask": "debug",
            "logging": {
                "moduleLoad": false
            },
            "program": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
            "args": ["/x", "${workspaceFolder}\\src\\bin\\Debug\\net7.0-windows\\KAT.Extensibility.xll"],
            "cwd": "${workspaceFolder}\\src\\bin\\Debug\\net7.0-windows",
            "console": "internalConsole",
            "stopAtEntry": false
        }
    ]
}
```

**tasks.json**

```json
{
    "version": "2.0.0",
    "tasks": [
		{
			"label": "debug",
			"hide": true,
			"command": "dotnet",
			"type": "process",
			"presentation": {
				"clear": true
			},
			"args": [
				"build",
				"${workspaceFolder}\\src\\Excel.AddIn.csproj",
				"/property:GenerateFullPaths=true",
				"/consoleloggerparameters:NoSummary"
			],
			"problemMatcher": "$msCompile"
		}
    ]
}
```

### Intellisense in Ribbon.xml

My add-in has a CustomUI ribbon and to enable intellisense in the `Ribbon.xml` file, I had to add the following to the `settings.json` file:

```json
{
	"xml.fileAssociations": [
		{
			"pattern": "Ribbon.xml",
			"systemId": "https://raw.githubusercontent.com/Excel-DNA/ExcelDna/master/Distribution/XmlSchemas/customUI.xsd"
		}
	]
}
```

## Features

### Ribbon Organization

The use of the `IRibbonUI` in the KAT Tools add-in is quite extensive.  There is state management of CustomUI elements via ribbon events, CustomUI element handlers, dynamic menus, and dynamic images to name a few.  In this section I will describe some of the challenges I faced with `IRibbonUI` and how I overcame them.

**Helpful Documenation Links**

1. [CustomUI Reference](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43)
1. [imageMso Reference](https://codekabinett.com/download/Microsoft-Office-2016_365-imageMso-Gallery.pdf)


**Managing the Sheer Amount of Code**

Given the amount of methods I had to implement to provide all the required functionality, the lines of code became quite overwhelming (at least given the way I organized the code).  To help alleviate this, I used [partial classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods#partial-classes) just as an organizational tool.  This made it easier to find and maintain the code for me.  Additionally, to make this separation easier to manage in the (Solution) Explorer side bar, I would suggest to enable file nesting.  To enable file nesting in Visual Studio Code, add the following to your `settings.json` file:

```json
{
	"explorer.fileNesting.patterns": {
		"*.cs": "${capture}.*.cs"
	}
}
```

I also used method prefixes that matched the CustomUI `group.id` as well to make code navigation easier (via `CTRL+T` keyboard shortcut).  For example, for my group with an `id` of `Navigation`, the methods all have the prefix `Navigation_`.


### ExcelIntegration.RegisterUnhandledExceptionHandler

I register a global exception handler to log diagnostic information to the `ExcelDna.Logging.LogDisplay` window and update a ribbon image with a badge count.  In the diagnostic information, I wanted to display the address and formula of the offending cell.  Since the error handler runs on Excel's calculation thread, but directly converting the `ExcelReference` to an address can't (easily) be done in this context, so `ExcelAsyncUtil.QueueAsMacro` is required to register a delegate to run in a safe context on the main thread, from any other thread or context (i.e. when the calculation completes).  This [conversation](https://groups.google.com/d/msg/exceldna/cHD8Tx56Msg/MdPa2PR13hkJ) explains why `QueueAsMacro` is required for other `XlCall` methods.

For our add-in, we wanted to have a badge count on a ribbon image to indicate the number of formula errors that had occurred during calculations.  This is tracked via the `auditShowLogBadgeCount` variable below.  This variable is also used before saving a workbook to determine if the error log should be displayed.  

```csharp

// IExcelAddIn implementation
public void AutoOpen() => ExcelIntegration.RegisterUnhandledExceptionHandler( UnhandledExceptionHandler );

private object UnhandledExceptionHandler( object exception )
{
	var caller = ExcelApi.GetCaller();

	ExcelAsyncUtil.QueueAsMacro( () =>
	{
		Ribbon.CurrentRibbon.LogFunctionError( caller, exception );
	} );

	return ExcelError.ExcelErrorValue;
}

// Excel Ribbon Implementation
private readonly Microsoft.Office.Interop.Excel.Application application;

public Ribbon()
{
	application = ( ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application )!;
}

public override void OnConnection( object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom )
{
	base.OnConnection( Application, ConnectMode, AddInInst, ref custom );

	application.WorkbookOpen += Application_WorkbookOpen;
	application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
}

public override void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
{
	base.OnDisconnection( RemoveMode, ref custom );

	application.WorkbookOpen -= Application_WorkbookOpen;
	application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
}

private void Application_WorkbookOpen( MSExcel.Workbook Wb )
{
	// Clear error info whenever a new workbook is opened.  Currenly, only show any 
	// errors after a cell is calculated.  Could call application.Calculate() to force everything
	// to re-evaluate, but that could be expensive, so for now, not doing it, the function log display
	// is just helpful information for CalcEngine developer to 'clean' up their formulas.
	auditShowLogBadgeCount = 0;
	cellsInError.Clear();
	ExcelDna.Logging.LogDisplay.Clear();
}

private void Application_WorkbookBeforeSave( MSExcel.Workbook Wb, bool SaveAsUI, ref bool Cancel )
{
	// If any errors, we want to show them first before the save, then user can hit save again if needed
	if ( auditShowLogBadgeCount > 0 )
	{
		RBLe_ShowLog( null );
		Cancel = true;
	}
}

private readonly ConcurrentDictionary<string, string?> cellsInError = new();

public void LogFunctionError( ExcelReference caller, object exception )
{
	var address = caller.GetAddress();
	var formula = caller.GetFormula();

	var reportError = !cellsInError.TryGetValue( address, out var failedFormula ) || failedFormula != formula;
	cellsInError[ address ] = formula;

	// Only report error if not already reported once for current formula/worksheet
	if ( reportError )
	{
		var message = $"Error: {address} {formula ?? "unavailable"}{Environment.NewLine}{exception}";
	
		ExcelDna.Logging.LogDisplay.RecordLine( message );

		auditShowLogBadgeCount++;
		ribbon.InvalidateControl( "auditShowLog" );
	}		
}

public void RBLe_ShowLog( IRibbonControl? _ )
{
	ExcelDna.Logging.LogDisplay.Show();
	auditShowLogBadgeCount = 0;
	ribbon.InvalidateControl( "auditShowLog" );
}

private int auditShowLogBadgeCount;
public Bitmap Ribbon_GetImage( IRibbonControl control )
{
	switch ( control.Id )
	{
		case "auditShowLog":
		{
			using var ms = new MemoryStream( auditShowLogImage );

			var img = System.Drawing.Image.FromStream( ms );

			if ( auditShowLogBadgeCount > 0 )
			{
				var flagGraphics = Graphics.FromImage( img );
				flagGraphics.FillEllipse(
					new SolidBrush( Color.FromArgb( 242, 60, 42 ) ),
					new Rectangle( 11, 0, 19, 19 )
				);
				flagGraphics.DrawString(
					auditShowLogBadgeCount.ToString(),
					new Font( FontFamily.GenericSansSerif, 6, FontStyle.Bold ),
					Brushes.White,
					x: auditShowLogBadgeCount < 10 ? 16 : 13,
					y: 3 
				);
			}

			return (Bitmap)img;
		}

		default: throw new ArgumentOutOfRangeException( nameof( control.Id ), $"The id {control.Id} does not support custom image generation." );
	}
}


// Utility class wrapping Excel C API calls
public static class ExcelApi
{
	public static ExcelReference GetCaller() => (ExcelReference)XlCall.Excel( XlCall.xlfCaller );

	public static string? GetFormula( this ExcelReference cell )
	{
		var formula = (string)XlCall.Excel( XlCall.xlfGetCell, 6, cell );
		return !string.IsNullOrEmpty( formula ) ? formula : null;
	}

	public static string GetAddress( this ExcelReference? reference )
	{
		try
		{
			var address = (string)XlCall.Excel( XlCall.xlfReftext, reference, true /* true - A1, false - R1C1 */ );
			return address;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"GetAddress failed.  reference.RowFirst:{reference?.RowFirst}, reference.RowLast:{reference?.RowLast}, reference.ColumnFirst:{reference?.ColumnFirst}, reference.ColumnLast:{reference?.ColumnLast}", ex );
		}
	}
}
```

### appsettings.json Support

The KAT add-in requires support for user settings and the most convenient way to provide that functionality was simply by leveraging an `appsettings.json` file.  In the previous .NET Framework version, I used an `app.config` file which was distributed as `*.xll.config` and `*64.xll.config`.

To enable this support:

1. Output the `appsettings.json` file to the output directory during build so that it could be used during debugging.  Not displayed here, but additionally, a default `appsettings.json` file is included in the project distribution, but is based on whatever mechanism you choose to distribute your add-in.
1. Read and monitor the `appsettings.json` file for changes.
1. Access settings throughout the add-in.

#### Output `appsettings.json` File

Simply add the following to the `.csproj` file and the `appsettings.json` file will be copied to the output directory during build.

```xml
<ItemGroup>
	<Content Include="appsettings.json">
		<CopyToOutputDirectory>Always</CopyToOutputDirectory>
	</Content>
</ItemGroup>
```

#### Read and Monitor `appsettings.json` File

This was probably the trickiest part of the process.  As Excel-DNA documentation has stated, it does not want to include Dependency Injection into the project.  This means that the `IConfiguration` interface is not available by default.  To get around this, I used the `Microsoft.Extensions.Configuration` package (and couple others) to read the `appsettings.json` file directly.  This strongly typed settings class is a singleton and is accessed throughout the add-in via `AddIn.Settings`.

To monitor for changes (since `IOptionsSnapshot<T>` pattern is not available), I used a `FileSystemWatcher` to monitor the `appsettings.json` file for changes.  When a change is detected, the settings are reloaded (with a little protection against multiple notifications).  

Below I will demonstrate what is needed to wire this all together.

1. The *.csproj file needs to include the following package references:

```xml
<PackageReference Include="Microsoft.Extensions.Configuration" Version="7.0.0" />
<PackageReference Include="Microsoft.Extensions.Configuration.Binder" Version="7.0.0" />
<PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="7.0.0" />
```

2.  For this documentation, assume the AddInSettings class simply has a single property.

```csharp
public class AddInSettings
{
	public bool ShowRibbon { get; init; }
}
```

3. In `IExcelAddIn.AutoOpen`, leverage the `FileWatcherNotification` class to monitor the `appsettings.json` file for changes and when a change is detected, reload the settings and invalidate the ribbon (the first time through, the ribbon might not be ready, but when subsequent 'file/settings' updates occur, it will be ready).

```csharp
public class AddIn : IExcelAddIn
{
	internal static AddInSettings Settings = new();
	private FileWatcherNotification settingsProcessor = null!;

	public void AutoOpen()
	{
		settingsProcessor = new( 
			notificationDelay: 300, 
			path: Path.GetDirectoryName( (string)XlCall.Excel( XlCall.xlGetName ) )!, 
			filter: "appsettings.json", 
			action: e => {
			try
			{
				IConfiguration configuration = new ConfigurationBuilder()
					.AddJsonFile( e.FullPath, optional: true )
					.Build();

				Settings = configuration.GetSection( "AddInSettings" ).Get<AddInSettings>() ?? new();
			}
			catch ( Exception ex )
			{
				// TODO: Need to log this somewhere...event viewer via Logging?
				Console.WriteLine( ex.ToString() );
				Settings = new();
			}

			// On the manual .Change() method called after this constructor, the Ribbon might not yet be initialized
			Ribbon.CurrentRibbon?.InvalidateSettings();
		} );

		settingsProcessor.Changed();
	}
}

/// <summary>
/// FileSystemWatcher notifications are capable of happening 'multiple' times for a single 'action'.  For example, if Notepad saves a file,
/// you might not get a single 'Changed' event when everything is 'done', you might get multiple 'Changed' events.  Similarily, "when a file is 
/// moved from one directory to another, several OnChanged and some OnCreated and OnDeleted events might be raised." (from MS Docs).  This class 
/// mitigates that by having an internal timer that starts/restarts on each event.  So once the timer is created (first event), if no other events
/// occur for notificationDelay milliseconds, then, and only then, is the event raised.
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
```

#### Access Settings

To access the settings, simply use `AddIn.Settings.*` properties when needed.  However, I had one property (the only one in this sample) that needed to update the ribbon immediately when the settings where changed.  The call in the previous sample code to `Ribbon.CurrentRibbon?.InvalidateSettings();` is what accomplishes this.

** Ribbon.xml **
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
	<ribbon startFromScratch="false">
		<tabs>
			<tab id="btrRBLe" keytip="K" label="KAT Tools" getVisible="Ribbon_GetVisible">
				
				<!-- All the group elements making up my ribbon omitted for brevity -->

			</tab>
		</tabs>
	</ribbon>
</customUI>
```

** Ribbon.cs **
```csharp
public partial class Ribbon : ExcelRibbon
{
	public static Ribbon CurrentRibbon { get; private set; } = null!;
	private bool showRibbon;

	public void Ribbon_OnLoad( IRibbonUI ribbon )
	{
		this.ribbon = ribbon;
		showRibbon = AddIn.Settings.ShowRibbon;
	}

	public void InvalidateSettings()
	{
		// Store new setting and invalidate the ribbon
		showRibbon = AddIn.Settings.ShowRibbon;
		ribbon.InvalidateControl( "btrRBLe" );
	}

	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"btrRBLe" => showRibbon,
			_ => true,
		};
	}
}
```