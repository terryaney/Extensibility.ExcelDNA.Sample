# Excel-DNA Add-In

An Excel-DNA add-in for the KAT teams that uses many of the features provided by Excel-DNA along with how I overcame many, not so obvious, obstacles.

## TODO

1. Ribbon_GetContent - is hacked to get it working
1. Propmt for DebugCalcEngines credentials and store in preferences.json - `var newHash = ( new PasswordHasher<string>().HashPassword( "terry.aney", "password" ) ).Dump();`
```
// Encrypt (hash MacAddress(length64))(encrypted(password))
// Decrypt, verify Mac with hash
// Decrypt, decrypted password and pass it along

				// Cryptography3DES.DefaultEncryptAsync
```
1. Save History - Load/Save window position in preferences
1. Implement Ribbon handlers
	1. exportMappedxDSData - Rename this better after you figure out what it is doing
1. [Custom Intellisense/Help Info](https://github.com/Excel-DNA/IntelliSense/issues/21) - read this and linked topic to see what's possible
	1. https://github.com/Excel-DNA/Tutorials/blob/master/SpecialTopics/IntelliSenseForVBAFunctions/README.md
1. [Possible Async Information](https://github.com/Excel-DNA/Samples/blob/master/Registration.Sample/AsyncFunctionExamples.cs)
	1. https://excel-dna.net/docs/guides-advanced/performing-asynchronous-work
	1. https://excel-dna.net/docs/tips-n-tricks/creating-a-threaded-modal-dialog
	1. Excel-DNA/Samples/Archive/Async/NativeAsyncCancellation.dna
1. BTR* functions...
	1. https://excel-dna.net/docs/guides-advanced/dynamic-delegate-registration - possible dynamic creation instead of having to create functions for each item and passing through?  Reference the SSG assembly detect custom functions.
1. Look for email/thread about excel not shutting down properly
1. Readme
	1. Badge count on ribbon image
	1. Dynamic menus

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

1. [Ribbon Organization](#ribbon-organization)
1. [async/await Issues](#asyncawait-issues)
1. [ExcelIntegration.RegisterUnhandledExceptionHandler](#excelintegrationregisterunhandledexceptionhandler)
1. [appsettings.json Support](#appsettingsjson-support)
1. [Fixing Workbook Links](#fixing-workbook-links)
1. [Changing Visible/Enabled State of Ribbon Controls](#changing-visibleenabled-state-of-ribbon-controls)
1. [Custom Ribbon Image with Badge Count](#custom-ribbon-image-with-badge-count)
1. [Using Windows Form Dialogs](#using-windows-form-dialogs)

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

Back to [Features listing](#features).

### async/await Issues

The KAT add-in uses async/await code to perform various tasks.

1. https://groups.google.com/g/exceldna/c/_pKphutWbvo - question asking about my different scenarios
2. https://stackoverflow.com/a/68303070/166231 - Cleary answer about async

Issues:

1. Availability of await/async - some places it was impossible, i.e. `Application_WorkbookBeforeSave` because of `ref bool Cancel`.
2. Desire to have code await a method and not continue running until complete, i.e. `Application_WorkbookBeforeSave` where I needed the `ProcessSaveHistoryAsync` method to complete before flow returned and `Application_WorkbookAfterSave` was called.
3. Ability to run async code without negatively affecting Excel's calculation thread - when problems happened and Excel attempts to shutdown, it does not terminate the msexcel.exe process, but when the current Excel window is closed it immediately launches a new window.

Async Tasks in Excel-DNA:

1. Calling `Camelot.Api.Excel` web application methods - during handlers of RibbonUI buttons, during Excel 'state' change to query information about the current CalcEngine, and during RibbonUI handlers (i.e. GetContent) to query information on demand about current CalcEngine.
2. Ability to launch and run a long running task in a separate thread that can be cancelled if needed and then report back information to the main thread (i.e. Local Batch Processes)

Code Locations and Requirements

| Method | Requirement | Strategy |
| --- | --- | --- |
| `Ribbon_GetContent` | `Camelot.Api.Excel` | `.GetAwaiter().GetResult()` |


Back to [Features listing](#features).


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

	application.WorkbookActivate += Application_WorkbookActivate;
	application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
}

public override void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
{
	base.OnDisconnection( RemoveMode, ref custom );

	application.WorkbookActivate -= Application_WorkbookActivate;
	application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
}

private void Application_WorkbookActivate( MSExcel.Workbook Wb )
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
		ribbon.InvalidateControl( "katShowDiagnosticLog" );
	}		
}

public void RBLe_ShowLog( IRibbonControl? _ )
{
	ExcelDna.Logging.LogDisplay.Show();
	auditShowLogBadgeCount = 0;
	ribbon.InvalidateControl( "katShowDiagnosticLog" );
}

private int auditShowLogBadgeCount;
public Bitmap Ribbon_GetImage( IRibbonControl control )
{
	switch ( control.Id )
	{
		case "katShowDiagnosticLog":
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

Back to [Features listing](#features).

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

Note: See [ExcelRna.Extensions.Hosting](https://github.com/altso/ExcelRna.Extensions.Hosting) for what seems like a possible solution for Dependency Injection in Excel-DNA.  The project looks very promising, but I wanted to try and only use Excel-DNA for this project until Dependency Injection was a requirement.

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

				Settings = configuration.GetSection( "addInSettings" ).Get<AddInSettings>() ?? new();
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
			<tab id="tabKat" keytip="K" label="KAT Tools" getVisible="Ribbon_GetVisible">
				
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
		ribbon.InvalidateControl( "tabKat" );
	}

	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"tabKat" => showRibbon,
			_ => true,
		};
	}
}
```

Back to [Features listing](#features).

### Fixing Workbook Links

Totally random topic here, but thought I'd document it.  During the original creation of addin (.NET Framework and/or *.xla files) we often had problems where links to add-ins would get broken because users who uploaded Spreadengines to be used by our APIs had different installation locations for the various addins that were required.

For example, let's say we have an addin named `rbl.xla` that exposed a function called `CalculateProjection`.  If the user had a formula of `=CalculateProjection(A1)` in their workbook all worked fine.  But when they uploaded, and the rbl.xla file was not in the same location as the workbook, the link would be broken **and the formula would be modified** to `=c:\installation\path\to\rbl.xla!CalculateProjection(A1)`.  This was compounded when the formula used several functions from `rbl.xla` because the length of the formula (after the path injections) would sometimes exceed the allowed limit for a formula expression and simply lose chunks of the formula.

To combat this, our 'calculation servers' would run a function like the `UpdateWorkbookLinks` below when a workbook was opened before it was prcoessed.  This process has continued to live on in the Excel-DNA add-in.  To be honest, I'm not 100% sure it is still needed (i.e. if Excel improved its addin installation location detection).  But we do run the `UpdateWorkbookLinks` during the `Application_WorkbookOpen` event and also expose a ribbon button to run it manually in case users are sharing Excel files and having link issues.

```csharp
private void UpdateWorkbookLinks( MSExcel.Workbook wb )
{
	if ( wb == null )
	{
		ExcelDna.Logging.LogDisplay.RecordLine( $"LinkToLoadedAddIns: ActiveWorkbook is null." );
		return;
	}

	if ( !WorkbookState.HasLinks ) return;

	var linkSources = ( wb.LinkSources( MSExcel.XlLink.xlExcelLinks ) as Array )!;

	var protectedInfo = wb.ProtectStructure
		? new[] { "Entire Workbook" }
		: wb.Worksheets.Cast<MSExcel.Worksheet>().Where( w => w.ProtectContents ).Select( w => string.Format( "Worksheet: {0}", w.Name ) ).ToArray();

	if ( protectedInfo.Length > 0 )
	{
		MessageBox.Show( 
			"Unable to update links due to protection.  The following items are protected:\r\n\r\n" + string.Join( "\r\n", protectedInfo ), 
			"Unable to Update", 
			MessageBoxButtons.OK, 
			MessageBoxIcon.Warning 
		);
		return;
	}

	foreach ( var addin in application.AddIns.Cast<MSExcel.AddIn>().Where( a => a.Installed ) )
	{
		var fullName = addin.FullName;
		var name = Path.GetFileName( fullName );

		foreach ( object o in linkSources )
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
}
```

Back to [Features listing](#features).

### Changing Visible/Enabled State of Ribbon Controls

Given the size of our ribbon, the visiblity and enabled states were toggling based on the current context of the workbook and or worksheet.  The following shows different parts of our addin demonstrating how we implemented this.

1. Changing `Ribbon.xml` to have `getEnabled` and `getVisible` attributes indicating which method to call to determine the state of the control.
1. Implementing the `Ribbon_OnLoad`, `Ribbon_GetVisible` and `Ribbon_GetEnabled` methods to use a built in `WorkbookState` class to determine the proper values for the requested state.
1. Implementing the `WorkbookState` class to determine the state of the workbook and worksheet.  Can example repository code, but this is just any mechanism you choose to determine the state of the workbook and worksheet (i.e. presence of named ranges, presence of tabs, etc.).
1. Refreshing the ribbon when `WorkbookState` was updated via the `InvalidateControl` method.

#### Modifying Ribbon.xml for Visibility and Enabled State

Part of our ribbon.xml file showing the `onLoad` specified on `customUI` element and then the `getEnabled` and `getVisible` attributes specified on a `button` element (but can be applied to `tab`, `group`, `button`, etc.).

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
	<ribbon startFromScratch="false">
		<tabs>
			<tab id="tabKat" keytip="K" label="KAT Tools">
				<group id="groupConfigurationExporting" keytip="ss" label="Configuration Exporting" imageMso="WorkspaceHtmlProperties">
					<button id="configurationExportingGlobalTables" keytip="G" label="Process Global Tables" imageMso="ExportMoreMenu" size="normal" onAction="Ribbon_OnAction" tag="ConfigurationExporting_ProcessGlobalTables" getVisible="Ribbon_GetVisible" getEnabled="Ribbon_GetEnabled"/>
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>
```

#### Handling Ribbon Events for Visibility and Enabled State

To manually refresh the state of the ribbon, we need to call the `InvalidateControl` method on the `IRibbonUI` object.  Therefore, in the `Ribbon_OnLoad` method, we store the `IRibbonUI` object in a class level variable.  Simple examples of the `Ribbon_GetVisible` and `Ribbon_GetEnabled` methods are also listed.

```csharp
public partial class Ribbon
{
	private IRibbonUI ribbon = null!;

	public void Ribbon_OnLoad( IRibbonUI ribbon ) => this.ribbon = ribbon;

	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"id1" => someCondition,

			"id2" => someCondition2,

			"id3" or "id4" => someCondition3,

			_ => true,
		};
	}

	public bool Ribbon_GetEnabled( IRibbonControl control )
	{
		return control.Id switch
		{
			"id1" => someCondition,

			"id2" => someCondition2,

			"id3" or "id4" => someCondition3,

			_ => true,
		};
	}
}
```

#### Refreshing Ribbon State on Demand

Nothing too complicated about manually refreshing the ribbon state.  Basically, you just have to indicate which controls to refresh (via `InvalidateControl` method on the `IRibbonUI` object) when the 'context' changes.  To accomplish this you probably will at a minimum have the following application event handlers:

```csharp
public partial class Ribbon : ExcelRibbon
{
	private readonly MSExcel.Application application;

	// This is a class that has a static `Current` method that detects all the current context indicators
	// that our ribbon needs to know about.  To force a refresh, you can set workbookState to null so it is regenerated.
	private WorkbookState? workbookState = null;
	public WorkbookState WorkbookState => workbookState ??= WorkbookState.Current( application ) ;

	public Ribbon()
	{
		application = ( ExcelDnaUtil.Application as MSExcel.Application )!;
	}

	public override void OnConnection( object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom )
	{
		base.OnConnection( Application, ConnectMode, AddInInst, ref custom );
		application.WorkbookActivate += Application_WorkbookActivate;
		application.SheetActivate += Application_SheetActivate;
	}

	public override void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
	{
		base.OnDisconnection( RemoveMode, ref custom );
		application.WorkbookActivate -= Application_WorkbookActivate;
		application.SheetActivate -= Application_SheetActivate;
	}

	private void Application_WorkbookActivate( MSExcel.Workbook wb )
	{
		workbookState = null;
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnWorkbookChange );
	}

	private void Application_SheetActivate( object sheet )
	{
		workbookState = null;
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnSheetChange );
	}

	// Sample list of button ids that should be refresheds 'OnCalcEngineManagement'.
	// Above, RibbonStatesToInvalidateOnWorkbookChange and RibbonStatesToInvalidateOnSheetChange
	// are just arrays for more IDs to manage visible/enabled state when the context changes.
	readonly string[] RibbonStatesToInvalidateOnCalcEngineManagement =
		new[] { "katDataStoreDownloadLatest", "katDataStoreCheckIn", "katDataStoreCheckOut" };

	public void KatDataStore_CheckInCalcEngine( IRibbonControl control )
	{
		// Omitted code that performs a 'Check In' of the current Spreadsheet, and when 
		// successful, manually trigger a ribbon invalidate on the controls that care about
		// 'checked in' state.

		workbookState = null;
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
	}

	// Sample button click handler that manually refreshes the *ENTIRE* ribbon.
	// This was useful when the ribbon was in a 'bad' state or the 'context' could have been
	// updated without Excel being aware and needed to be refreshed.
	public void RBLe_RefreshRibbon( IRibbonControl _ )
	{
		workbookState = null;
		ribbon.Invalidate();
	}
}

// Simple extension method to wrap Invalidating more than one control.
public static class ExcelExtensions
{
	public static void InvalidateControls( this IRibbonUI ribbon, params string[] controlIds )
	{
		foreach ( var controlId in controlIds )
		{
			ribbon.InvalidateControl( controlId );
		}
	}
}
```

Back to [Features listing](#features).

### Custom Ribbon Image with Badge Count

To help promote cleaner Spreadsheet development, we wanted to show a badge count on a ribbon image when there were errors in the workbook formulas as an indicator to the developers.  See [ExcelIntegration.RegisterUnhandledExceptionHandler](#excelintegrationregisterunhandledexceptionhandler) section for the sample code that demonstrates how to accomplish this via the ribbons `getContent` method.

Back to [Features listing](#features).

### Using Windows Form Dialogs

Our addin makes use of Windows Form dialogs when information from the user is required to perform a task.  I've never looked into implementing custom Task Panes, but in terms of the 'problems' I face in VS Code, I don't think it matters.  At the time of writing, VS Code did not have a visual designer for Windows Forms.  There are two solutions to this.  First, you can manually code up all the controls by hand inside a 'designer' partial class, or you can use Visual Studio to create the form UI and then bring the files over to VS Code.  I chose the latter.

After the form is created in Visual Studio, small tweaks to the UI can probably be accomplished fairly easily in VS Code.  But if major UI changes are required, I'll probably end up back in Visual Studio.

You can look at the `SaveHistory.cs` file under the `Views` folder and the `ProcessSaveHistory` method to examine the code, but in general I use the following pattern.

```csharp
// Method that needs input from the user...
private CalcEngineUploadInfo? ProcessSaveHistory( MSExcel.Workbook workbook )
{
	// Constructor to my forms take in 'models' they need to properly initialize themselves
	// if there are differing 'modes' of the form.
	using var saveHistory = new SaveHistory( workbook, WorkbookState );

	// Then I expose a public method that will return an object that contains the 'information'
	// gathered from the user/form along with an additional `DialogResult Result` property that
	// indicates the users resopnse/intent to the dialog.
	var saveHistoryInfo = saveHistory.GetHistoryInformation();

	// Omitted code - perform the appropriate code (or not) based on the `Result` property.
	if ( saveHistoryInfo.Result == DialogResult.Cancel )
	{
		return null;
	}

	// Peform action with information provided in saveHistoryInformation...
}
```

Back to [Features listing](#features).
