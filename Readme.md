# Excel-DNA Add-In

An Excel-DNA add-in for the KAT teams that uses many of the features provided by Excel-DNA along with how I overcame many, not so obvious, obstacles.

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
            "args": ["/x", "${workspaceFolder}\\src\\bin\\Debug\\net7.0-windows\\KAT.Extensibility.Excel.xll"],
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
			"systemId": "https://raw.githubusercontent.com/Excel-DNA/ExcelDna/master/Distribution/XmlSchemas/customui14.xsd"
		}
	]
}
```

## Features

1. [Ribbon Organization](#ribbon-organization)
1. [Thread Context Issues](#thread-context-issues)
1. [ExcelIntegration.RegisterUnhandledExceptionHandler](#excelintegrationregisterunhandledexceptionhandler)
1. [Dependency Injection](#dependency-injection)
1. [Changing Visible/Enabled State of Ribbon Controls](#changing-visibleenabled-state-of-ribbon-controls)
1. [Using Windows Form Dialogs](#using-windows-form-dialogs)
1. [Fixing Workbook Links](#fixing-workbook-links)

### Ribbon Organization

The use of the `IRibbonUI` functionality in the KAT Tools add-in is quite extensive.  There is state management (enabled/visible) of CustomUI elements via ribbon events, CustomUI element handlers, dynamic menus with content from async method calls, and dynamic images to name a few.  In this section I will describe some of the challenges I faced with `IRibbonUI` and how I overcame them.

**Helpful Documenation Links**

1. [CustomUI Reference](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43)
1. [imageMso Reference](https://codekabinett.com/download/Microsoft-Office-2016_365-imageMso-Gallery.pdf)

**Managing the Sheer Amount of Code**

Given the amount of code I had to implement to provide all the required functionality, the number of lines/methods became quite overwhelming (at least given the way I organized the code - leaving vast majority of the code within the `Ribbon` class).  To help alleviate this, I used [partial classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods#partial-classes) as an organizational tool.  This made it easier to find and maintain the code for me but your mileage may vary.  Additionally, to make this separation easier to manage in the (Solution) Explorer side bar, I would suggest enabling file nesting.  To enable file nesting in Visual Studio Code, add the following to your `settings.json` file:

```json
{
	"explorer.fileNesting.patterns": {
		"*.cs": "${capture}.*.cs"
	}
}
```

I also used method prefixes that matched the CustomUI `group.id` as well to make code navigation easier (via `CTRL+T` keyboard shortcut).  For example, for my group with an `id` of `groupNavigation`, the methods all have the prefix `Navigation_`.

Back to [Features listing](#features).

### Thread Context Issues

The following requirements were some of the challenges I faced regarding Excel thread context switching.  In some cases, the code would not function correctly (or at all), but in many cases, the code would execute as I expected, but when I closed Excel, it attempts to shutdown but the `msexcel.exe` process is not terminated and after the current Excel window is closed it immediately launches a new window with no spreadsheet and the add-in is not displayed.  Attempting to view the add-in in Excel's add-in dialog caused Excel to GPF and shutdown, requiring the user to re-add the add-in via Excel's dialog.


1. **Api Functionality via HttpClient** - The `Camelot.Api.Excel` api project is used both to manage state of the ribbon as well as provide functionality for some of the button events.  See [`WorkbookState.RefreshAsync`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Models/WorkbookState.cs) or [`Ribbon.KatDataStore_DownloadLatestCalcEngine`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Handlers.KatDataStore.cs) for some examples.
2. **Long Running Tasks** - Some ribbon handlers require launching a long running task that can be cancelled if needed, does not block the main Excel threads, and then reports back information to the main thread (i.e. Local Batch Processes).
3. **async/await Support in Handlers** - The ability to run async code in Ribbon button handlers and follow up interactions with Excel COM objects to manipulate the Excel application in some fashion (i.e. `ribbon.Invalidate()` or other `application.*` methods (`application.Workbooks.Open`)).
4. **FileSystemWatcher Event Support** - A `FileSystemWatcher` that is monitoring `appsettings*.json` file(s) for changes triggers change events when necessary that need to reload settings and then invalid the ribbon.  See [AddIn.AutoOpen](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/AddIn.cs) for more detail.
5. **Global Exception Handler** - A global exception handler is added via `ExcelIntegration.RegisterUnhandledExceptionHandler` to provide the ability to catch all formula/calculation exceptions.  In the registered [UnhandledExceptionHandler](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/AddIn.cs) handler, the error is logged and a ribbon badge is updated with current count of errors.  See the [ExcelIntegration.RegisterUnhandledExceptionHandler](#excelintegrationregisterunhandledexceptionhandler) section for more detail.

**Reference Links for Thread Context and async/await Issues**

1. https://groups.google.com/g/exceldna/c/_pKphutWbvo - question asking about my different scenarios
1. https://groups.google.com/g/exceldna/c/ILgL-dW47A4/m/9HrOyClJAQAJ - Thread about ensuring Excel shuts down properly.
1. https://stackoverflow.com/a/68303070/166231 - Stephen Cleary's answer about async/await best practices.
1. https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development - Async Programming article by Stephen Cleary.

The following are some of the specific issues I had to overcome/workaround.  To be honest, I do not have my head wrapped around Thread context switching/blocking.  So if anyone reads this and can explain it better, please contact me.

1. **Availability of await/async** - In some places (i.e. [`Application_WorkbookBeforeSave`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Events.Excel.cs)), it was impossible to use `await`/`async`.  To work around this, I used [The Thread Pool Hack](https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development#the-thread-pool-hack) to run the async method from within synchronous code.  
	1. [`Ribbon_GetContent`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Events.cs)
		1. Originally, I simply tried changing the handler signature to async (`public async Task<string?> Ribbon_GetContent( IRibbonControl control )`) and use `await` as normal hoping Excel Integration would properly call my method.  The code execution *did* wait until the api call was complete before returning the `menu.ToString(), however, before it completed, Excel had immediately shown a single, empty menu item and never refreshed it.
		1. I then tried using `ExcelAsyncUtil.QueueAsMacro`, of course this didn't work because I couldn't `await` this method call, so the handler returned back to Excel immediately before the `menu.ToString()` was complete.
		1. I then tried using `ExcelAsyncTaskScheduler` after reading [this discussion](https://groups.google.com/g/exceldna/c/9OkHWILuFMo/m/RpilXElgAQAJ), but it had similar results to `ExcelAsyncUtil.QueueAsMacro`.
		1. I then tried [The Blocking Hack](https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development#the-blocking-hack) but if `ExcelDna.Logging.LogDisplay.WriteLine` or `Show` were called, the subsequent call to 'The Blocking Hack' would hang indefinitely.
		1. Finally, I used 'The Thread Pool Hack' to get the desired behavior (menu populating correctly, and Excel closing gracefully).
	1. [`Application_WorkbookBeforeSave`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Events.Excel.cs) - Most of the `public void Application_*` event handlers seemed to function properly simply by placing an `async` keyboard in the method signature.  However, the `Application_WorkbookBeforeSave` handler was a special case because of the `ref bool Cancel` parameter preventing the use of the keyword.
		1. Originally, I was using the 'The Blocking Hack' and it seemed to run correctly.
		1. After the 'The Blocking Hack' was replaced in `Ribbon_GetContent`, I decided to use the 'The Thread Pool Hack' in `Application_WorkbookBeforeSave` for consistency and it worked as expected.
1. **Needing to await asynchronous code** - There were cases where I desired code flow to behave like normal `async/await` code (i.e. code flow pauses until async code is complete).
	1. `Application_WorkbookBeforeSave` - I needed the `ProcessSaveHistoryAsync` method to complete and assign the `calcEngineUploadInfo` variable before flow returned, allowing `Application_WorkbookBeforeSave` to complete and ultimately call `Application_WorkbookAfterSave` which performs asynchronous code if `calcEngineUploadInfo` is assigned.
	1. `Ribbon_GetContent` - This same requirement was needed in this method as well, waiting for `menu` to be created before returning the result to Excel.
1. **Access `ExcelDna.Logging.LogDisplay`/`ribbon.Invalidate()` While On Calculation Thread** - In the `UnhandledExceptionHandler` global exception handler, the use of both of these methods must be from within an `ExcelAsyncUtil.QueueAsMacro` wrapper otherwise Excel would not close gracefully.

Below is a summary of the code locations that faced thread context issues and what was used to overcome them.

| Method | Requirement | Strategy |
| --- | --- | --- |
| `AddIn.AutoOpen` | `FileSystemWatcher` | `ExcelAsyncUtil.QueueAsMacro` |
| `UnhandledExceptionHandler` | `Excel COM Access` | `ExcelAsyncUtil.QueueAsMacro` |
| `Ribbon_GetContent` | `Camelot.Api.Excel` | `Task.Run( () => *Async() ).GetAwaiter().GetResult()` |
| `Application_WorkbookBeforeSave` | `Camelot.Api.Excel` | `Task.Run( () => *Async() ).GetAwaiter().GetResult()` |
| `Ribbon_OnAction` | `Camelot.Api.Excel`<br/>`Async Functionality`<br/>`Long Running Task` | `ExcelAsyncUtil.QueueAsMacro` |
| `KatDataStore_DownloadLatestCalcEngine` | `Camelot.Api.Excel`<br/>`Excel COM Access` | `ExcelAsyncUtil.QueueAsMacro` |
| `KatDataStore_DownloadDebugFile` | `Camelot.Api.Excel`<br/>`Excel COM Access` | `ExcelAsyncUtil.QueueAsMacro` |

Back to [Features listing](#features).

### ExcelIntegration.RegisterUnhandledExceptionHandler

A global exception handler is registered to log diagnostic information to the `ExcelDna.Logging.LogDisplay` window and update a ribbon image with a badge count.  In the diagnostic information, I wanted to display the address and formula of the offending cell.  Since the error handler runs on Excel's calculation thread, but directly converting the `ExcelReference` to an address can't (easily) be done in this context, so `ExcelAsyncUtil.QueueAsMacro` is required to register a delegate to run in a safe context on the main thread, from any other thread or context (i.e. when the calculation completes).  This [conversation](https://groups.google.com/d/msg/exceldna/cHD8Tx56Msg/MdPa2PR13hkJ) explains why `QueueAsMacro` is required for other `XlCall` methods.

To help promote cleaner Spreadsheet development, we wanted to show a badge count on a ribbon image when there were errors in the workbook formulas as an indicator to the developers.  The following code demonstrates a possible solution.

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

	application.WorkbookDeactivate += Application_WorkbookDeactivate;
	application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
}

private void Application_WorkbookDeactivate( MSExcel.Workbook Wb )
{
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

// Only report error if not already reported once for current formula/worksheet
private readonly ConcurrentDictionary<string, string?> cellsInError = new();

public void LogFunctionError( ExcelReference caller, object exception )
{
	var address = caller.GetAddress();
	var formula = caller.GetFormula();

	var reportError = !cellsInError.TryGetValue( address, out var failedFormula ) || failedFormula != formula;
	cellsInError[ address ] = formula;

	if ( reportError )
	{
		var message = $"Error: {address} {formula ?? "unavailable"}{Environment.NewLine}{exception}";
	
		ExcelDna.Logging.LogDisplay.RecordLine( message );
		auditShowLogBadgeCount++;

		ribbon.InvalidateControl( "katShowDiagnosticLog" );
	}		
}

// Ribbon button to show log manually
public void RBLe_ShowLog( IRibbonControl? _ )
{
	ExcelDna.Logging.LogDisplay.Show();
	auditShowLogBadgeCount = 0;
	ribbon.InvalidateControl( "katShowDiagnosticLog" );
}

// Ribbon event handler to draw badge count on ribbon image if needed
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

	public static string GetAddress( this ExcelReference? reference ) => (string)XlCall.Excel( XlCall.xlfReftext, reference, true /* true - A1, false - R1C1 */ );
}
```

Back to [Features listing](#features).

### Dependency Injection

As Excel-DNA documentation has stated, [it does not want to include Dependency Injection into the project](https://github.com/Excel-DNA/ExcelDna/issues/20#issuecomment-135950407).  This means that classes/interfaces like `IConfiguration` and `IHttpClientFactory` are not available by default in the normal usage pattern.

As suggested in that discussion, I used static and/or global classes to provide the functionality I needed.

Note: See [ExcelRna.Extensions.Hosting](https://github.com/altso/ExcelRna.Extensions.Hosting) for what seems like a possible solution for Dependency Injection in Excel-DNA.  The project looks very promising, but I wanted to try and only use Excel-DNA for this project until Dependency Injection was a requirement.

#### IHttpClientFactory Support

As documented in the [Thread Context Issues](#thread-context-issues) section, I needed to use `HttpClient` to make api calls to the Camelot.Api.Excel project.  As discussed in the [Use IHttpClientFactory to implement resilient HTTP requests](https://learn.microsoft.com/en-us/dotnet/architecture/microservices/implement-resilient-applications/use-httpclientfactory-to-implement-resilient-http-requests) article by Microsoft, there can be some problems with using the `HttpClient` class directly.  I'm not sure if creating a single static `HttpClient` would suffice for the lifetime of the add-in, but I decided to use the `IHttpClientFactory` to manage the `HttpClient` instances which allowed the class to follow the same pattern used in our web sites and apis when working with `HttpClient` class.

To accomplish that, I use the have a single instance of the classes that need an `IHttpClientFactory` stored in the `Ribbon` class which are set up during the `Ribbon` constructor.  To enable this support, the following code creates a `ServiceCollection` and build a `ServiceProvider` to enable the supoprt of `IHttpClientFactory` and then creates the `HttpClient` instances as needed.

```csharp
// Create service collection
var services = new ServiceCollection();
services.AddHttpClient();
var serviceProvider = services.BuildServiceProvider();
var clientFactory = serviceProvider.GetService<IHttpClientFactory>()!;
apiService = new ApiService( clientFactory );
```

#### IConfiguration / appsettings.json Support

The KAT add-in requires support for user settings and secrets and the most convenient way to provide that functionality was simply by leveraging an `Microsoft.Extensions.Configuration.IConfiguration` capabilities.

To enable this support:

1. Output the `appsettings.json` file to the output directory during build so that it could be used during debugging.  
1. For settings that should not be shared between users, they will be stored in an `appsettings.secrets.json` file that is not distributed/shared amoung users nor present in source control.
1. Read and monitor the `appsettings.json` and `appsettings.secrets.json` files for changes and updates object/UI on demand when files change instead of requiring a restart.
1. Access settings throughout the code base via the `AddIn.Settings` static object.

##### Output `appsettings.json` File

Simply add the following to the `.csproj` file and the `appsettings.json` file will be copied to the output directory during build.

```xml
<ItemGroup>
	<Content Include="appsettings.json">
		<CopyToOutputDirectory>Always</CopyToOutputDirectory>
	</Content>
</ItemGroup>
```

##### Read and Monitor `appsettings.json` File

This was probably the trickiest part of the process.  I used the `Microsoft.Extensions.Configuration` package (and couple others) to read the `appsettings.json` file directly and bind it to a strongly typed settings object.  This strongly typed settings object is a singleton and is accessed throughout the add-in via `AddIn.Settings`.

To monitor for changes (since `IOptionsSnapshot<T>` pattern is not available), I used a `FileSystemWatcher` to monitor the `appsettings.json` file for changes.  When a change is detected, the settings are reloaded (with a little protection against multiple notifications inside [FileWatcherNotification](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Configuration/FileWatcherNotification.cs)).  

Below I will demonstrate what is needed to wire this all together.

1. The *.csproj file needs to include the following package references: `Microsoft.Extensions.Configuration`, `Microsoft.Extensions.Configuration.Binder`, and `Microsoft.Extensions.Configuration.Json`.

2.  For this documentation, assume the AddInSettings class has a single property (see [AddInSettings](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Configuration/AddInSettings.cs) for all the settings that are supported).

```csharp
public class AddInSettings
{
	public bool ShowRibbon { get; init; }
}
```

3. In `IExcelAddIn.AutoOpen`, leverage the `FileWatcherNotification` class to monitor the `appsettings.json` files for changes and when a change is detected, reload the settings and invalidate the ribbon (the first time through, the ribbon might not be ready, but when subsequent 'file/settings' updates occur, it will be ready).

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
			filter: "appsettings*.json", 
			action: e => {
				try
				{
					IConfiguration configuration = new ConfigurationBuilder()
						.AddJsonFile( Path.Combine( XllPath, "appsettings.json" ), optional: true )
						.AddJsonFile( Path.Combine( XllPath, "appsettings.secrets.json" ), optional: true )
						.Build();

					Settings = configuration.GetSection( "addInSettings" ).Get<AddInSettings>() ?? new();
				}
				catch ( Exception ex )
				{
					Ribbon.LogError( "Unable to initialize IConfiguraiton for appsettings.json.  Using default settings.", ex );
					Settings = new();
				}

				ExcelAsyncUtil.QueueAsMacro( () => ribbon.InvalidateControl( "tabKat" ) );
			} );

		settingsProcessor.Changed();
	}
}
```

##### Access Settings

To access the settings, simply use `AddIn.Settings.*` properties when needed.  However, I had one property (the only one in this sample) that needed to update the ribbon immediately when the settings where changed.  The call in the previous sample code to `Ribbon.CurrentRibbon?.InvalidateFeatures();` is what accomplishes this.

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
	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"tabKat" => AddIn.Settings.ShowRibbon,
			_ => true,
		};
	}
}
```

Back to [Features listing](#features).

### Changing Visible/Enabled State of Ribbon Controls

Given the size of our ribbon, the visiblity and enabled states were toggling based on the current context of the workbook and or worksheet.  The following shows different parts of our addin demonstrating how we implemented this.

1. Changing `Ribbon.xml` to have `getEnabled` and `getVisible` attributes indicating which method to call to determine the state of the control.
1. Implementing the [`WorkbookState`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Models/WorkbookState.cs) class to determine the state of the active workbook and worksheet.
1. Implementing the `Ribbon_OnLoad`, `Ribbon_GetVisible` and `Ribbon_GetEnabled` methods to use a built in `WorkbookState` class to determine the proper values for the requested state.
1. Refreshing the ribbon when `WorkbookState` was updated via the `ribbon.Invalidate()` method.

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

To manually refresh the state of the ribbon, we need to call the `IRibbonUI.Invalidate` method.  Therefore, in the `Ribbon_OnLoad` method, we store the `IRibbonUI` object in a class level variable.  Simple examples of the `Ribbon_GetVisible` and `Ribbon_GetEnabled` methods are also listed.

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

Nothing too complicated about manually refreshing the ribbon state.  To accomplish this you probably will at a minimum have the following application event handlers:

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
		ribbon.Invalidate();
	}

	private void Application_SheetActivate( object sheet )
	{
		workbookState = null;
		ribbon.Invalidate();
	}

	public void KatDataStore_CheckInCalcEngine( IRibbonControl control )
	{
		// Code omitted for brevity
		workbookState = null;
		ribbon.Invalidate();
	}

	public void RBLe_RefreshRibbon( IRibbonControl _ )
	{
		// Code omitted for brevity
		workbookState = null;
		ribbon.Invalidate();
	}
}
```

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

### Fixing Workbook Links

During the original creation of addin (.NET Framework and/or *.xla files) we often had problems where links to add-ins would get broken because users who uploaded Spreadengines to be used by KAT services had different installation locations for the various addins that were required.

For example, assume there is an addin named `rbl.xla` that exposed a function called `CalculateProjection`.  If the user had a formula of `=CalculateProjection(A1)` in their workbook all worked fine.  But when they uploaded file to be used in KAT services, and the rbl.xla file was not in the same location on the server as it was for user, the link would be broken **and the formula would be modified** to `=c:\user\installation\path\to\rbl.xla!CalculateProjection(A1)`.  If not corrected upon opening, this was compounded when the formula used several functions from `rbl.xla` because the length of the formula (after the path injections) would sometimes exceed the allowed limit for a formula expression and simply lose chunks of the formula.

To combat this, our 'calculation servers' would run a function like the [UpdateWorkbookLinks](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Handlers.CalcEngineUtilities.cs) when a workbook was opened, *before it was prcoessed*.  This process has continued to live on in the Excel-DNA add-in as a utility for users, but is no longer required in KAT services since Excel automation is no longer the server functionaliy used to process spreadsheets.

Back to [Features listing](#features).
