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
1. [Extending Optional Parameters and Default Values](#extending-optional-parameters-and-default-values)
1. [My Intellisense Journey](#my-intellisense-journey)
1. [Creating a Setup Program](#creating-a-setup-program)

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


1. **Api Functionality via HttpClient** - The `Camelot.Api.Excel` api project is used both to manage state of the ribbon as well as provide functionality for some of the button events.  See [`WorkbookState.UpdateCalcEngineInfoAsync`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Models/WorkbookState.cs) or [`Ribbon.KatDataStore_DownloadLatestCalcEngine`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Handlers.KatDataStore.cs) for some examples.
2. **Long Running Tasks** - Some ribbon handlers require launching a long running task that can be cancelled if needed, does not block the main Excel threads, and then reports back information to the main thread (i.e. Local Batch Processes).
3. **async/await Support in Handlers** - The ability to run async code in Ribbon button handlers and follow up interactions with Excel COM objects to manipulate the Excel application in some fashion (i.e. `ribbon.Invalidate()` or other `application.*` methods (`application.Workbooks.Open`)).
4. **FileSystemWatcher Event Support** - A `FileSystemWatcher` that is monitoring `appsettings*.json` file(s) for changes triggers change events when necessary that need to reload settings and then invalid the ribbon.  See [AddIn.AutoOpen](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/AddIn.cs) for more detail.
5. **Global Exception Handler** - A global exception handler is added via `ExcelIntegration.RegisterUnhandledExceptionHandler` to provide the ability to catch all formula/calculation exceptions.  In the registered [UnhandledExceptionHandler](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/AddIn.cs) handler, the error is logged and a ribbon badge is updated with current count of errors.  See the [ExcelIntegration.RegisterUnhandledExceptionHandler](#excelintegrationregisterunhandledexceptionhandler) section for more detail.

After a [long discussion](https://groups.google.com/g/exceldna/c/_pKphutWbvo/m/aaH_Z-QZAAAJ) and answers from Govert, I finally wrapped my head around support for async/await in Excel-DNA.  In the simplest terms, the `ExcelAsyncUtil.QueueAsMacro` method **must be** used to [access all 'UI' and 'Excel COM' objects on the main Excel thread](https://groups.google.com/g/exceldna/c/_pKphutWbvo/m/-Xl3imi5AAAJ) (vs the calculation thread(s)) otherwise unexpected results will occur - most notably/frustrating is Excel not closing gracefully after improper code flow.  To initial any async/await code from the add-in, the use of `Task.Run( () => { } )` is used.  I created [RunRibbonTask](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.cs) as a helper to wrap this that also logs exceptions to the `ExcelDna.Logging.LogDisplay` window and toggles the main UI thread's cursor.

At times, following the pattern described above required a different approach/flow of code than I would have preferred.

1. When modals or `MessageBox` messages are displayed I still wanted Excel to be the 'owner' of the window.  To accomplish this, I followed the [Creating a Threaded Modal Dialog](https://excel-dna.net/docs/tips-n-tricks/creating-a-threaded-modal-dialog) example.  I didn't follow the exact pattern in the sample because I was surprised that `Application.Hwnd` was accessed on a new thread.  I've [asked Excel-DNA](https://groups.google.com/g/exceldna/c/lY5e-CtgFiI) and will update this section when I get an answer.  You can see a demonstration of my suggested pattern in [ExportGlobalTables](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Handlers.ConfigurationExporting.cs).  This pattern was used whenever a dialog was displayed from with the context of `RunRibbonTask`.

2. Nested calls of `RunRibbonTask` and `QueueAsMacro`. In some cases, I needed to call `RunRibbonTask` to start an async method, then later call `QueueAsMacro` to access UI/COM components, then start another `RunRibbonTask` to support async support.  This was done in the `ExportGlobalTables`.
	1. Use `RunRibbonTask` to display my custom form/dialog.  Almost all my dialogs were prompting for user credentials that would be used, and async support was required to decrypt a local secret.
	1. If the dialog was confirmed, I would then call `QueueAsMacro` to access `UI` components or `Excel COM` objects to perform my task (in this case multiple times, first to toggle the cursor, then to do the work).
	1. After performing the above mentioned work (building an api payload from current worksheet information), I needed to call async api methods, so started a nested `RunRibbonTask` delegate.
	1. From this context, I needed to call `QueueAsMacro` to display a message box and conditionally close a `Workbook` object.

Obviously, this is not as clean as straight forward linear code, but it is not terribly difficult to following.  Mostly it is just a matter of indentation in the code.

Note: I started to refactor code that used current password to instead have a variable that was ready to use without async requirement (except upon password updates) but all dialogs usually resulted in async work when confirmed, so I didn't bother.  *At time of writing*, this example was only 'nested' flow problem that I had.  I may revisit.

3. I have some helpers in `Ribbon.cs` that wrap some common UI thread access `QueueAsMacro` calls (i.e. `InvalidateRibbon`, `ClearStatusBar`, etc.) and there are times when the 'common helper' was only part of the work flow so I called the common helper then immediately made another call to `QueueAsMacro` to access other UI/COM components.  Govert mentioned that these multiple calls should be ok.  This was done in the [UploadCalcEngineToManagementSiteAsync](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Events.Excel.cs) method where I used multiple calls to `SetStatusBar` along with direct calls to `QueueAsMacro`.

4. The`System.Windows.Forms.OpenFileDialog` control (and other Windows Forms controls) requires that the thread it is created and used on is marked as a Single Threaded Apartment (STA) thread. This is typically done by adding the `[STAThread]` attribute to the Main method of your application.  However, since I'm not in a position to control the Main method, I need to make sure that any dialogs that use this control are opened within a `QueueAsMacro` delegate.  See [EmailBlast](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Views/EmailBlast.cs) for example.

There was one location where I did not follow the best practices advice from Govert.  In the [Ribbon_GetContent](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/Ribbon.Events.cs) handler, I call our async api to get list of debug CalcEngines to populate a dropdown ribbon menu.  Govert [described](https://groups.google.com/g/exceldna/c/_pKphutWbvo/m/ae_0hSR_AAAJ) how any 'blocking' technique was not the best approach and that I should return immediately from this handler providing 'empty' menu or something indicating 'work is being' done.  I never tried to change my flow, so I'm not sure how it would have worked/looked in terms of 're-populating' the dropdown to remove the 'empty/working' menu item with the updated list (or none available) of CalcEngines menus.

Instead, I used [The Thread Pool Hack](https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development#the-thread-pool-hack) to run the async method from within synchronous code and block and wait.  I decided it was acceptable to have user/Excel blocked/waiting for a response since they understand that it is querying our api and are expecting a slight delay.

**Original Reference Links for Thread Context and async/await Issues**

1. https://groups.google.com/g/exceldna/c/_pKphutWbvo - question asking about my different scenarios
1. https://groups.google.com/g/exceldna/c/ILgL-dW47A4/m/9HrOyClJAQAJ - Thread about ensuring Excel shuts down properly.
1. https://stackoverflow.com/a/68303070/166231 - Stephen Cleary's answer about async/await best practices.
1. https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development - Async Programming article by Stephen Cleary.

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

### Extending Optional Parameters and Default Values

Just wanted to note an updated method of handling optional parameters and default values in Excel-DNA.  See [`OptionalValues.Check`](https://github.com/terryaney/Extensibility.Camelot.Excel.KAT/blob/main/src/[Functions]/OptionalValues.cs) for the implementation.  It is based on Excel-DNA guide of [Optional Parameters and Default Values
](https://excel-dna.net/docs/guides-basic/optional-parameters-and-default-values) but it reduces the functions into a single C# generic function and also handles converting Excel 'numbers' (always doubles) to `int` if that is the requested type.  The signature of the function includes the name of the variable simply for diagnostic purposes.

```csharp
// Function Signature
static T Check<T>( this object? arg, string argumentName, T defaultValue ) { }

// Usage
[ExcelFunction()]
public static int MapOrdinal( 
	[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null
) => defaultValue.Check( nameof( defaultValue ), 1 );
```

Back to [Features listing](#features).

### My Intellisense Adventure

Making intellisense has been made much easier as compared to when I originally implemented intellisense with a separate ExcelDNA Intellisense Host add-in that was required.  That being said, there are still some frustrations to overcome due to limitations in Excel (and not within ExcelDNA itself).  The 'Function Wizard' has similar issues that are discussed below as well.

There are probably a few discussions on this topic, but the bottom line is Excel will truncate function and argument descriptions if they exceed 255 characters.  ExcelDNA has been accomodating and has provided a few workarounds for Excel's limitations.

- [.intellisense.xml](https://groups.google.com/g/exceldna/c/BaoE8psmTxo) has priority over compiled descriptions from the `ExcelFunction` and `ExcelArgument` attributes.
- [IntelliSense shows full descriptions](https://github.com/Excel-DNA/ExcelDna/issues/85) even though Excel truncates them.

Given these capabilities, one could probably just manage an `.intellisense.xml` file and not use the `ExcelFunction` and `ExcelArgument` attributes for documentation.  I didn't want truncated descriptions to occur in the Excel Functionn Wizard even though intellisense displayed them correctly, so I had to come up with a solution that would support 'short(er)' descriptions for Excel Function Wizard and full descriptions in intellisense.

In addition to function and attribute description limitations, there is a limitation in the Function Wizard with regard to the length of the names of all the arguments in the function.  If exceeded, the Function Wizard will simply truncate an argument name and/or full arguments from being displayed.  So again, I needed a solution for this problem as well.  What I came up with was this:

1. Extend `ExcelFunction` and `ExcelArgument` attributes to include a 'short' name and description and a 'full' name and description so the Function Wizard could display 'all' information (albiet in abbreviated fashion) while IntelliSense and 'help' could display the full information.
1. Function Wizard uses compiled `ExcelFunctionAttribute.Description`, `ExcelArgumentAttribute.Name`, and `ExcelArgumentAttribute.Description` properties.
1. Using a post build script, enerate an `.intellisense.xml` file that includes the 'full' names and descriptions for IntelliSense to display and MarkDown documentation files for functions.
1. During `AutoOpen`, register my functions and update the `ExcelFunctionAttribute.HelpTopic` property to point to the hosted MarkDown file in github.

My extended attributes are as follows:

```csharp
public class KatExcelFunctionAttribute : ExcelFunctionAttribute
{
	public string? Returns = null; // Used in MarkDown documentation for more explicit information on the return type
	public string? Summary = null; // Full/Long description of the function used in IntelliSense and MarkDown documentation
	public string? Remarks = null; // Used in MarkDown documentation for additional remarks about the funciton
	public string? Exceptions = null; // Used in MarkDown documentation to explain possible Exception types that are thrown.
	public string? Example = null; // Used in MarkDown documentation to provide an example of the function usage
	public bool CreateDebugFunction = false; // Used to create a debug version of the function that returns either the valid return value or Exception.Message as text (instead of #VALUE)
	
	public KatExcelFunctionAttribute() { }

	public KatExcelFunctionAttribute( string description )
	{
		Description = description; // Base description shown in Function Wizard
	}
}
public class KatExcelArgumentAttribute : ExcelArgumentAttribute
{
	public string? Summary = null; // Full/Long description of the argument used in IntelliSense and MarkDown documentation
	public string? DisplayName = null; // Full/Long name of argument used in IntelliSense and MarkDown documentation
	public Type? Type = null; // Used in MarkDown documentation to provide the type of an (optional) argument (instead of just `object`)
	public string? Default = null; // Used in MarkDown documentation to provide the default value of an (optional) argument

	public KatExcelArgumentAttribute() { }

	public KatExcelArgumentAttribute( string description )
	{
		Description = description; // Base description shown in Function Wizard
	}
}
```

Here are some usage examples of the extended attributes:

```csharp
	[KatExcelFunction( 
		Category = "Formatting", 
		Description = "Formats a numeric value to a string representation using the specified format and culture-specific format information.",
		Returns = "The string representation of the value of this instance as specified by `format` and `culture`.",
		Remarks = @"The `BTRNumberFormat` method is similar to Excel's `Format()` function with the exception that `BTRNumberFormat` can dynamically format a number based on `culture` using the same `format` string.

*See Also*
[Standard Numeric Format Strings](http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx)
[Custom Numeric Format Strings](http://msdn.microsoft.com/en-us/library/0c899ak8(v=vs.110).aspx)",
		Example = @"This sample shows how to format a numeric value to currency format with a single format string but changes based on culture.

\```
// Assume this comes from the iCurrentCulture input.
string culture = ""en-US"";
// Assume this comes from a calculated result.
double value = 10.5;
// currencyValue would have ""$10.50"" for a value.
string currencyValue = BTRNumberFormat( value, ""c"", culture );
// If culture was French...
culture = ""fr-FR"";
// currencyValue would have ""10,50 €"" for a value.
currencyValue = BTRNumberFormat( value, ""c"", culture );
\\\"
	)]
	public static string BTRNumberFormat(
		[ExcelArgument( "The numeric value to apply formatting to." )]
		double value,
		[ExcelArgument( "The C# string format to apply.  View the function's help for more detail on possible values." )]
		string format,
		[KatExcelArgument(
			Description = "Optional.  The culture name in the format languagecode2-country/regioncode2 (default of `en-US`).  See 'National Language Support (NLS) API Reference' for available names.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null
	)
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		return Utility.LocaleFormat( value, format, cultureArg );
	}
```

This function generates [this MarkDown documentation](#https://github.com/terryaney/Documentation.Camelot/blob/main/RBLe/RBLeFormatting.BTRNumberFormat.md).

The post build script is simply a [LINQPad](#https://www.linqpad.net/) script that I run via the command line.  It is pretty trivial C# code that can be found [here](.vscode/DnaDocumentation/Generate.linq), but the meat of the script is simply getting appropriate functions and examining custom attributes.

```csharp
var assembly = typeof(Ribbon).Assembly;
var info =
	assembly.GetTypes()
		.SelectMany(t => t.GetMethods())
		.Where(m => (m.GetCustomAttribute<KatExcelFunctionAttribute>() ?? m.GetCustomAttribute<ExcelFunctionAttribute>()) != null)
		.Select(m =>
		{
			var katFunc = m.GetCustomAttribute<KatExcelFunctionAttribute>();
			var dnaFunc = (katFunc as ExcelFunctionAttribute) ?? m.GetCustomAttribute<ExcelFunctionAttribute>()!;

			return new
			{
				Name = dnaFunc.Name ?? m.Name,
				Category = dnaFunc.Category,
				Description = katFunc?.Summary ?? dnaFunc.Description,
				Returns = katFunc?.Returns,
				Remarks = katFunc?.Remarks,
				Example = katFunc?.Example,
				HelpTopic = dnaFunc.HelpTopic,
				CreateDebugFunction = katFunc?.CreateDebugFunction ?? false,
				Arguments =
					m.GetParameters()
						.Select(p =>
						{
							var katArg = p.GetCustomAttribute<KatExcelArgumentAttribute>();
							var dnaArg = (katArg as ExcelArgumentAttribute) ?? p.GetCustomAttribute<ExcelArgumentAttribute>();
							return new
							{
								Name = katArg?.DisplayName ?? dnaArg?.Name ?? p.Name!,
								Description = katArg?.Summary ?? dnaArg?.Description ?? "TODO: Document this parameter.",
								Type = katArg?.Type ?? p.ParameterType,
								IsOptional = p.IsOptional,
								DefaultValue = katArg?.Default ?? p.DefaultValue?.ToString()
							};
						})
			};
		})
		.ToArray();

XNamespace ns = "http://schemas.excel-dna.net/intellisense/1.0";

var fileName = isDebug ? "KAT.Extensibility.Excel.Debug.intellisense.xml" : "KAT.Extensibility.Excel.intellisense.xml";
WriteLine($"Generating {fileName}...", true);
var intelliSense =
	new XElement(ns + "IntelliSense",
		new XElement(ns + "FunctionInfo",
			info.Select(i =>
				new XElement(ns + "Function",
					new XAttribute("Name", i.Name),
					new XAttribute("Description", i.Description),
					i.HelpTopic != null ? new XAttribute("HelpTopic", i.HelpTopic) : null,
					i.Arguments.Select(a =>
						new XElement(ns + "Argument",
							new XAttribute("Name", a.Name),
							new XAttribute("Description", a.Description)
						)
					)
				)
			)
		)
	);

// Omitting MarkDown file generation for brevity
```

Back to [Features listing](#features).

### Creating a Setup Program

Back to [Features listing](#features).
