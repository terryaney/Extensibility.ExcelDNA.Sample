# Excel-DNA Add-In

An Excel-DNA add-in for the KAT teams that uses many of the features provided by Excel-DNA along with how I overcame many, not so obvious, obstacles.

## TODO

1. Implement ribbon 'state'
	1. Get 'app settings' working, note: AddInSettings don't work in debug mode :(
		1. https://excel-dna.net/docs/guides-advanced/user-settings-and-the-xllconfig-file
	1. Get 'State Coded'
	1. Ribbon_GetContent - is hacked to get it working
1. Implement Ribbon handlers
	1. exportMappedxDSData - Rename this better after you figure out what it is doing
	1. [Using Windows Forms](https://groups.google.com/g/exceldna/c/84IIhcdAPRk/m/8cRtFOvvAAAJ)	
1. [Custom Intellisense/Help Info](https://github.com/Excel-DNA/IntelliSense/issues/21) - read this and linked topic to see what's possible
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

## Features

### Ribbon Organization

The use of the `IRibbonUI` in the KAT Tools add-in is quite extensive.  There is state management of CustomUI elements via ribbon events, CustomUI element handlers, dynamic menus, and dynamic images to name a few.  In this section I will describe some of the challenges I faced with `IRibbonUI` and how I overcame them.

1. **Helpful Documenation Links**
	1. [CustomUI Reference](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43)
	1. [imageMso Reference](https://codekabinett.com/download/Microsoft-Office-2016_365-imageMso-Gallery.pdf)
1. **Sheer Amount of Code**
	1. Given the amount of methods I had to implement to provide all the required functionality, the lines of code became quite overwhelming (at least given the way I organized the code).  To help alleviate this, I used [partial classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods#partial-classes) just as an organizational tool.  This made it easier to find and maintain the code for me.  Additionally, to make this separation easier to manage in the (Solution) Explorer side bar, I would suggest to enable file nesting.<sup>1</sup>
	1. I also used method prefixes that matched the CustomUI `group.id` as well to make code navigation easier (via `CTRL+T` keyboard shortcut).  For example, for my group with an `id` of `Navigation`, the methods all have the prefix `Navigation_`.

<sup>1</sup> To enable file nesting in Visual Studio Code, add the following to your `settings.json` file:

```json
{
	"explorer.fileNesting.patterns": {
		"*.cs": "${capture}.*.cs"
	}
}
```

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