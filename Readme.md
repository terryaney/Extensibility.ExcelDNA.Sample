# Excel-DNA Add-In

An Excel-DNA add-in for Excel demonstrates many of the features provided by Excel-DNA along with how I overcame many, not so obvious, obstacles.

## TODO

1. Implement ribbon 'state'
	1. Get 'app settings' working, note: AddInSettings don't work in debug mode :(
	1. ribbon_GetContent - is hacked to get it working
1. Implement Ribbon handlers
	1. exportMappedxDSData - Rename this better after you figure out what it is doing
	1. [Using Windows Forms](https://groups.google.com/g/exceldna/c/84IIhcdAPRk/m/8cRtFOvvAAAJ)	
1. Why do I have Registerfunctions?
1. [Path of Xll](https://groups.google.com/g/exceldna/c/1rScvDdeVOk) - `XlCall.Excel( XlCall.xlGetName )` to get the name of the add-in

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