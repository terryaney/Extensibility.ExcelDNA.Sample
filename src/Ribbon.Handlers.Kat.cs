using System.Diagnostics;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Kat_BlastEmail( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void Kat_ShowLog( IRibbonControl? _ )
	{
		ExcelDna.Logging.LogDisplay.Show();
		auditShowLogBadgeCount = 0;
		ribbon.InvalidateControl( "katShowDiagnosticLog" );
	}

	public void Kat_OpenHelp( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void Kat_RefreshRibbon( IRibbonControl _ ) => Application_WorkbookActivate( application.ActiveWorkbook );

	public void Kat_HelpAbout( IRibbonControl _ )
	{
		// https://stackoverflow.com/a/14498889/166231 - discusses versioning in xll
		var fvi = FileVersionInfo.GetVersionInfo( AddIn.XllName );
		var versionParts = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyInformationalVersionAttribute>().First().InformationalVersion.Split( '.' );
		var version = versionParts.Length == 3 && versionParts[ 2 ].Contains( '+' )
			? string.Join( ".", versionParts.Take( 2 ).Concat( new [] { $"{versionParts[ 2 ].Split( '+' )[ 0 ]}+{versionParts[ 2 ].Split( '+' )[ 1 ][ ..6 ]}" } ) )
			: string.Join( ".", versionParts );

		MessageBox.Show(
			$"KAT Excel Add-In: {version}{Environment.NewLine}ExcelDna.AddIn Package: {fvi.FileVersion}",
			"KAT Excel Add-In",
			MessageBoxButtons.OK,
			MessageBoxIcon.Information
		);
	}
}