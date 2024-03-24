using ExcelDna.Integration.CustomUI;
using System.Xml.Linq;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	// Need reference to IRibbonUI so I can change the enable/disable state of buttons and 
	// dynmically update the ribbon (i.e. debug CalcEngine dropdown).
	private IRibbonUI ribbon = null!;

	public void ribbon_OnLoad( IRibbonUI ribbon )
	{
		this.ribbon = ribbon;
	}

	public bool ribbon_GetVisible( IRibbonControl control )
	{
		switch ( control.Id )
		{
			default: return true;
		}
	}

	public bool ribbon_GetEnabled( IRibbonControl control )
	{
		switch ( control.Id )
		{
			default: return true;
		}
	}

	public string? ribbon_GetScreentip( IRibbonControl control )
	{
		switch ( control.Id )
		{
			default: return null;
		}
	}

	public string? ribbon_GetContent( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "debugCalcEngines":
			{
				// TODO: Need real settings
				var historyAuthor = (string?)null; // AddInSettings.Settings?.HistoryAuthor;
				if ( string.IsNullOrEmpty( historyAuthor ) )
				{
					historyAuthor = null;
				}
				if ( string.IsNullOrEmpty( historyAuthor ) && System.Diagnostics.Debugger.IsAttached )
				{
					// TODO: AddInSettings don't work in debug mode :(
					historyAuthor = "terry.aney";
				}
				if ( historyAuthor == "thomas.aney" )
				{
					historyAuthor = "tom.aney";
				}

				var userDirectory = new[] { "terry.aney", "tom.aney" }.Contains( historyAuthor )
					? $"btr.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}"
					: $"conduent.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}";

				var ceName = "Conduent_Nexgen_Home_SE.xlsm"; // workbookState.ManagementName;
				var debugFiles = // LibraryHelpers.GetDebugCalcEngines( userDirectory, ceName );
					Enumerable.Range( 0, 3 ).Select( i => new { Name = $"{Path.GetFileNameWithoutExtension( ceName )} Debug at {DateTime.Now.AddHours( -1 * ( i + 1 ) ):yyyy-MM-dd hh-mm-sstt} for 011391001.xlsm" } );

				XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
				var menu =
					new XElement( ns + "menu",
						debugFiles.Any()
							? debugFiles.Select( ( f, i ) =>
								new XElement( ns + "button",
									new XAttribute( "id", "managementDownloadFile" + i ),
									new XAttribute( "keytip", i ),
									new XAttribute( "imageMso", "CustomizeXSLTMenu" ),
									new XAttribute( "onAction", "DownloadDebugFile" ),
									new XAttribute( "tag", $"{userDirectory}|{f.Name}" ),
									new XAttribute( "label", f.Name )
								)
							)
							: new[] {
									new XElement( ns + "button",
										new XAttribute( "id", "managementDownloadFile0" ),
										new XAttribute( "imageMso", "CustomizeXSLTMenu" ),
										new XAttribute( "label", "No files available" ),
										new XAttribute( "enabled", "false" )
									)
							}
						);

				return menu.ToString();

				// foreach ( var c in ribbonStatesDebugCEs ) { ribbon.InvalidateControl( c ); }
			}
			default: return null;
		}
	}

	private int auditShowLogBadgeCount;
	public Bitmap ribbon_GetImage( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "auditShowLog":
			{
				if ( auditShowLogBadgeCount > 0 )
				{
					var flagGraphics = Graphics.FromImage( auditShowLogImage );
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

				return (Bitmap)auditShowLogImage;
			}

			default: throw new ArgumentOutOfRangeException( nameof( control.Id ), $"The id {control.Id} does not support custom image generation." );
		}

	}
}