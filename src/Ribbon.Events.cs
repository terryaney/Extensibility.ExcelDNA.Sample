using ExcelDna.Integration.CustomUI;
using System.Xml.Linq;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"btrRBLe" => showRibbon,
			_ => true,
		};
	}

	public bool Ribbon_GetEnabled( IRibbonControl control )
	{
		switch ( control.Id )
		{
			default: return true;
		}
	}

	public string? Ribbon_GetScreentip( IRibbonControl control )
	{
		switch ( control.Id )
		{
			default: return null;
		}
	}

	public string? Ribbon_GetContent( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "debugCalcEngines":
			{
				var historyAuthor = AddIn.Settings.SaveHistory.Name;
				if ( string.IsNullOrEmpty( historyAuthor ) )
				{
					historyAuthor = null;
				}
				if ( historyAuthor == "thomas.aney" )
				{
					historyAuthor = "tom.aney";
				}

				var userDirectory = new[] { "terry.aney", "tom.aney" }.Contains( historyAuthor )
					? $"btr.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}"
					: $"conduent.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}";

				var ceName = "Conduent_Nexgen_Home_SE.xlsm"; // TODO: workbookState.ManagementName;
				var debugFiles = // TODO: LibraryHelpers.GetDebugCalcEngines( userDirectory, ceName );
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
			}
			default: return null;
		}
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

			default: throw new ArgumentOutOfRangeException( nameof( control ), $"The id {control.Id} does not support custom image generation." );
		}
	}
}