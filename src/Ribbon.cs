using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using System.Reflection;
using System.Runtime.InteropServices;

namespace KAT.Extensibility.Excel.AddIn;

/// <summary>
/// TODO: Add a description of the Ribbon class here.
/// 
/// Additionally, this class was converted to a partial class for easier maintenance and readability due the the amount
/// of code needed to support all 'ribbon functionality' that is required for the add-in.
/// 
/// 1. The Ribbon.RibbonEvents.cs partial class contains all the events for ribbon elements (i.e. OnLoad, GetVisible, GetEnabled, etc.)
/// 2. The Ribbon.Handlers.*.cs partial class files contain ribbon handlers for each 'group' specified in the CustomUI ribbon xml.  
/// </summary>
[ComVisible( true )]
public partial class Ribbon : ExcelRibbon
{
	private Image auditShowLogImage = null!;

	public override string GetCustomUI( string RibbonID )
	{
		var assembly = Assembly.GetExecutingAssembly();

		using var stream = assembly.GetManifestResourceStream( "KAT.Extensibility.Excel.AddIn.Resources.Ribbon.xml" )!;
		using var reader = new StreamReader( stream );
		var customUi = reader.ReadToEnd();

		using var s = assembly.GetManifestResourceStream( "KAT.Extensibility.Excel.AddIn.Resources.ShowScriptBlockMark.png" )!;
		auditShowLogImage = Image.FromStream( s );

		return customUi;
	}

	public override void OnConnection( object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom )
	{
		// TODO: Need to hook up event handlers...

		base.OnConnection( Application, ConnectMode, AddInInst, ref custom );
	}

	public override void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
	{
		// TODO: Need to unhook event handlers...
		
		base.OnDisconnection( RemoveMode, ref custom );
	}
}