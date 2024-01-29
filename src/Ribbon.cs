using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using System.Reflection;
using System.Runtime.InteropServices;

namespace KAT.Extensibility.Excel.AddIn;

[ComVisible( true )]
public class Ribbon : ExcelRibbon
{
	public override string GetCustomUI( string RibbonID )
	{
		var resourceName = "KAT.Extensibility.Excel.AddIn.Resources.Ribbon.xml";
		var assembly = Assembly.GetExecutingAssembly();
		using var stream = assembly.GetManifestResourceStream( resourceName )!;
		using var reader = new StreamReader( stream );
		return reader.ReadToEnd();
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

	public void OnButtonPressed( IRibbonControl control )
	{
		MessageBox.Show( "Hello from control " + control.Id );
	}
}