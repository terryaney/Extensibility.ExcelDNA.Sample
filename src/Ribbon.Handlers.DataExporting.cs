using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void DataExporting_ExportXmlData( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void DataExporting_ExportResultDocGenXml( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void DataExporting_ExportJsonData( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void DataExporting_ExportMappedXmlData( IRibbonControl _ )
	{
		var owner = new NativeWindow();
		owner.AssignHandle( new IntPtr( application.Hwnd ) );

		ExcelAsyncUtil.QueueAsMacro( () =>
		{

			var config = GetWindowConfiguration( nameof( XmlMapping ) );
			using var xmlMapping = new XmlMapping( config );

			var ws = application.ActiveWorksheet();
			var clientName = (string?)ws.RangeOrNull( "ClientName" )?.Offset[ 0, 1 ].Text;
			var inputFile = (string?)ws.RangeOrNull( "InputFile" )?.Offset[ 0, 1 ].Text;
			var outputFile = (string?)ws.RangeOrNull( "OutputFile" )?.Offset[ 0, 1 ].Text;
			
			var info = xmlMapping.GetInfo( clientName, inputFile, outputFile, owner );

			if ( info == null ) return;

			SaveWindowConfiguration( nameof( XmlMapping ), info.WindowConfiguration );

			new XmlMappingService().ExportXmlData( ws, info );
		} );
	}
	public void DataExporting_ExportResultJsonData( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void DataExporting_AuditDataExportHeaders( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}
