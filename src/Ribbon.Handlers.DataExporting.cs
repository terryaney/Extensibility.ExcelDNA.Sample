using System.Diagnostics;
using System.Text.Json;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Domain.IO;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void DataExporting_ExportXmlData( IRibbonControl _ ) => ExportData( true );
	public void DataExporting_ExportJsonData( IRibbonControl _ ) => ExportData( false );

	public void ExportData( bool isXml )
	{
		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			var owner = new NativeWindow();
			owner.AssignHandle( new IntPtr( application.Hwnd ) );

			var selection = ExcelApi.Selection;
			selection = string.IsNullOrEmpty( selection.GetText() )
				? "DataExport".GetReferenceOrNull( selection.SheetName() ) ?? selection
				: selection;

			if ( string.IsNullOrEmpty( selection.GetText() ) )
			{
				MessageBox.Show( "To export data, you must select the first column header of either a single sheet export configuration (usually the Auth ID) or the 'Sheet' configuration cell of a multi-sheet export configuration.", "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
				return;
			}

			var dataExportPath = AddIn.Settings.DataExport.Path;
			var exportPath = string.IsNullOrWhiteSpace( dataExportPath )
				? application.ActiveWorkbook.Path
				: dataExportPath;

			var ext = isXml ? ".xml" : ".json";

			var ws = application.ActiveWorksheet();
			var clientName = ws.RangeOrNull<string>( "ClientName" );
			var outputFile =
				ws.RangeOrNull<string>( "OutputFile" ) ??
				Path.Combine( exportPath, Path.GetFileNameWithoutExtension( application.ActiveWorkbook.Name ) + ext );

			var config = GetWindowConfiguration( nameof( ExportData ) );
			using var exportData = new ExportData( config );

			var info = exportData.GetInfo( clientName, outputFile, AddIn.Settings.DataExport.AppendDateToName, isXml, owner );

			if ( info == null ) return;

			var sw = Stopwatch.StartNew();
			
			SaveWindowConfiguration( nameof( ExportData ), info.WindowConfiguration );

			var exportFiles = DataExport.Provider.Export( selection, info );

			if ( exportFiles != null && !isXml )
			{
				var client = info.ClientName;
				var authIdToExport = info.AuthIdToExport;

				foreach ( var file in exportFiles )
				{
					using ( var fs = File.Create( Path.Combine( Path.GetDirectoryName( file )!, Path.GetFileNameWithoutExtension( file ) + ".json" ) ) )
					using ( var writer = new Utf8JsonWriter( fs, new JsonWriterOptions { Indented = true } ) )
					using ( var reader = new xDataDefReader( file ) )
					{
						if ( string.IsNullOrEmpty( authIdToExport ) )
						{
							writer.WriteStartArray();
						}

						while ( reader.Read() )
						{
							var xDataDef = reader.xDataDef!;
							var authId = xDataDef.AuthId();

							if ( string.IsNullOrEmpty( authIdToExport ) || string.Compare( authIdToExport, authId, true ) == 0 )
							{
								writer.WriteStartObject();

								writer.WritePropertyName( "AuthID" );
								writer.WriteStringValue( authId );

								writer.WritePropertyName( "Client" );
								writer.WriteStringValue( client );

								writer.WritePropertyName( "Profile" );
								writer.WriteStartObject();

								foreach ( var f in xDataDef.Elements( "Profile" ).Elements().Where( e => e.Attribute( "delete" ) == null ) )
								{
									writer.WritePropertyName( f.Name.LocalName );
									writer.WriteStringValue( (string)f );
								}

								writer.WriteEndObject();

								var historyTypes =
									xDataDef.Elements( "HistoryData" ).Elements( "HistoryItem" )
										.Where( h => h.Attribute( "hisClear" ) == null )
										.GroupBy( h => h.hisType() );

								writer.WritePropertyName( "History" );
								writer.WriteStartObject();

								foreach ( var historyType in historyTypes )
								{
									writer.WritePropertyName( historyType.Key );
									writer.WriteStartArray();

									foreach ( var row in historyType )
									{
										writer.WriteStartObject();

										foreach ( var f in row.Elements().Where( e => e.Attribute( "delete" ) == null ) )
										{
											writer.WritePropertyName( f.Name.LocalName );
											writer.WriteStringValue( (string)f );
										}

										writer.WriteEndObject();
									}

									writer.WriteEndArray();
								}

								writer.WriteEndObject();

								writer.WriteEndObject();
							}
						}

						if ( string.IsNullOrEmpty( authIdToExport ) )
						{
							writer.WriteEndArray();
						}
					}

					File.Delete( file );
				}
			}
			MessageBox.Show( string.Format( "Data successfully exported in {0:0.000} seconds.", sw.Elapsed.TotalSeconds ), "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information );
		} );
	}

	public void DataExporting_ExportResultDocGenXml( IRibbonControl control )
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
}
