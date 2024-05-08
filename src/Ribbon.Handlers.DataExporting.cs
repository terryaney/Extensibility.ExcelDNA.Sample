using System.Linq.Expressions;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

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

			var info = xmlMapping.GetInfo( owner );

			if ( info == null ) return;

			SaveWindowConfiguration( nameof( XmlMapping ), info.WindowConfiguration );

			var mappingExpressions = new Dictionary<string, Func<XmlContext, object>>();
			var specification = GetXmlMappingConfiguration( ws );

			var preFile = PreprocessXmlMappingInputFile( info.ClientName, info.InputFile );

			try
			{
				var settings = new XmlWriterSettings
				{
					Encoding = Encoding.UTF8,
					// OmitXmlDeclaration = true,
					Indent = System.Diagnostics.Debugger.IsAttached
				};

				var xDataDefConfig =
					specification
						.Elements( "Profile" )
						.Elements().Where( e => (string?)e.Attribute( "IsAuthId" ) == "Y" )
						.Select( e => new
						{
							ClientElementName = ( (string)e.Parent!.Attribute( "PathToProfile" ) )!.Split( '/' ).Last(),
							AuthIdFieldName = (string)e.Attribute( "ClientField" )!
						} )
						.First();

				var totalRows = 0;

				var profileFieldMappings = specification.Element( "Profile" )!.Elements().ToArray();
				var historyFieldMappings = specification.Elements().Where( m => m.Name.LocalName != "Profile" ).ToArray();

				using ( var xw = XmlWriter.Create( info.OutputFile, settings ) )
				using ( var reader = new XmlTextReader( preFile ?? info.InputFile ) )
				{
					xw.WriteStartElement( "xDataDefs" );

					var configuration =
						new XElement( "Configuration",
							new XElement( "Profile",
								profileFieldMappings.Select( f => new XElement( f.Name.LocalName ) )
							),
							historyFieldMappings.Select( h =>
								new XElement( "HistoryData",
									new XAttribute( "type", h.Name.LocalName ),
									h.Elements().Select( f => new XElement( f.Name.LocalName ) )
								)
							)
						);

					configuration.WriteTo( xw );

					while ( reader.Read() )
					{
						if ( reader.LocalName == xDataDefConfig.ClientElementName )
						{
							totalRows++;
							var clientElement = XElement.Load( reader.ReadSubtree() );
							var authId = GetMappingValue( clientElement, xDataDefConfig.AuthIdFieldName )!;

							var xDataDefElement =
								new XElement( "xDataDef",
									new XAttribute( "id-auth", authId ),
									new XElement( "Profile",
										from f in profileFieldMappings
										select GetMappingValue( clientElement, f, mappingExpressions )
									),
									new XElement( "HistoryData",
										from h in historyFieldMappings
										select GetHistoryRows( h, clientElement, mappingExpressions )
									)
								);

							xDataDefElement.WriteTo( xw );
						}
					}

					xw.WriteEndElement();
				}

				PostprocessXmlMappingOutputFile( info.ClientName, info.OutputFile, totalRows );

				MessageBox.Show( "Export complete." );
			}
			finally
			{
				if ( !string.IsNullOrEmpty( preFile ) )
				{
					File.Delete( preFile );
				}
			}
		} );
	}

	private const string AuthIdElement = "AuthIdElement";
	private const string PathToProfileElement = "PathToProfileElement";
	private const string MappingLayouts = "MappingLayouts";
	private const string clientField = "Client Field";
	private const string xDSField = "xDS Field";
	private const string isDateField = "Is Date";
	private const string deleteIfBlank = "Delete If Blank";
	private const string expression = "Expression";

	private static XElement GetXmlMappingConfiguration( MSExcel.Worksheet worksheet )
	{
		var authIdElement = worksheet.Range[ AuthIdElement ].GetReference().Offset( 0, 1 ).GetText();
		var pathToProfileElement = worksheet.Range[ PathToProfileElement ].GetReference().Offset( 0, 1 ).GetText()!;
		var mappingLayout = worksheet.Range[ MappingLayouts ].GetReference();

		var specification = new XElement( "Specification" );

		string? mappingLayoutText = null;

		while ( !string.IsNullOrEmpty( mappingLayoutText = mappingLayout.GetText() ) && mappingLayout.RowFirst < 10000 )
		{
			if ( mappingLayoutText == "xDS Table" )
			{
				var currentxDSTable = mappingLayout.Offset( 0, 1 ).GetText() ?? "Profile";
				var pathToElement = mappingLayout.Offset( 1, 1 ).GetText()!;

				var headers =
					mappingLayout
						.Offset( 2, 0 )
						.Extend( mappingLayout.Offset( 2, 0 ).End( DirectionType.ToRight ) )
						.GetValues<string>()
						.Select( ( h, i ) => new { Header = h!, Ordinal = i } )
						.ToDictionary( h => h.Header, h => h.Ordinal );

				ExcelReference? getSpecColumn( string name, ExcelReference reference )
				{
					if ( !headers.TryGetValue( name, out var value ) ) return null;
					return reference.Offset( 0, value );
				}

				string? getSpecValue( string name, ExcelReference reference ) =>
					getSpecColumn( name, reference )?.GetText();

				XAttribute? getSpecAttribute( string name, ExcelReference reference )
				{
					var value = getSpecValue( name, reference );

					return !string.IsNullOrEmpty( value )
						? new XAttribute( name.Replace( " ", "" ), value )
						: null;
				}

				XAttribute? getSpecFormulaAttribute( string name, ExcelReference reference )
				{
					var specColumn = getSpecColumn( name, reference );
					var value = specColumn?.GetFormula() ?? specColumn?.GetText();

					return !string.IsNullOrEmpty( value )
						? new XAttribute( name.Replace( " ", "" ), value )
						: null;
				}

				var mappingContainer =
					new XElement( currentxDSTable,
						currentxDSTable == "Profile"
							? new XAttribute( "PathToProfile", pathToProfileElement )
							: new XAttribute( "RelativePath", pathToElement )
					);

				specification.Add( mappingContainer );

				mappingLayout = mappingLayout.Offset( 3, 0 );

				var rowOffset = 0;
				ExcelReference? currentRow = null;

				while ( !string.IsNullOrEmpty( ( currentRow = mappingLayout.Offset( rowOffset, 0 ) ).GetText() ) )
				{
					var field = getSpecValue( xDSField, currentRow )!;
					var authIdCheck = getSpecValue( clientField, currentRow );

					mappingContainer.Add(
						new XElement( field,
							getSpecAttribute( clientField, currentRow ),
							new[] { authIdElement, $"@{authIdElement}" }.Contains( authIdCheck ) ? new XAttribute( "IsAuthId", "Y" ) : null,
							getSpecAttribute( isDateField, currentRow ),
							getSpecAttribute( deleteIfBlank, currentRow ),
							getSpecFormulaAttribute( expression, currentRow )
						)
					);

					rowOffset++;
				}
			}

			// Mapping layout current sitting on 'first row', so down twice
			mappingLayout = mappingLayout.End( DirectionType.Down ).End( DirectionType.Down, true );
		}

		return specification;
	}

	private static string? PreprocessXmlMappingInputFile( string groupName, string inputFile )
	{
		var settings = new XmlWriterSettings
		{
			Encoding = Encoding.UTF8,
			// OmitXmlDeclaration = true,
			Indent = System.Diagnostics.Debugger.IsAttached
		};

		if ( groupName == "AnnArbor" )
		{
			var outputFile = inputFile + ".pre.xml";
			string? reportEndDate = null;
			using ( var xw = XmlWriter.Create( outputFile, settings ) )
			using ( var reader = new XmlTextReader( inputFile ) )
			{
				while ( reader.Read() )
				{
					reader.StreamXmlNode( xw );

					if ( reader.NodeType == XmlNodeType.Element && reader.IsStartElement( "Batch" ) )
					{
						reportEndDate = reader.GetAttribute( "ReportEndDate" );
					}
					else if ( reportEndDate != null && reader.NodeType == XmlNodeType.Element && reader.IsStartElement( "Member" ) && string.IsNullOrEmpty( reader.GetAttribute( "ReportEndDate" ) ) )
					{
						xw.WriteAttributeString( "ReportEndDate", reportEndDate );
					}
				}
			}

			return outputFile;
		}

		return null;
	}

	private static IEnumerable<XElement> GetHistoryRows( XElement historyConfiguration, XElement clientElement, Dictionary<string, Func<XmlContext, object>> mappingExpressions )
	{
		var indexElement = historyConfiguration.Element( "index" )!;

		return 
			from h in clientElement.XPathSelectElements( (string?)historyConfiguration.Attribute( "RelativePath" ) ?? "." )
			let index = (string)GetMappingValue( h, indexElement, mappingExpressions )!
			select new XElement( "HistoryItem", 
				new XAttribute( "hisType", historyConfiguration.Name.LocalName ), 
				new XAttribute( "hisIndex", index ),
				from f in historyConfiguration.Elements()
				select GetMappingValue( h, f, mappingExpressions )
			);
	}

	private static string? GetMappingValue( XElement container, string fieldName )
	{
		if ( container == null ) return null;

		try
		{
			var value = container.XPathSelectValue( fieldName, "" );

			return ( (string)value ).Trim();
		}
		catch ( Exception ex )
		{
			throw new ArgumentOutOfRangeException( $"Invalid fieldName: {fieldName}.  Container: {container}", ex );
		}
	}

	private static XElement? GetMappingValue( XElement container, XElement configuration, Dictionary<string, Func<XmlContext, object>> mappingExpressions )
	{
		var valueContainer = container;

		var clientElement = (string?)configuration.Attribute( "ClientElement" );
		if ( !string.IsNullOrEmpty( clientElement ) && (string?)configuration.Attribute( "IsAuthId" ) != "Y" ) valueContainer = valueContainer.XPathSelectElement( clientElement )!;

		var clientField = (string?)configuration.Attribute( "ClientField" );
		var expression = (string?)configuration.Attribute( "Expression" );

		if ( string.IsNullOrEmpty( clientField ) && string.IsNullOrEmpty( expression ) ) return null;

		var value = 
			!string.IsNullOrEmpty( clientField ) ? GetMappingValue( valueContainer, clientField ) : 
			!string.IsNullOrEmpty( expression ) 
				? GetMappingExpression( string.Format( "{0}:{1}", configuration.Parent!.Name.LocalName, configuration.Name.LocalName ), expression[ 1.. ], mappingExpressions )( new XmlContext { Element = container } ).ToString() 
				: null;

		if ( string.IsNullOrEmpty( value ) )
		{
			return (string?)configuration.Attribute( "DeleteIfBlank" ) == "Y"
				? new XElement( configuration.Name.LocalName, new XAttribute( "delete", 1 ) )
				: null;
		}

		if ( (string?)configuration.Attribute( "IsDate" ) == "Y" ) value = DateTime.Parse( value ).ToString( "yyyy-MM-dd" );

		return new XElement( configuration.Name.LocalName, value );
	}

	private static Func<XmlContext, object> GetMappingExpression( string key, string formula, Dictionary<string, Func<XmlContext, object>> mappingExpressions )
	{
		if ( !mappingExpressions.ContainsKey( key ) ) mappingExpressions.Add( key, Parse( formula ) );
		return mappingExpressions[ key ];
	}

	private static Func<XmlContext, object> Parse( string expressionText )
	{
		return Parse<object>( expressionText );
	}

	private static Func<XmlContext, TResult> Parse<TResult>( string expressionText )
	{
		var scanner = new Scanner();
		var parser = new Parser( scanner );
		var tree = parser.Parse( expressionText );

		var visitor = new XmlMappingVisitor();
		var expression = visitor.GetExpression( tree );

		if ( expression.Type != typeof( TResult ) ) expression = Expression.Convert( expression, typeof( TResult ) );

		var lambda = Expression.Lambda<Func<XmlContext, TResult>>( expression, visitor.Context );

		return lambda.Compile();
	}

	private static void PostprocessXmlMappingOutputFile( string groupName, string outputFile, int totalRows )
	{
		var settings = new XmlWriterSettings
		{
			Encoding = Encoding.UTF8,
			// OmitXmlDeclaration = true,
			Indent = System.Diagnostics.Debugger.IsAttached
		};

		var resultFile = outputFile + ".post.xml";

		using ( var xw = XmlWriter.Create( resultFile, settings ) )
		using ( var reader = new XmlTextReader( outputFile ) )
		{
			xw.WriteStartElement( "xDataDefs" );
			xw.WriteAttributeString( "TotalRows", totalRows.ToString() );

			while ( reader.Read() )
			{
				switch ( reader.LocalName )
				{
					case "Configuration":
						xw.WriteNode( reader.ReadSubtree(), true );
						break;
					case "xDataDef":
						var xDataDef = XElement.Load( reader.ReadSubtree() );

						if ( groupName == "AnnArbor" )
						{
							var statusToRemove = xDataDef
													.Elements( "HistoryData" )
													.Elements( "HistoryItem" )
													.Where( e => e.hisType() == "Status" &&
																	e.Parent!.Elements( "HistoryItem" )
																			.Any( eh => eh.hisType() == "EmpHist" && eh.Element( "date-end" ) != null )
													);
							statusToRemove.Remove();
						}

						xDataDef.WriteTo( xw );
						break;
				}
			}

			xw.WriteEndElement();
		}
		File.Delete( outputFile );
		File.Move( resultFile, outputFile );
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
