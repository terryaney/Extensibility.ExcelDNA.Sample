using System.Text;
using System.Xml;
using System.Xml.Linq;
using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.DataExport;

class Provider
{
	public static string[]? Export( ExcelReference selection, ExportDataInfo info )
	{
		var tablesToExport = TablesToExport( selection );
		var validationErrors = Validate( tablesToExport, true );

		if ( validationErrors.Any() )
		{
			ExcelDna.Logging.LogDisplay.Clear();
			ExcelDna.Logging.LogDisplay.WriteLine( "*****Data Exporting Issues*****" );

			foreach ( var error in validationErrors )
			{
				ExcelDna.Logging.LogDisplay.WriteLine( "\t" + error );
			}

			ExcelDna.Logging.LogDisplay.Show();

			return null;
		}

		if ( info.Action == ExportDataAction.Validate )
		{
			MessageBox.Show( "Validation successful.  No issues found.", "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information );
			return null;
		}

		var configuration = GetConfiguration( tablesToExport );
		var authIds = tablesToExport.SelectMany( t => t.AuthIds ).Distinct().OrderBy( a => a );

		var exportedFiles = ProcessChunks( tablesToExport, authIds.GetEnumerator(), info.OutputFile!, info.MaxFileSize, configuration );

		return exportedFiles.ToArray();
	}

	private static ExportTable[] TablesToExport( ExcelReference topLeftRef )
	{
		var topRightRef = topLeftRef.End( DirectionType.ToRight );
		var configValues = topLeftRef.Extend( topRightRef ).GetValues<string>();

		if ( configValues.Any( v => string.IsNullOrEmpty( v ) ) )
		{
			throw new ArgumentOutOfRangeException( $"There appears to be hidden columns within {topLeftRef.GetAddress()} - {topRightRef.GetAddress()}.  Please ensure there are no hidden columns and no blank column headers." );
		}

		var isMultiSheet = IsMultiSheetConfiguration( topLeftRef );

		var multiSheetInfo = isMultiSheet
			? topLeftRef.Extend( topLeftRef.End( DirectionType.ToRight ).End( DirectionType.Down ) ).GetArray<string>()
			: null;

		var exportTables = !isMultiSheet
			? new[] {
				new ExportTable(
					topLeftRef,
					// TCA: IsRowExport - see if there is a Range.Find I can use
					configValues.Any( v => v!.Contains( "/Table:", StringComparison.InvariantCultureIgnoreCase ) ) ? ExportType.RowBased : ExportType.ColumnBased
				)
			}
			: Enumerable
				.Range( 1, multiSheetInfo!.GetUpperBound( 0 ) )
				.Select( r => new { Row = r, Reference = multiSheetInfo[ r, 1 ]!.GetReference( multiSheetInfo[ r, 0 ]! ) } )
				.Select( r =>
					new ExportTable
					(
						r.Reference,
						bool.Parse( multiSheetInfo[ r.Row, 2 ]!.ToLower() ) ? ExportType.RowBased : ExportType.ColumnBased
					)
				).ToArray();

		return exportTables;
	}

	private static bool IsMultiSheetConfiguration( ExcelReference selection )
	{
		var isMultiSheet = (string)selection.GetValue() == "Sheet" &&
			(string)selection.Offset( 0, 1 ).GetValue() == "Address" &&
			(string)selection.Offset( 0, 2 ).GetValue() == "Row Format";

		return isMultiSheet;
	}

	private static IEnumerable<string> Validate( ExportTable[] tables, bool validateDuplicateColumns )
	{
		foreach ( var error in tables.SelectMany( t => t.Validate() ) )
		{
			yield return error;
		}

		/*
		- Validate duplicate fields (across flat and each history type)
		- Validate authid col names on col based exports match
		*/
		var flatFields =
			tables.Where( t => t.ExportType != ExportType.RowBased )
				  .SelectMany( t =>
						t.ExportableColumns
							.Where( c => string.IsNullOrEmpty( c.Field.HistoryTable ) )
							.Select( c =>
								new
								{
									t.TableKey,
									c.IsAuthId,
									c.Field.Field,
									c.Address
								}
							)
				  );

		var duplicates =
			flatFields.Where( f => !f.IsAuthId )
					  .GroupBy( f => f.Field )
					  .Where( g => g.Count() > 1 );

		if ( validateDuplicateColumns )
		{
			foreach ( var g in duplicates )
			{
				var original = g.First().Address;
				yield return string.Format( "DUPLICATE FIELD: Duplicate {0} field detected.  Originally at {1}.  Duplicate column(s) specified: {2}",
					g.Key,
					original,
					string.Join( ", ", g.Skip( 1 ).Select( d => d.Address ) )
				);
			}
		}

		duplicates =
			flatFields.Where( f => f.IsAuthId )
					  .GroupBy( f => f.Field )
					  .ToArray();

		if ( duplicates.Count() > 1 )
		{
			yield return string.Format( "CONFLICTING AuthID columns: '{0}' specified at {1}.",
				string.Join( ", ", duplicates.Select( d => d.Key ).Distinct() ),
				string.Join( ", ", duplicates.SelectMany( d => d.Select( i => i.Address ) ) )
			);
		}

		if ( validateDuplicateColumns )
		{
			// Don't allow same history table to be configured on multiple row based sheets, would lead towards index overlapping
			// which I don't validate against and would probably cause maintenance issue as person who exported would be surprised why
			// data isn't what they expect it to be (last sheet processed would be applied)

			var historyOnMultipleTables =
				tables.Where( t => t.ExportType == ExportType.RowBased )
					  .GroupBy( g => g.RowExportHistoryTable )
					  .Where( g => g.Count() > 1 )
					  .Select( g => new { g.Key, Addresses = g.Select( i => i.TableKey ) } );

			foreach ( var d in historyOnMultipleTables )
			{
				yield return string.Format( "MULTIPLE Row Based History Configurations: {0} history table export configured on multiple row based data export tables.  {0} history table configurations found at {1}.",
					d.Key,
					string.Join( ", ", d.Addresses )
				);
			}

			var historyFields =
				tables.SelectMany( t => t.ExportType != ExportType.RowBased
					? t.ExportableColumns
						.Where( c => !string.IsNullOrEmpty( c.Field.HistoryTable ) )
						.Select( c => new { c.Field.HistoryTable, c.Field.Field, c.Field.Index, c.Address, t.TableKey } )
					: t.ExportableColumns
						.Select( c => new { HistoryTable = t.RowExportHistoryTable, c.Field.Field, Index = (string?)t.TableKey, c.Address, t.TableKey } )
				);


			// Only when indexes are same (column based could have [Status:2015]status, [Status:2016]status, etc.)
			var duplicateHistory =
				historyFields
					.GroupBy( g => new { g.HistoryTable, g.Index, g.Field } )
					.Where( g => g.Count() > 1 );

			foreach ( var d in duplicateHistory )
			{
				var original = d.First().Address;
				var name = string.Format( "{0}.{1}", d.Key.HistoryTable, d.Key.Field );

				yield return string.Format( "DUPLICATE FIELD: Duplicate {0} field detected.  Originally at {1}.  Duplicate column(s) specified at {2}.",
					name,
					original,
					string.Join( ", ", d.Skip( 1 ).Select( f => f.Address ) )
				);
			}
		}
	}

	private static XElement GetConfiguration( ExportTable[] tables )
	{
		var configuration = new XElement( "Configuration" );

		var flatFields =
			tables.Where( t => t.ExportType != ExportType.RowBased )
				.SelectMany( t =>
					t.ExportableColumns
						.Where( c => string.IsNullOrEmpty( c.Field.HistoryTable ) )
						.Select( c => c.Field.Field )
				)
				.Distinct()
				.OrderBy( f => f );

		var container = new XElement( "Profile" );
		configuration.Add( container );

		foreach ( var f in flatFields.Where( f => !new[] { "profileDateCreated", "profileDateUpdated" }.Contains( f ) ) )
		{
			container.Add( new XElement( f! ) );
		}

		var historyTables =
			tables.SelectMany( t => t.ExportType != ExportType.RowBased
				? t.ExportableColumns
					.Where( c => !string.IsNullOrEmpty( c.Field.HistoryTable ) )
					.Select( c => new { c.Field.HistoryTable, c.Field.Field, c.Field.Index, c.Address, t.TableKey } )
				: t.ExportableColumns
					.Select( c => new { HistoryTable = t.RowExportHistoryTable, c.Field.Field, Index = (string?)t.TableKey, c.Address, t.TableKey } )
			)
			.GroupBy( g => g.HistoryTable )
			.OrderBy( g => g.Key );

		foreach ( var table in historyTables )
		{
			container = new XElement( "HistoryData", new XAttribute( "type", table.Key ?? "Not available" ), new XElement( "index" ) );
			configuration.Add( container );

			var fieldNames =
				table.Select( f => f.Field )
					.Where( f => !new[] { "hisDateCreated", "hisDateUpdated", "index" }.Contains( f ) )
					.Distinct()
					.OrderBy( f => f );

			foreach ( var fieldName in fieldNames )
			{
				container.Add( new XElement( fieldName! ) );
			}
		}

		return configuration;
	}

	private static IEnumerable<string> ProcessChunks(
		ExportTable[] exportTables,
		IEnumerator<string> enumerator,
		string saveLocation,
		int? maxFileSize,
		XElement configuration,
		int iteration = 1 )
	{
		var settings = new XmlWriterSettings
		{
			Encoding = Encoding.UTF8,
			// OmitXmlDeclaration = true,
			Indent = System.Diagnostics.Debugger.IsAttached
		};

		var totalRows = 0;

		var exportFileName = ( maxFileSize ?? 0 ) > 0
			? Path.Combine(
				Path.GetDirectoryName( saveLocation )!,
				string.Format( "{0}.{1}{2}.NoConfig",
					Path.GetFileNameWithoutExtension( saveLocation ),
					iteration,
					Path.GetExtension( saveLocation )
				)
			  )
			: Path.Combine(
				Path.GetDirectoryName( saveLocation )!,
				string.Format( "{0}.NoConfig",
					Path.GetFileName( saveLocation )
				)
			  );

		using ( var fs = new FileStream( exportFileName, FileMode.Create ) )
		using ( var writer = XmlWriter.Create( fs, settings ) )
		{
			writer.WriteStartElement( "xDataDefs" );
			configuration.WriteTo( writer );

			while ( enumerator.MoveNext() )
			{
				var authId = enumerator.Current;
				try
				{
					var values = exportTables.SelectMany( t => t.GetValues( authId ) );
					var xDataDef = ToXElement( values, authId );

					if ( xDataDef != null )
					{
						totalRows++;
						xDataDef.WriteTo( writer );
						writer.Flush();

						if ( maxFileSize > 0 && fs.Length >= maxFileSize )
						{
							break;
						}
					}
				}
				catch ( ExcelErrorException ex )
				{
					throw new ApplicationException(
						$"Unable to export values for {authId}.",
						ex
					);

				}
			}

			writer.WriteEndElement();
		}

		if ( totalRows > 0 )
		{
			var fileReturn = InjectxDSCount( exportFileName, totalRows, settings );

			yield return fileReturn;

			foreach ( var l in ProcessChunks( exportTables, enumerator, saveLocation, maxFileSize, configuration, ++iteration ) )
			{
				yield return l;
			}
		}

		if ( File.Exists( exportFileName ) )
		{
			File.Delete( exportFileName );
		}
	}

	private static XElement? ToXElement( IEnumerable<ExportValue> values, string authId )
	{
		var xDataDef =
			new XElement( "xDataDef",
				new XAttribute( "id-auth", authId ),
				new XElement( "Profile" ),
				new XElement( "HistoryData" )
			);
		var profile = xDataDef.Element( "Profile" )!;
		var historyData = xDataDef.Element( "HistoryData" )!;
		var isProfileDelete = false;

		var historySkips = new Dictionary<string, XElement?>();
		var historyDeletes = new Dictionary<string, XElement>();
		var historyModels = new Dictionary<string, XElement>();
		var notes = new List<XElement>();

		string? lastDataType = null;
		string? lastFieldName = null;
		object? lastValue = null;

		try
		{
			foreach ( var v in values )
			{
				lastDataType = null;
				lastFieldName = null;
				lastValue = null;

				if ( v.SkipExport && v.HistoryTable == null )
				{
					return null; // Skip exporting the profile
				}
				else if ( v.SkipExport && v.Field == null )
				{
					lastDataType = "History:SkipExport";
					lastFieldName = "Index";
					lastValue = v.Index;

					var key = v.HistoryTable + v.Index;

					// Skip exporting history row
					if ( !historySkips.ContainsKey( key ) )
					{
						historySkips.Add( key, null );
						historyData.Elements( "HistoryItem" ).Where( h => h.hisType() == v.HistoryTable && h.hisIndex() == v.Index ).Remove();
					}
				}
				else if ( v.Clear && v.Field == null ) // delete profile or history table or history row
				{
					if ( v.HistoryTable == null )
					{
						lastDataType = "Profile:Delete";

						if ( !isProfileDelete ) // if not already added
						{
							profile.Parent!.Add( new XAttribute( "delete", 1 ) );
							profile.Elements().Remove();
							isProfileDelete = true;
						}
					}
					else if ( v.Index == null )
					{
						lastDataType = $"History:{v.HistoryTable}:Delete";

						if ( !historyDeletes.ContainsKey( v.HistoryTable ) )
						{
							var clear = new XElement( "HistoryItem",
												new XAttribute( "hisType", v.HistoryTable ),
												new XAttribute( "hisClear", "Table" ) );

							historyDeletes.Add( v.HistoryTable, clear );

							historyData.Elements( "HistoryItem" ).Where( h => h.hisType() == v.HistoryTable && h.Attribute( "hisClear" ) != null ).Remove();
							historyData.AddFirst( clear );
						}
					}
					else
					{
						lastDataType = $"History:{v.HistoryTable}:Delete";
						lastFieldName = "Index";
						lastValue = v.Index;

						if ( !historyDeletes.ContainsKey( v.HistoryTable ) && !historyDeletes.ContainsKey( v.HistoryTable + v.Index ) )
						{
							var clear = new XElement( "HistoryItem",
												new XAttribute( "hisType", v.HistoryTable ),
												new XAttribute( "hisIndex", v.Index ),
												new XAttribute( "hisClear", "Row" ) );
							historyDeletes.Add( v.HistoryTable + v.Index, clear );
							historyData.AddFirst( clear );
						}
					}
				}
				else if ( !isProfileDelete ) // If not flagged to delete this profile
				{
					XElement getFieldElement( ExportValue xds ) =>
						new( ( lastFieldName = xds.Field )!,
							!xds.AllowReplace ? new XAttribute( "noreplace", 1 ) : null,
							xds.Clear ? new XAttribute( "delete", 1 ) : null,
							!xds.Clear ? lastValue = xds.Value : null
						);

					if ( !string.IsNullOrEmpty( v.Subject ) )
					{
						lastDataType = "Note";
						lastFieldName = "Subject/Body";
						lastValue = $"Subject: {v.Subject}, Body: {v.Body}";

						notes.Add( new XElement( "FolderItem", new XElement( "note", new XElement( "subject", v.Subject ), new XElement( "body", v.Body ) ) ) );
					}
					else if ( v.HistoryTable == null )
					{
						lastDataType = "Profile" + ( v.Clear ? ":DeleteField" : "" );

						if ( v.Field == "profileDateUpdated" )
						{
							profile.Add( new XAttribute( v.Field, XmlConvert.ToString( DateTime.Parse( v.Value! ), XmlDateTimeSerializationMode.Unspecified ) ) );
						}
						else if ( v.Field == "profileDateCreated" )
						{
							profile.Add( new XAttribute( v.Field, XmlConvert.ToString( DateTime.Parse( v.Value! ), XmlDateTimeSerializationMode.Unspecified ) ) );
						}
						else
						{
							var fieldExists = profile.Element( v.Field! ) != null;

							if ( fieldExists && !v.IsAuthID /* multiple column based sheets will each have authid */ )
							{
								throw new IndexOutOfRangeException( string.Format( "You have provided Profile.{0} multiple times for {1}.  Please fix your export configuration to only include it once.", v.Field, authId ) );
							}

							if ( !fieldExists )
							{
								profile.Add( getFieldElement( v ) );
							}
						}
					}
					else
					{
						lastDataType = $"History:{v.HistoryTable}:{v.Index}";
						lastFieldName = "Index";
						lastValue = v.Index;

						var key = v.HistoryTable + v.Index;

						if ( !historySkips.ContainsKey( key ) )
						{
							var modelExists = historyModels.ContainsKey( key );
							var historyItem = modelExists
								? historyModels[ key ]
								: new XElement( "HistoryItem",
									new XAttribute( "hisType", v.HistoryTable ),
									new XAttribute( "hisIndex", v.Index! ),
									!string.IsNullOrEmpty( v.NewIndex ) ? new XAttribute( "new-index", v.NewIndex ) : null,
									new XElement( "index", v.Index ) );

							if ( !modelExists )
							{
								historyModels[ key ] = historyItem;
								historyData.Add( historyItem );
							}

							// If 'index' is returned, then a row based history table was dumped with only index column
							// being populated
							if ( !( v.Field == "index" && !string.IsNullOrEmpty( v.HistoryTable ) ) )
							{
								if ( historyItem.Element( v.Field! ) != null )
								{
									throw new IndexOutOfRangeException( string.Format( "You have provided [{0}:{1}]{2} multiple times for {3}.  Please fix your export configuration to only include it once.", v.HistoryTable, v.Index, v.Field, authId ) );
								}
								if ( v.Field == "hisDateUpdated" )
								{
									historyItem.Add( new XAttribute( v.Field, XmlConvert.ToString( DateTime.Parse( v.Value! ), XmlDateTimeSerializationMode.Unspecified ) ) );
								}
								else if ( v.Field == "hisDateCreated" )
								{
									historyItem.Add( new XAttribute( v.Field, XmlConvert.ToString( DateTime.Parse( v.Value! ), XmlDateTimeSerializationMode.Unspecified ) ) );
								}
								else
								{
									historyItem.Add( getFieldElement( v ) );
								}
							}
						}
					}
				}
			}
		}
		catch ( ExcelErrorException )
		{
			throw;
		}
		catch ( IndexOutOfRangeException )
		{
			throw;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException(
				$"Unable to generate xDS element for {authId}.  Processing {lastDataType}, field: {lastFieldName}, value: {lastValue}",
				ex
			);
		}

		if ( !isProfileDelete && !profile.HasElements )
		{
			profile.Remove();
		}
		if ( isProfileDelete || !historyData.HasElements )
		{
			historyData.Remove();
		}
		if ( !isProfileDelete && notes.Any() )
		{
			xDataDef.Add( notes );
		}

		return xDataDef.HasElements ? xDataDef : null;
	}

	private static string InjectxDSCount( string noConfigFile, int totalRows, XmlWriterSettings settings )
	{
		var configFile = Path.Combine( Path.GetDirectoryName( noConfigFile )!, Path.GetFileNameWithoutExtension( noConfigFile ) );

		using ( var xw = XmlWriter.Create( configFile, settings ) )
		using ( var reader = new XmlTextReader( noConfigFile ) )
		{
			while ( reader.Read() )
			{
				reader.StreamXmlNode( xw );

				if ( reader.NodeType == XmlNodeType.Element && reader.IsStartElement( "xDataDefs" ) )
				{
					xw.WriteAttributeString( "TotalRows", totalRows.ToString() );
				}
			}
		}

		return configFile;
	}
}