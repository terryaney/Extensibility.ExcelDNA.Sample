using System.Text.RegularExpressions;
using ExcelDna.Integration;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.DataExport;

class ExportTable
{
	public ExportType ExportType { get; private set; }
	public string TableKey { get; private set; }
	public string? RowExportHistoryTable { get; private set; }

	private readonly ExportColumn[] headers;
	private readonly ExcelReference excelReference;
	private readonly int authIdCol;
	private readonly int rowExportHistoryIndexCol;
	private readonly int rowExportNewHistoryIndexCol;
	private readonly bool clearRowExportHistoryBeforeLoad;
	private readonly Dictionary<string, ControlColumn> exportColumns;
	private readonly Dictionary<string, ControlColumn> deleteColumns;
	private InteropArray data = null!;
	private ILookup<string, int> profileRows = null!;

	private static readonly Regex noteRegex = new ( @"^/subject:(?<subject>.+)/body:(?<body>.+)$", RegexOptions.Compiled );

	public ExportTable( ExcelReference topLeft, ExportType exportType )
	{
		excelReference = topLeft;
		ExportType = exportType;
		TableKey = topLeft.GetAddress();

		var headerInfo = new ExcelReference(
			topLeft.RowFirst,
			topLeft.RowFirst,
			topLeft.ColumnFirst,
			topLeft.End( DirectionType.ToRight ).ColumnFirst,
			topLeft.SheetId
		).GetArray<string>();

		var indexReplacements = new Dictionary<string, string>();

		string indexReplace( string table, string index )
		{
			var key = string.Format( "{0}.{1}", table, index );

			if ( indexReplacements.TryGetValue( key, out var v ) )
			{
				return v;
			}

			var value = InputBox.Show(
				string.Format( "Please provide a replacement index for {0}:", key ),
				string.Format( "Dynamic {0} Index", table ),
				index[ 1..^1 ]
			);

			if ( value.ReturnCode != DialogResult.OK )
			{
				throw new ArgumentNullException( string.Format( "You must provide a replacement index for {0} table, index {1}.", table, index ) );
			}

			return indexReplacements[ key ] = value.Text;
		}

		var isRowBased = exportType == ExportType.RowBased;

		var parsedHeaders =
			Enumerable.Range( 0, headerInfo.Length )
				.Select( c => new { Label = headerInfo[ 0, c ], Info = headerInfo[ 0, c ]!.Split( '/' ) } )
				.Select( c => new
				{
					Field = ExportField.Parse( c.Info[ 0 ], !isRowBased ? (Func<string, string, string>)indexReplace : null ),
					c.Label,
					SwitchInfo = c.Info.Skip( 1 ).Select( i => i.Split( ':' ) )
				} )
				.Select( c => new
				{
					c.Field,
					c.Label,
					IsSsn = string.Compare( c.Field.Field, "ssn", true ) == 0,
					Switches = c.SwitchInfo.Select( i => new { Key = i[ 0 ], Value = i.Length > 1 ? i[ 1 ] : i[ 0 ] } )
				} )
				.ToArray();

		var hasAuthIdSwitch = parsedHeaders.Any( h => h.Switches.Any( s => string.Compare( s.Key, "key", true ) == 0 ) );
		var hasHistoryIndexSwitch = isRowBased && parsedHeaders.Any( h => h.Switches.Any( s => string.Compare( s.Key, "index", true ) == 0 ) );

		headers = parsedHeaders.Select( ( h, i ) =>
			new ExportColumn
			{
				Address = topLeft.Offset( 0, i ).GetAddress(),

				Field = h.Field,
				Label = h.Label,
				Ordinal = i,
				IgnoreColumn = h.Switches.Any( s => string.Compare( s.Key, "off", true ) == 0 || string.Compare( s.Key, "new-index", true ) == 0 ),

				IsAuthId = hasAuthIdSwitch
					? h.Switches.Any( s => string.Compare( s.Key, "key", true ) == 0 )
					: h.IsSsn,

				IsText = h.Switches.Any( s => string.Compare( s.Key, "text", true ) == 0 ),
				IsDate = h.Switches.Any( s => string.Compare( s.Key, "date", true ) == 0 ),
				IsDateTime = h.Switches.Any( s => string.Compare( s.Key, "dateTime", true ) == 0 ),
				DeleteIfBlank = h.Switches.Any( s => string.Compare( s.Key, "deleteIfBlank", true ) == 0 ),
				AllowReplace = !h.Switches.Any( s => string.Compare( s.Key, "noReplace", true ) == 0 ),
				DateConvertFormat = h.Switches.Where( s => string.Compare( s.Key, "dateConvert", true ) == 0 ).Select( s => s.Value ).FirstOrDefault(),
				Format = h.Switches.Where( s => string.Compare( s.Key, "format", true ) == 0 ).Select( s => s.Value ).FirstOrDefault() ?? ( h.IsSsn ? "000000000" : null ),
				DefaultValue = h.Switches.Where( s => string.Compare( s.Key, "default", true ) == 0 ).Select( s => s.Value ).FirstOrDefault(),
				ToUpper = h.Switches.Any( s => string.Compare( s.Key, "upper", true ) == 0 ),
				ToLower = h.Switches.Any( s => string.Compare( s.Key, "lower", true ) == 0 ),
				IgnoreZero = h.Switches.Any( s => string.Compare( s.Key, "ignoreZero", true ) == 0 ),
				DecimalPlacesToInsert = h.Switches.Where( s => string.Compare( s.Key, "numberConvert", true ) == 0 ).Select( s => (int?)int.Parse( s.Value[ 1.. ] ) ).FirstOrDefault(),

				IsExportControl = h.Switches.Any( s => string.Compare( s.Key, "export", true ) == 0 ),
				IsDeleteControl = h.Switches.Any( s => string.Compare( s.Key, "delete", true ) == 0 ),
				IsProfileNote = h.Field.Field?.StartsWith( "ProfileNotes", StringComparison.InvariantCultureIgnoreCase ) ?? false,

				RowExportHistoryTable = isRowBased ? h.Switches.Where( s => string.Compare( s.Key, "table", true ) == 0 ).Select( s => s.Value ).FirstOrDefault() : null,
				ClearRowExportHistoryBeforeLoad = isRowBased && h.Switches.Any( s => string.Compare( s.Key, "clearAll", true ) == 0 ),
				IsRowExportHistoryIndex = isRowBased &&
				(
					( !hasHistoryIndexSwitch && string.Compare( h.Field.Field, "index", true ) == 0 && !h.Switches.Any( s => string.Compare( s.Key, "off", true ) == 0 ) ) ||
					( hasHistoryIndexSwitch && h.Switches.Any( s => string.Compare( s.Key, "index", true ) == 0 ) && !h.Switches.Any( s => string.Compare( s.Key, "off", true ) == 0 ) )
				),
				IsRowExportHistoryNewIndex = isRowBased && h.Switches.Any( s => string.Compare( s.Key, "new-index", true ) == 0 ) && !h.Switches.Any( s => string.Compare( s.Key, "off", true ) == 0 )
			} ).ToArray();

		authIdCol =
			headers.Where( h => h.IsAuthId )
				.Select( h => (int?)h.Ordinal )
				.FirstOrDefault() ?? -1;

		var duplicates =
			headers.Where( h => h.IsExportControl ).GroupBy( g => g.Field.ToString() ).Where( s => s.Count() > 1 ).ToArray();

		if ( duplicates.Any() )
		{
			throw new IndexOutOfRangeException( $"You have provided duplicate export columns for {TableKey}." );
		}

		exportColumns =
			headers
				.Where( h => h.IsExportControl )
				.ToDictionary( h => h.Field.ToString(), h => new ControlColumn { Field = h.Field, Ordinal = h.Ordinal } );

		duplicates =
			headers.Where( h => h.IsDeleteControl ).GroupBy( g => g.Field.ToString() ).Where( s => s.Count() > 1 ).ToArray();

		if ( duplicates.Any() )
		{
			throw new IndexOutOfRangeException( $"You have provided duplicate delete columns for {TableKey}." );
		}

		deleteColumns =
			headers.Where( h => h.IsDeleteControl )
				.ToDictionary( h => h.Field.ToString(), h => new ControlColumn { Field = h.Field, Ordinal = h.Ordinal } );

		if ( isRowBased )
		{
			var rowBasedHeaderInfo = headers.FirstOrDefault( h => !string.IsNullOrEmpty( h.RowExportHistoryTable ) );

			if ( rowBasedHeaderInfo != null )
			{
				RowExportHistoryTable = rowBasedHeaderInfo.RowExportHistoryTable;
				clearRowExportHistoryBeforeLoad = rowBasedHeaderInfo.ClearRowExportHistoryBeforeLoad;
				rowExportHistoryIndexCol = headers.FirstOrDefault( h => h.IsRowExportHistoryIndex )?.Ordinal ?? -1;
				rowExportNewHistoryIndexCol = headers.FirstOrDefault( h => h.IsRowExportHistoryNewIndex )?.Ordinal ?? -1;
			}
		}
	}

	public IEnumerable<string> Validate()
	{
		var columnsToTest = headers.Where( h => !h.IgnoreColumn && !h.IsExportControl && !h.IsDeleteControl );

		var validationCols = columnsToTest.Where( c => c.IsAuthId ).ToArray();

		if ( !validationCols.Any() )
		{
			yield return string.Format( "MISSING AuthID: Export table starting at {0} does not have an AuthID specified.  Please provide a ssn column or an AuthID column via the /key switch.", TableKey );
		}
		else if ( validationCols.Length > 1 )
		{
			yield return string.Format( "MULTIPLE AuthID COLUMN's: Export table starting at {0} provided more than one AuthID.  The columns specified were {1}.", TableKey, string.Join( ", ", validationCols.Select( c => c.Address ) ) );
		}
		else if ( ExportType == ExportType.ColumnBased )
		{
			var bottomLeft = excelReference.End( DirectionType.Down );
			var exportHeader = headers.FirstOrDefault( h => h.IsExportControl );

			string?[] duplicates;

			var dataRange = new ExcelReference(
				excelReference.RowFirst + 1,
				bottomLeft.RowFirst,
				excelReference.ColumnFirst + authIdCol,
				excelReference.ColumnFirst + authIdCol,
				excelReference.SheetId
			);

			var authIds = dataRange.GetValues<string>();

			if ( exportHeader == null )
			{
				duplicates =
					authIds
						.GroupBy( a => a )
						.Where( g => g.Count() > 1 )
						.Select( g => g.Key )
						.ToArray();
			}
			else
			{
				// Weed out /export=0 people
				dataRange = new ExcelReference( excelReference.RowFirst + 1, bottomLeft.RowFirst, excelReference.ColumnFirst + exportHeader.Ordinal, excelReference.ColumnFirst + exportHeader.Ordinal, excelReference.SheetId );
				var exportAllowed = dataRange.GetValues<string>();

				duplicates =
					authIds
						.Where( ( a, i ) => exportAllowed[ i ] != "0" )
						.GroupBy( a => a )
						.Where( g => g.Count() > 1 )
						.Select( g => g.Key )
						.ToArray();
			}

			if ( duplicates.Any() )
			{
				yield return $"DUPLICATE AuthID's: Export table starting at {TableKey} contains duplicate AuthID columns.  The duplicates specified were {string.Join( ", ", duplicates )}.";
			}
		}

		if ( ExportType == ExportType.RowBased )
		{
			validationCols = columnsToTest.Where( c => !string.IsNullOrEmpty( c.RowExportHistoryTable ) ).ToArray();

			if ( !validationCols.Any() )
			{
				yield return string.Format( "MISSING /table: Row based Export table starting at {0} does not have a history table specified.  Please provide a table name via the /table:tableName switch.", TableKey );
			}
			else if ( validationCols.Length > 1 )
			{
				yield return string.Format( "MULTIPLE /table's: Row based Export table starting at {0} provided more than one history table columns.  The columns specified were {1}.", TableKey, string.Join( ", ", validationCols.Select( c => c.Address ) ) );
			}

			validationCols = columnsToTest.Where( c => c.IsRowExportHistoryIndex ).ToArray();

			if ( false /* will do UniqueN if missing */ && !validationCols.Any() )
			{
				yield return string.Format( "MISSING /index: Row based Export table starting at {0} does not have an history index column specified.  Please provide an index column; either named 'index' or via the /index switch.", TableKey );
			}
			else if ( validationCols.Length > 1 )
			{
				yield return string.Format( "MULTIPLE /index's: Row based Export table starting at {0} provided more than one history index columns.  The columns specified were {1}.", TableKey, string.Join( ", ", validationCols.Select( c => c.Address ) ) );
			}

			validationCols = columnsToTest.Where( c => c.ClearRowExportHistoryBeforeLoad && string.IsNullOrEmpty( c.RowExportHistoryTable ) ).ToArray();

			if ( validationCols.Any() )
			{
				yield return string.Format( "INVALID /clearAll: Row based Export table starting at {0} provided a /clearAll switch on an invalid column.  You can only supply this switch on the /table:tableName column.  The column(s) specified were {1}.", TableKey, string.Join( ", ", validationCols.Select( c => c.Address ) ) );
			}

			validationCols = columnsToTest.Where( c => c.DeleteIfBlank && c.IsRowExportHistoryIndex ).ToArray();

			if ( validationCols.Any() )
			{
				yield return string.Format( "INVALID /deleteIfBlank: Row based Export table starting at {0} provided a /deleteIfBlank switch on an 'index' column.  You cannot supply this switch on a history table's index column.  The column(s) specified were {1}.", TableKey, string.Join( ", ", validationCols.Select( c => c.Address ) ) );
			}

			foreach ( var f in ExportableColumns.Where( c => !string.IsNullOrEmpty( c.Format ) ) )
			{
				var validFormat = true;
				try
				{
					var testFormat = ( 0d ).ToString( f.Format );
				}
				catch
				{
					validFormat = false;
				}

				if ( !validFormat )
				{
					yield return string.Format( "INVALID /format: Invalid format provided for {0}.  Format specified was '{1}'.", f.Address, f.Format );
				}
			}
		}

		yield break;
	}

	public ExportColumn[] ExportableColumns =>
		ExportableColumnsAndExportDelete
			.Where( h => !h.IsExportControl && !h.IsDeleteControl )
			.ToArray();

	private ExportColumn[] ExportableColumnsAndExportDelete
	{
		get
		{
			var isRowBased = ExportType == ExportType.RowBased;
			return
				headers
					.Where( h => !h.IgnoreColumn && !h.IsProfileNote )
					// If row based, skip authid column
					.Where( h => !( isRowBased && h.IsAuthId ) )

					// No longer skipping b/c if a history substitution index column, i.e. [Table:1]index, is only thing provided, it wasn't exporting

					// Skip 'index' if not on profile
					// .Where( h => !( h.Field.Field == "index" && ( h.Field.HistoryTable != null || isRowBased ) ) )
					.ToArray();
		}
	}

	private void EnsureDataRead()
	{
		if ( data == null )
		{
			var bottomLeft = excelReference.End( DirectionType.Down );
			var emptyData = bottomLeft == excelReference;

			var topRight = excelReference.End( DirectionType.ToRight );

			var dataRange = new ExcelReference(
				excelReference.RowFirst + 1,
				emptyData ? excelReference.RowFirst + 1 : bottomLeft.RowFirst,
				excelReference.ColumnFirst,
				topRight.ColumnFirst,
				excelReference.SheetId
			);

			data = dataRange.GetValueArray();

			if ( !emptyData && authIdCol > -1 )
			{
				profileRows =
					Enumerable.Range( 0, data.RowCount )
						.Where( r => !string.IsNullOrEmpty( data[ r, authIdCol ]?.ToString() ) )
						.Select( r =>
							new
							{
								AuthId = GetExportValue( data[ r, authIdCol ], headers[ authIdCol ] )!,
								Row = r
							}
						).ToLookup( k => k.AuthId, e => e.Row );
			}
			else if ( emptyData )
			{
				profileRows = Enumerable.Empty<int>().ToLookup( x => "Empty" );
			}
		}
	}

	public IEnumerable<string> AuthIds
	{
		get
		{
			EnsureDataRead();

			foreach ( var lookup in profileRows )
			{
				yield return lookup.Key;
			}
		}
	}

	public IEnumerable<ExportValue> GetValues( string authId )
	{
		EnsureDataRead();

		if ( profileRows.Contains( authId ) )
		{
			var isRowBased = ExportType == ExportType.RowBased;
			var clearInstructionIssued = false;

			foreach ( var row in profileRows[ authId ] )
			{
				var rowBasedIndex = isRowBased
					? rowExportHistoryIndexCol == -1
						? $"{{Unique{row + 1}}}"
						: GetExportValue( data[ row, rowExportHistoryIndexCol ], headers[ rowExportHistoryIndexCol ] )
					: null;

				var rowBasedNewIndex = isRowBased && rowExportNewHistoryIndexCol != -1
					? GetExportValue( data[ row, rowExportNewHistoryIndexCol ], headers[ rowExportNewHistoryIndexCol ] )
					: null;

				if ( clearRowExportHistoryBeforeLoad && !clearInstructionIssued )
				{
					yield return new ExportValue { HistoryTable = RowExportHistoryTable, Clear = true };
					clearInstructionIssued = true;
				}

				// Row export for Profile or history table has Field == null, meaning this is just /export or /delete
				// with no field name before the switch.
				var skipExportingRow = false;
				foreach ( var exportColumn in exportColumns.Where( i => i.Value.Field.Field == null && i.Value.Field.Index == null /* If index != null, then it is like [Status:2005]/export on a col based export */ ) )
				{
					var field = exportColumn.Value.Field;
					var export = data[ row, exportColumn.Value.Ordinal ];

					if ( export != null && (double)export == 0 )
					{
						var isProfileExport = !isRowBased && ( field.HistoryTable ?? "Profile" ) == "Profile"; // HistoryTable is null when doing column based and using /Export
																											   // When *just* /export is provided, if History (row based), then I can simply skip exporting anything from this row
						skipExportingRow = !isProfileExport && isRowBased;

						// If on profile and export=0, can just exit out of function
						if ( isProfileExport )
						{
							yield return new ExportValue
							{
								/*
								 *  Update: This should always only be profile now, so don't see need to assign type/index
								 *
								// When *just* /export is provided, if column based export, would be applied to Profile, otherwise it is row based export and applied to
								// History, and would think RowExportHistoryTable is always set, but leaving code as is.
								HistoryTable = isProfileExport ? null : ( RowExportHistoryTable ?? field.HistoryTable ),
								Index = rowBasedIndex ?? field.Index,
								*/
								SkipExport = true
							};

							yield break;
						}
					}
				}

				if ( !skipExportingRow )
				{
					var hadHistoryRowBasedDeleteColumn = false;

					foreach ( var deleteColumn in deleteColumns.Where( i => i.Value.Field?.Field == null ) )
					{
						var field = deleteColumn.Value.Field;
						var delete = data[ row, deleteColumn.Value.Ordinal ];

						if ( delete != null && (double)delete == 1 && !deleteColumn.Value.Field.DynamicIndex )
						{
							var isProfileDelete = field.HistoryTable == "Profile";
							hadHistoryRowBasedDeleteColumn = !string.IsNullOrEmpty( RowExportHistoryTable ); // 'profile' will yield break anyway

							yield return new ExportValue
							{
								HistoryTable = isProfileDelete ? null : ( RowExportHistoryTable ?? field.HistoryTable ),
								Index = rowBasedIndex ?? field.Index,
								Clear = true
							};

							if ( isProfileDelete )
							{
								yield break;
							}
						}
					}

					foreach ( var header in headers.Where( h => h.IsProfileNote ) )
					{
						var subject = string.Format( "Default Note Exported on {0:MM/dd/yyyy hh:mm tt}", DateTime.Now );
						var body = GetExportValue( data[ row, header.Ordinal ], header );

						if ( !string.IsNullOrEmpty( body ) )
						{
							if ( noteRegex.IsMatch( body ) )
							{
								var matches = noteRegex.Match( body ).Groups;
								subject = matches[ "subject" ].Value;
								body = matches[ "body" ].Value;
							}

							yield return new ExportValue
							{
								Subject = subject,
								Body = body
							};
						}
					}

					bool canExportColumn( string key )
					{
						if ( exportColumns.TryGetValue( key, out var value ) )
						{
							var v = data[ row, value.Ordinal ];
							return string.IsNullOrEmpty( v?.ToString() ) || (double)v != 0;
						}

						return true;
					}

					var hasRowBasedData = false;

					if ( !hadHistoryRowBasedDeleteColumn )
					{
						foreach ( var header in ExportableColumnsAndExportDelete.Where( h => !h.IsExportControl ) )
						{
							var headerKey = header.Field.ToString();
							var headerHistoryRowKey = header.Field.ToString( FieldFormatType.TableIndexOnly );
							// Check [Status:2018] first, then check *this* field, [Status:2018]field
							var exportColumn = canExportColumn( headerHistoryRowKey ) && canExportColumn( headerKey );

							if ( exportColumn )
							{
								var deleteColumn = false;
								if ( deleteColumns.TryGetValue( headerKey, out var dValue ) )
								{
									var v = data[ row, dValue.Ordinal ];
									deleteColumn = !string.IsNullOrEmpty( v?.ToString() ) && (double)v == 1;
								}

								string? value;

								try
								{
									var dataValue = header.IsText
										? excelReference.Offset( row + 1, header.Ordinal ).GetValue<string>()
										: data[ row, header.Ordinal ];

									value = !deleteColumn && !( header.IsDeleteControl && header.Field.Field == null ) // If no .Field, it is profile delete or history row delete
										? GetExportValue( dataValue, header )
										: null;
								}
								catch ( InteropValueException ex ) when ( header.IsText )
								{
									value = ex.ErrorText;
								}
								catch ( Exception ex )
								{
									throw new ApplicationException( $"Unable to read value from {excelReference.Offset( row + 1, header.Ordinal ).GetAddress()}.", ex );
								}

								if ( header.DeleteIfBlank || deleteColumn || !string.IsNullOrEmpty( value ) )
								{
									hasRowBasedData = true;
									var index = isRowBased ? rowBasedIndex : header.Field.Index;

									if ( !isRowBased && !string.IsNullOrEmpty( index ) )
									{
										var substitutionIndexCol = headers.FirstOrDefault( h => h.Field.HistoryTable == header.Field.HistoryTable && h.Field.Field == "index" && h.Field.Index == header.Field.Index );
										if ( substitutionIndexCol != null )
										{
											var substitutionIndexValue = substitutionIndexCol.IsText
												? excelReference.Offset( row + 1, substitutionIndexCol.Ordinal ).GetValue<string>()
												: data[ row, substitutionIndexCol.Ordinal ];

											index = GetExportValue( substitutionIndexValue, substitutionIndexCol );

											if ( string.IsNullOrEmpty( index ) )
											{
												throw new ApplicationException( $"Unable to export {header.Field} for {authId} because the index value in column {substitutionIndexCol.Field} was blank." );
											}
										}
									}

									var isIndexColumn = header.Field.Field == "index" && header.Field.HistoryTable != null;
									yield return new ExportValue
									{
										IsAuthID = header.IsAuthId,
										Field = header.Field.Field,
										HistoryTable = isRowBased ? RowExportHistoryTable : header.Field.HistoryTable,
										Index = index,
										NewIndex = rowBasedNewIndex,
										Value = !isIndexColumn ? value : null,
										Clear = deleteColumn || ( value == null && header.DeleteIfBlank ),
										AllowReplace = !isIndexColumn && header.AllowReplace
									};
								}
							}
						}

						// If nothing except the index is provided, need to return this value b/c all other 
						// fields above returned blank.
						if ( isRowBased && !hasRowBasedData )
						{
							yield return new ExportValue
							{
								IsAuthID = false,
								Field = "index",
								HistoryTable = RowExportHistoryTable,
								Index = rowBasedIndex,
								NewIndex = rowBasedNewIndex,
							};
						}
					}
				}
			}
		}
	}

	private static string? GetExportValue( object? value, ExportColumn header )
	{
		if ( value == null ) return header.DefaultValue;

		var type = value.GetType();
		var isString = type == typeof( string );
		var isDouble = type == typeof( double );
		var isDecimal = type == typeof( decimal ); // data probably formatted as currency
		var isDateTime = type == typeof( DateTime );
		var isInt = type == typeof( int );

		if ( isString && (string)value == "" ) return header.DefaultValue;

		value.ThrowOnInteropError( type );

		if ( isDateTime )
		{
			return ( (DateTime)value ).ToString( header.IsDateTime ? "yyyy-MM-ddThh:mm:ss" : "yyyy-MM-dd" );
		}
		else if ( isString && header.IsDateTime )
		{
			return Convert.ToDateTime( value ).ToString( "yyyy-MM-ddThh:mm:ss" );
		}
		else if ( isString && header.IsDate )
		{
			return Convert.ToDateTime( value ).ToString( "yyyy-MM-dd" );
		}
		else if ( header.Format != null )
		{
			if ( isInt )
			{
				return ( (int)value ).ToString( header.Format );
			}
			else if ( isDouble )
			{
				return ( (double)value ).ToString( header.Format );
			}

			// Slowest mechanism...
			return string.Format( string.Format( "{{0:{0}}}", header.Format ), value );
		}
		else if ( header.DecimalPlacesToInsert != null )
		{
			if ( isDouble )
			{
				value = (double)value / Math.Pow( 10, header.DecimalPlacesToInsert.Value );
			}
			else
			{
				throw new ArgumentOutOfRangeException( nameof( value ), "To use the /numberConvert flag, you must provide an numerical value." );
			}
		}

		if ( isDecimal && (decimal)value == 0 )
		{
			value = 0d;
			isDouble = true;
		}

		if (
			header.IgnoreZero && (
				( isInt && (int)value == 0 ) ||
				( isDouble && (double)value == 0 )
			)
		)
		{
			return null;
		}

		var @out = value.ToString()!.Trim();

		if ( !string.IsNullOrEmpty( header.DateConvertFormat ) )
		{
			return header.DateConvertFormat switch
			{
				"yyyymmdd" => string.Format( "{0}-{1}-{2}", @out[ ..4 ], @out[ 4..6 ], @out[ 6.. ] ),
				"mmddyyyy" => string.Format( "{0}-{1}-{2}", @out[ 4.. ], @out[ ..2 ], @out[ 2..4 ] ),
				_ => throw new ApplicationException( string.Format( "Invalid format provided: {0}", header.DateConvertFormat ) ),
			};
		}

		if ( header.ToLower )
		{
			return @out.ToLower();
		}
		else if ( header.ToUpper )
		{
			return @out.ToUpper();
		}
		else if ( !header.IsText && header.Field.Field!.StartsWith( "date-" ) && @out.Split( '/' ).Length == 3 )
		{
			return Convert.ToDateTime( value ).ToString( "yyyy-MM-dd" );
		}
		else
		{
			return @out;
		}
	}
}