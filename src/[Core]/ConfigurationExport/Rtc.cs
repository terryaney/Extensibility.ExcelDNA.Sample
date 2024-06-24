using System.Text.Json.Nodes;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

class Rtc
{
	public void Export( IEnumerable<MSExcel.Worksheet> sheets )
	{
		var fileName = @"C:\BTR\Camelot.Old\Websites\Personal\RTC\rtcsettings.json";

		var clubInfo = sheets.First( s => s.Name == "Club Info" );
		var indoorStartDate = clubInfo.Range[ "B1" ].GetText();
		var indoorHours = clubInfo.Range[ "A2" ];
		var itemCosts = indoorHours.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ];
		var indoorDues = itemCosts.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ];
		var outdoorStartDate = indoorDues.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ].Offset[ 0, 1 ].GetText();
		var outdoorHours = indoorDues.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ].Offset[ 1 ];
		var outdoorDues = outdoorHours.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ];

		// First Hours Of Operation
		var settings = new JsonObject
		{
			{ "RTCSettings", new JsonObject
				{
					{ "Indoor", new JsonObject
						{
							{ "StartDate", !string.IsNullOrEmpty( indoorStartDate ) ? indoorStartDate : null },
							{ "HoursOfOperation", GetKeyValueArrays( indoorHours.Offset[ 1 ] ).ToJsonArray() }
						}
						.AddProperties( new [] { new JsonKeyProperty( "MembershipDues", GetKeyValueArrays( indoorDues.Offset[ 1 ] ).ToJsonArray() ) } )
						.AddProperties( GetPropertyValueArrays( itemCosts.Offset[ 1 ] ) )
					},
					{ "Outdoor", new JsonObject
						{
							{ "StartDate", !string.IsNullOrEmpty( outdoorStartDate ) ? outdoorStartDate : null },
							{ "HoursOfOperation", GetKeyValueArrays( outdoorHours.Offset[ 1 ] ).ToJsonArray() },
							{ "MembershipDues", GetKeyValueArrays( outdoorDues.Offset[ 1 ] ).ToJsonArray() }
						}
					},
					{ "Programs", (
						from sheet in sheets.Where( s => s.Name != "Club Info" && s.Name != "Schedule.Template" && s.Name != "Schedule_Template" )
						select GetSchedules( sheet )
					).ToJsonArray() }
				}
			}
		};

		File.WriteAllText( fileName, JsonExtensions.ToJsonString( settings, writeIndented: true, ignoreNulls: true ) );
	}

	private static IEnumerable<JsonObject> GetKeyValueArrays( MSExcel.Range start )
	{
		var currentRow = 0;
		string key;

		while ( !string.IsNullOrEmpty( key = start.Offset[ currentRow ].GetText() ) )
		{
			yield return
				new JsonObject {
					{ "Key", key },
					{ "Values", GetValues( start.Offset[ currentRow, 1 ], DirectionType.ToRight ).ToJsonArray() }
				};

			currentRow++;
		}
	}

	private static string[] GetValues( MSExcel.Range start, DirectionType direction )
	{
		var startRef = start.GetReference();
		var endRef = startRef.End( direction );

		var data =
			Enumerable.Range( 0, direction == DirectionType.Down ? endRef.RowFirst - startRef.RowFirst + 1 : endRef.ColumnFirst - startRef.ColumnFirst + 1 )
				.Select( i => (string)start.Offset[ direction == DirectionType.Down ? i : 0, direction == DirectionType.ToRight ? i : 0 ].Text )
				.ToArray();

		return data;
	}

	private static JsonObject GetSchedules( MSExcel.Worksheet sheet )
	{
		var name = sheet.Range[ "B1" ];
		var description = sheet.Range[ "B2" ];
		var outdoor = sheet.Range[ "A4" ];
		var indoor = sheet.Range[ "A4" ];

		while ( indoor.GetText() != "Indoor Information" && indoor.Row < 200 )
		{
			indoor = indoor.End[ MSExcel.XlDirection.xlDown ];
		}

		if ( indoor.GetText() != "Indoor Information" )
		{
			throw new ApplicationException( $"Unable to find Indoor Information on {sheet.Name}." );
		}

		return new JsonObject {
			{ "ID", sheet.Name.Replace( "_", "." ) },
			{ "Name", name.GetText() },
			{ "Description", GetValues( description, DirectionType.ToRight ).ToJsonArray() },
			{ "Indoor", GetSchedule( indoor ) },
			{ "Outdoor", GetSchedule( outdoor ) }
		};
	}

	private static JsonObject GetSchedule( MSExcel.Range scheduleInformation )
	{
		var registrationDate = scheduleInformation.Offset[ 1, 1 ].GetText();
		var cost = scheduleInformation.Offset[ 2, 1 ].GetText();
		var sessionsStart = scheduleInformation.Offset[ 4, 0 ];
		var sessions = GetKeyValueArrays( sessionsStart ).ToArray();

		var schedulesStart = sessionsStart.End[ MSExcel.XlDirection.xlDown ];
		while ( schedulesStart.GetText() != "Schedules" && schedulesStart.Row < 200 )
		{
			schedulesStart = schedulesStart.End[ MSExcel.XlDirection.xlDown ];
		}
		if ( schedulesStart.GetText() != "Schedules" )
		{
			throw new ApplicationException( $"Unable to find schedules for {scheduleInformation.Text} on {scheduleInformation.Worksheet.Name}." );
		}
		var schedules = GetScheduleTables( schedulesStart.Offset[ 1 ] ).ToArray();

		var notesStart = schedulesStart.End[ MSExcel.XlDirection.xlDown ];
		while ( notesStart.GetText() != "Notes" && notesStart.Row < 200 )
		{
			notesStart = notesStart.End[ MSExcel.XlDirection.xlDown ];
		}
		if ( notesStart.GetText() != "Notes" )
		{
			throw new ApplicationException( $"Unable to find notes for {scheduleInformation.Text} on {scheduleInformation.Worksheet.Name}." );
		}
		var notes = GetValues( notesStart.Offset[ 1 ], DirectionType.Down );

		return new JsonObject {
			{ "RegistrationDate", !string.IsNullOrEmpty( registrationDate ) ? registrationDate : null },
			{ "Cost", !string.IsNullOrEmpty( cost ) ? cost : null },
			{ "Sessions", sessions.Length > 0 ? new JsonArray( sessions ) : null },
			{ "Schedules", schedules.Length > 0 ? new JsonArray( schedules ) : null },
			{ "Notes", !string.IsNullOrEmpty( notes[ 0 ] ) ? notes.ToJsonArray() : null }
		};
	}

	private static IEnumerable<JsonObject> GetScheduleTables( MSExcel.Range scheduleStart )
	{
		var currentRow = 0;

		while (
			!string.IsNullOrEmpty( scheduleStart.Offset[ currentRow ].GetText() ) && // [ 0, 0 ]
			scheduleStart.Offset[ currentRow ].GetText() != "Notes" && // [ 0, 0 ]
			!string.IsNullOrEmpty( scheduleStart.Offset[ currentRow, 1 ].GetText() ) && // [ 0, 1 ]
			!string.IsNullOrEmpty( scheduleStart.Offset[ currentRow + 1 ].GetText() ) // [1, 0 ]
		)
		{
			var currentScheduleStart = scheduleStart.Offset[ currentRow ];
			var rows = currentScheduleStart.End[ MSExcel.XlDirection.xlDown ].Row - currentScheduleStart.Row + 1;
			var cols = currentScheduleStart.End[ MSExcel.XlDirection.xlToRight ].Column - currentScheduleStart.Column + 1;

			yield return new JsonObject {
				{ "ID", currentScheduleStart.GetText() },
				{ "Rows", 
					Enumerable.Range( 0, rows )
						.Select( r => 
							Enumerable.Range( 0, cols )
								.Select( c => scheduleStart.Offset[ currentRow + r, c ].GetText() )
								.ToJsonArray() 
						)
						.ToJsonArray()
				}
			};

			currentRow += rows + 1;
		}
	}

	private static IEnumerable<JsonKeyProperty> GetPropertyValueArrays( MSExcel.Range start )
	{
		var currentRow = 0;
		string key;

		while ( !string.IsNullOrEmpty( key = start.Offset[ currentRow ].GetText() ) )
		{
			yield return
				new JsonKeyProperty (
					key.Replace( " ", "" ),
					GetValues( start.Offset[ currentRow, 1 ], DirectionType.ToRight ).ToJsonArray()
				);

			currentRow++;
		}
	}
}