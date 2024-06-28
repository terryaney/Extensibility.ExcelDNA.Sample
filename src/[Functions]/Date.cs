using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class Date
{
	[KatExcelFunction( 
		Category = "Date Helpers", 
		Description = "Given a `startDate` date and day of week, find the next date whose day of the week equals `desiredDay` and `dateType`.",
		Returns = "`startDate` converted to the first day of the month coincident or following.",
		Remarks = @"1. If `startDate.DayOfWeek` equals `desiredDay` and `dateType` is `Next` or `Previous`, `startDate` is returned.
1. If `dateType` is `PreviousWeek` or `NextWeek`, the `desiredDay` before the previous Sunday or after the next Sunday, respectively, will be returned.
1. If `dateType` is `Previous` or `Next`, the first occurrence of the `desiredDay` in the appropriate direction will be returned.
1. `desiredDay` can be abbreviated to `mon`, `tue`, `wed`, `thu`, `fri`, `sat`, and `sun`.
1. `dateType` can be abbreviated to `pw`, `prevweek`, `nw`, `pd`, `prevday`, `previous`, `prev`, `nd`, and `next`."
	)]
	public static double BTRGetDateForDay(
		[ExcelArgument( "The target date." )]
		DateTime startDate,
		[ExcelArgument( "Monday, Tuesday, ..., Friday representing which day you want." )]
		string desiredDay,
		[ExcelArgument( "Date increment type.  `PreviousWeek`, `NextWeek`, `PreviousDay`, `NextDay`." )]
		string dateType 
	) => Utility.GetDateForDay(  startDate < new DateTime( 1900, 3, 1 ) ? startDate.AddDays( 1 ) : startDate , desiredDay, dateType );
	

	[ExcelFunction( Category = "Date Helpers", Description = "Returns a DateTime object converted to the first day of the month coincident or following." )]
	public static double BTRFirstOfMonthOrCoincident(
		[ExcelArgument( "The target date." )]
		DateTime target 
	) => target.FirstOfMonthOrCoincident().FromLotus();
	

	[ExcelFunction( Category = "Date Helpers", Description = "Returns a target date after adding specified years." )]
	public static double BTRAddYears(
		[ExcelArgument( "The target date." )]
		DateTime target,
		[ExcelArgument( "Number of years to add." )]
		int years 
	) => target.AddYears( years ).FromLotus();

	[ExcelFunction( Category = "Date Helpers", Description = "Returns a target date after adding specified years." )]
	public static double BTRAddMonths(
		[ExcelArgument( "The target date." )]
		DateTime target,
		[ExcelArgument( "Number of months to add." )]
		int months 
	) => target.AddMonths( months ).FromLotus();	

	[ExcelFunction( Category = "Date Helpers", Description = "TODO: HAN: Argument documentation. Determine retirement eligibility." )]
	public static string BTRRetirementEligible(
		int mode,
		long appendix,
		DateTime dateTerm,
		double age,
		double service,
		DateTime dateParticipation,
		object[,] range
	)
	{
		var eligibility = new Utility.EligibilityRow[ range.GetUpperBound( 0 ) + 1 ];
		for ( var i = 0; i < eligibility.Length; i++ )
		{
			eligibility[ i ] = new Utility.EligibilityRow
			{
				Appendix = Convert.ToInt64( range[ i, 0 ] ),
				DateTerm = range[ i, 1 ] != ExcelEmpty.Value ? DateTime.FromOADate( (double)range[ i, 1 ] ) : new DateTime( 1900, 1, 1 ),
				Age = Convert.ToDouble( range[ i, 2 ] ),
				Service = Convert.ToDouble( range[ i, 3 ] ),
				DateParticipation = range[ i, 4 ] != ExcelEmpty.Value ? DateTime.FromOADate( (double)range[ i, 4 ] ) : null,
				Eligibilty = Convert.ToString( range[ i, 5 ] )!
			};
		}
		return Utility.RetirementEligible( mode, appendix, dateTerm, age, service, dateParticipation, eligibility );
	}

	[ExcelFunction( Category = "Date Helpers", Description = "Returns age at target date as fractional years." )]
	public static double BTRAgeAtDate(
		[ExcelArgument( "The date of birth." )]
		DateTime dob,
		[ExcelArgument( "The target date." )]
		DateTime target 
	) => dob.AgeAtDate<double>( target );

	[ExcelFunction( Category = "Date Helpers", Description = "Returns current age as fractional years." )]
	public static double BTRCurrentAge(
		[ExcelArgument( "The date of birth." )]
		DateTime dob 
	) => dob.AgeAtDate<double>( DateTime.Today );

	[ExcelFunction( Category = "Date Helpers", Description = "Returns difference in years, months, full months (integer) or days between two dates." )]
	public static double BTRDateDiff(
		[ExcelArgument( "The start date." )]
		DateTime start,
		[ExcelArgument( "The end date." )]
		DateTime end,
		[KatExcelArgument( Description = "Optional.  1 for Years, 2 for Months, 3 for FullMonths and 4 for Days.  Default is 1.", Type = typeof( int ), Default = "1" )]
		object? interval = null,
		[KatExcelArgument( Description = "Optional.  Whether or not to include the last day as part of the calculation.  Default is false.", Type = typeof( bool ), Default = "false" )]
		object? inclusive = null
	) => start.DateDiff( end, interval.Check( nameof( interval ), 1 ), inclusive.Check( nameof( inclusive ), false ) );
}