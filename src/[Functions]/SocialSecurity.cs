using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaSocialSecurity
{
	[ExcelFunction( Category = "Financial", Description = "Last law year for which yearly indexed functions were updated." )]
	public static int BTRLastLawYear() => SocialSecurity.CurrentYear;

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!SSNRA() function.  Returns a decimal value representing the Social Security normal retirement age." )]
	public static double BTRSSNRA(
		[ExcelArgument( "The member's date of birth." )] 
		DateTime dateBirth,
		[ExcelArgument( "Simplified integer values (e.g. for Covered Compensation and IRC 415 limit purposes), otherwise actual values for Social Security use." )]
		object? simplifiedResultsArg = null )
	{
		var simplifiedResults = simplifiedResultsArg.Check( "simplifiedResults", false );
		return SocialSecurity.NormalRetirementAge( dateBirth, simplifiedResults );
	}

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!SSNRD() function.  Returns a decimal value representing the Social Security normal retirement date." )]
	public static DateTime BTRSSNRD(
		[ExcelArgument( "The member's date of birth." )] 
		DateTime dateBirth,
		[ExcelArgument( "Simplified integer values (e.g. for Covered Compensation and IRC 415 limit purposes), otherwise actual values for Social Security use." )] 
		object? simplifiedResultsArg = null )
	{
		var simplifiedResults = simplifiedResultsArg.Check( "simplifiedResults", false );
		return SocialSecurity.NormalRetirementDate( dateBirth, simplifiedResults );
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!SSExactPays() function using a single current pay value.  Returns a decimal value representing the age 65 Social Security monthly benefit using exact pays." )]
	public static double BTRSSPIA( 
		[ExcelArgument( "The member's date of birth." )]
		DateTime dateBirth,
		[ExcelArgument( "The member's date of event (i.e. Date Term)." )]
		DateTime dateEvent,
		[ExcelArgument( "The member's age at retirement." )]
		double ageRetire,
		[ExcelArgument( "The member's annual compensation for current Social Security year.  Annual compensation will be projected for any missing years from age 18 through the year before payment start." )]
		double payCurrent,
		[ExcelArgument( "NAW increase rate." )]
		object? rateNAW = null,
		[ExcelArgument( "Future pay increase rate." )]
		object? rateFuturePay = null,
		[ExcelArgument( "Backward pay increase rate." )]
		double rateBackPay = 0,
		[ExcelArgument( "Add NAW to backward pay increase rate (TRUE or FALSE)." )]
		object? addNAWToBackPay = null,
		[ExcelArgument( "CPI increase rate." )]
		object? rateCOLA = null,
		[ExcelArgument( "Social Security law year." )]
		int lawYear = 0,
		[ExcelArgument( "Type of pay after the year before termination year: C=project one more year at rateFuturePay then stay constant till the year before commencement year, L= stay constant till the year before commencement year., Z=zero pay starting termination year." )]
		object? futurePayType = null,
		[ExcelArgument( "The first age when emeber started receiving compensation." )]
		object? firstPayAge = null,
		[ExcelArgument( "Stop NAW growth after termination?" )]
		object? stopNAWGrowth = null,
		[ExcelArgument("Last pay year")]
		object? lastPayYear = null,
		[ExcelArgument("Include post NRD Increase?")]
		object? postNRDIncrease = null,
		[ExcelArgument("Year COLA Stops")]
		object? yearCOLAStops = null 
	)
	{
		var lastPayYearArg = lastPayYear.Check( nameof( lastPayYear ), 0 );
		var postNRDIncreaseArg = postNRDIncrease.Check( nameof( postNRDIncrease ), true );
		var yearCOLAStopsArg = yearCOLAStops.Check( nameof( yearCOLAStops ), 0 );

		return BTRSSPIASalHist( dateBirth, dateEvent, ageRetire, new[] { payCurrent }, rateNAW, rateFuturePay, rateBackPay, addNAWToBackPay, rateCOLA, lastPayYearArg, lawYear, futurePayType, firstPayAge, stopNAWGrowth, postNRDIncreaseArg, yearCOLAStopsArg );
	}

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!SSTableProj() function using a pay array.  Returns a decimal value representing the age 65 Social Security monthly benefit using exact pays." )]
	public static double BTRSSPIASalHist(
		[ExcelArgument( "The member's date of birth." )]
		DateTime dateBirth,
		[ExcelArgument( "The member's date of event (i.e. Date Term)." )]
		DateTime dateEvent,
		[ExcelArgument( "The member's age at retirement." )]
		double ageRetire,
		[ExcelArgument( "The member's annual compensations ending at current Social Security year.  Annual compensation will be projected for any missing years from age 18 through the year before payment start." )]
		double[] actualPay,
		[ExcelArgument( "NAW increase rate." )]
		object? rateNAW = null,
		[ExcelArgument( "Future pay increase rate." )]
		object? rateFuturePay = null,
		[ExcelArgument( "Backward pay increase rate." )]
		double rateBackPay = 0,
		[ExcelArgument( "Add NAW to backward pay increase rate (TRUE or FALSE)." )]
		object? addNAWToBackPay = null,
		[ExcelArgument( "CPI increase rate." )]
		object? rateCOLA = null,
		[ExcelArgument( "Ending year of compensation in the compensation array." )]
		int lastPayYear = 0,
		[ExcelArgument( "Social Security law year." )]
		int lawYear = 0,
		[ExcelArgument( "Type of pay after the year before termination year: C=project one more year at rateFuturePay then stay constant till the year before commencement year, L= stay constant till the year before commencement year., Z=zero pay starting termination year." )]
		object? futurePayType = null,
		[ExcelArgument( "The first age when meeber started receiving compensation." )]
		object? firstPayAge = null,
		[ExcelArgument( "Stop NAW growth after termination?" )]
		object? stopNAWGrowth = null,
		[ExcelArgument("Include post NRD increase?")]
		object? postNRDIncrease = null,
		[ExcelArgument("Year COLA Stops")]
		object? yearCOLAStops = null 
	)
	{
		var rateNAWArg = rateNAW.Check( nameof( rateNAW ), 0.045 );
		var rateFuturePayArg = rateFuturePay.Check( nameof( rateFuturePay ), 0.05 );
		var addNAWToBackPayArg = addNAWToBackPay.Check( nameof( addNAWToBackPay ), true );
		var rateCOLAArg = rateCOLA.Check( nameof( rateCOLA ), 0.04 );
		var futurePayTypeArg = futurePayType.Check( nameof( futurePayType ), "C" );
		var firstPayAgeArg = firstPayAge.Check( nameof( firstPayAge ), 18 );
		var stopNAWGrowthArg = stopNAWGrowth.Check( nameof( stopNAWGrowth ), true );
		var postNRDIncreaseArg = postNRDIncrease.Check( nameof( postNRDIncrease ), true );
		var yearCOLAStopsArg = yearCOLAStops.Check( nameof( yearCOLAStops ), 0 );

		return SocialSecurity.PrimaryInsuranceAmount( dateBirth, dateEvent, ageRetire, actualPay, rateNAWArg, rateFuturePayArg, rateBackPay, addNAWToBackPayArg, rateCOLAArg, lastPayYear, lawYear, futurePayTypeArg, firstPayAgeArg, stopNAWGrowthArg, postNRDIncreaseArg, yearCOLAStopsArg ) ??
			throw new NullReferenceException( "arrayNAW not setup correctly." );
	}

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!WageBase() function.  Returns a decimal value of Social Security Wage Base at yearEvent parameter." )]
	public static double BTRWageBase(
		[ExcelArgument( "Year of requested wage base." )]
		int yearEvent,
		[ExcelArgument( "NAW increase rate, defaulted to 4.5%." )]
		object? rateNAW = null,
		[ExcelArgument( "SS law year, defaulted to current law year." )]
		int lawYear = 0,
		[ExcelArgument( "Whether or not to apply the $300 rounding." )]
		bool unrounded = false 
	)
	{
		var rateNAWArg = rateNAW.Check( nameof( rateNAW ), 0.045 );

		return SocialSecurity.WageBase( yearEvent, rateNAWArg, lawYear, unrounded );
	}

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!SSNAW() function.  Returns a decimal value of Social Security National Average Wage Base at yearEvent parameter." )]
	public static double BTRNAW(
		[ExcelArgument( "Year of requested wage base." )]
		int yearEvent,
		[ExcelArgument( "NAW increase rate, defaulted to 4.5%." )]
		object? rateNAW = null,
		[ExcelArgument( "SS law year, defaulted to current law year." )]
		int lawYear = 0 
	)
	{
		var rateNAWArg = rateNAW.Check( nameof( rateNAW ), 0.045 );

		return SocialSecurity.AverageWageBase( yearEvent, rateNAWArg, lawYear );
	}

	[ExcelFunction( Category = "Financial", Description = "A replacement function for the Cfgena.xla!CoveredComp() function.  Returns a decimal value of covered compensation at Social Security Normal Retirement Age or at yearEvent parameter." )]
	public static double BTRCoveredComp(
		[ExcelArgument( "The member's year of birth." )]
		int yearBirth,
		[ExcelArgument( "NAW increase rate, defaulted to 4.5%." )]
		object? rateNAW = null,
		[ExcelArgument( "Year of requested covered compensation." )]
		int yearEvent = 0,
		[ExcelArgument( "SS law year, defaulted to current law year." )]
		int lawYear = 0,
		[ExcelArgument( "Transition rule (when value is 1, else defaulted to 0), when value is 1 then the averaging of 35 years of wage base ended at the year before yearEvent else it ended at 'yearEvent'." )]
		int flagTransition = 0,
		[ExcelArgument( "Rounding, defaulted to 0, means rounded down to 12, other alues are 12 (same as 0), 1, 300, 600 or 6000 means rounded to the nearest value, else unrounded." )]
		int optRounding = 0 
	)
	{
		var rateNAWArg = rateNAW.Check( nameof( rateNAW ), 0.045 );
		return SocialSecurity.CoveredCompensation( yearBirth, rateNAWArg, yearEvent, lawYear, flagTransition, optRounding );
	}
}