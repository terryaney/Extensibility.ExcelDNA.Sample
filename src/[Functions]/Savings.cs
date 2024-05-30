using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaSavings
{
	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns plan investment annual return based on Fund allocations and how Funds are mapped to assets classes." )]
	public static double BTRAnnualReturn(
		[ExcelArgument( "Required.  Fund Table Name." )]
		string tableName,
		[ExcelArgument( "Required.  Plan type." )]
		int planType,
		[ExcelArgument( "Required.  Year of requested returns." )]
		int year,
		[ExcelArgument( "Required.  Current year." )]
		int yearCurrent,
		[ExcelArgument( "Required.  Current fund allocations." )]
		double[ , ] fundAllocations,
		[ExcelArgument( "Required.  Entered fund allocations (this will override current or future allocations)." )]
		double[ , ] inputAllocations,
		[ExcelArgument( "Required.  Array of investment returns by assets classes." )]
		double[ , ] returnByClass,
		[ExcelArgument( "Optional.  If provided allocation will be changed 100% to that target fund." )]
		string fundIndex 
	) => Savings.AnnualReturn( tableName, (FundPlanType)planType, year, yearCurrent, fundAllocations, inputAllocations, returnByClass, fundIndex ) ??
			throw new NullReferenceException( "MapFundToAssetClass/GetClassAllocations is null." );

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns plan asset class allocation based on Fund allocations and how Funds are mapped to assets classes.", IsMacroType = true )]
	public static double[] BTRAssetClassAllocation(
		[ExcelArgument( "Required.  Fund Table Name." )]
		string tableName,
		[ExcelArgument( "Required.  Plan type." )]
		int planType,
		[ExcelArgument( "Required.  Year of requested allocation." )]
		int year,
		[ExcelArgument( "Required.  Current fund allocations." )]
		double[ , ] fundAllocations,
		[ExcelArgument( AllowReference = true, Description = "Optional.  Entered fund allocations (this will override current or future allocations)." )]
		object? inputAllocations = null,
		[ExcelArgument( "Optional.  If provided allocation will be changed 100% to that target fund." )]
		object? fundIndex = null )
	{
		var inputAllocationsArg = inputAllocations as ExcelReference; //  OptionalValues.Check<double[ , ]>( inputAllocations, nameof( inputAllocations ), null );
		var inputAllocationValues = inputAllocationsArg?.GetArray<double>(); // This will be null if no param passed in.
		var fundIndexArg = fundIndex.Check<string?>( nameof( fundIndex ), null );
		var allocations = 
			Savings.AssetClassAllocation( tableName, (FundPlanType)planType, year, fundAllocations, inputAllocationValues, fundIndexArg ) ?? 
			throw new NullReferenceException( "Allocation is null."	);

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( allocations.Any( r => double.IsNaN( r ) ) )
		{
			return allocations.Select( r => double.NaN ).ToArray();
		}

		return allocations;
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns index of Target Fund based on retirement year and target fund group." )]
	public static string BTRTargetFundID(
		[ExcelArgument( "Fund Table Name." )]
		string tableName,
		[ExcelArgument( "Year of retirement." )]
		int yearRetirement,
		[ExcelArgument( "Target fund group ID." )]
		int targetFundGroup 
	) => Savings.TargetFundID( tableName, yearRetirement, targetFundGroup );
	
	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns an array with 3 values: balance at EOY, principle payments with investment return at EOY and principle payments with investment return at EOB." )]
	public static double[] BTRCalculateLoan(
		[ExcelArgument( "Required. Loan balance." )]
		double balance,
		[ExcelArgument( "Required. Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required. Payment per pay period." )]
		double paymentPerPayPeriod,
		[ExcelArgument( "Required. Loan interest rate." )]
		double interestRate,
		[ExcelArgument( "Required. Calculate contributions as of the end of this month." )]
		int monthEnd,
		[ExcelArgument( "Required. Investment rate of return." )]
		double rateOfReturn,
		[ExcelArgument( Description = "Optional. Starting pay period. Default to 1." )]
		object? startPayPeriod = null,
		[ExcelArgument( Description = "Optional. Use mid point contributions timing calculation method." )]
		object? midPointContribution = null,
		[ExcelArgument( "Optional. Rounding, Defaulted to 0 decimals" )]
		int rounding = 0 )
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );
		var midPointContributionArg = midPointContribution.Check( nameof( midPointContribution ), false );
		return Savings.CalculateLoan( balance, payPeriod, paymentPerPayPeriod, interestRate, monthEnd, rateOfReturn, startPayPeriodArg, midPointContributionArg, rounding );
	}

	[DebugFunction]
	[ExcelFunctionDoc(
		Category = "Financial",
		Description = "DOC: Han, Cfgena replacement?  Returns 401k contributions/match.",
		Remarks = "<p>Both 'matchParam' and 'contributionParam' are a | delimited list of periods.  Each period is in the form of <i>M:P:P:P</i> where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values.</p><p>The 'matchType' and 'contributionType' determine how each tier of numbers are processed.  Each tier's value pairs are described below:</p><ul><li>CalculationMatchType.MultiplierBasedOnPercent - ContrbutionPercent, Multiplier</li><li>CalculationMatchType.MultiplierBasedOnDollars - ContributionDollars, Multiplier</li><li>CalculationMatchType.ERMatchPercentBasedOnPercent - EEContributionPercent, ERContributionPercent</li><li>CalculationMatchType.ERMatchDollarsBasedOnDollars - ContributionDollars, ERContributionDollars</li><li>CalculationContributionType.PercentBasedOnAge - Age, ContributionPercent</li><li>CalculationContributionType.PercentBasedOnService - Service, ContributionPercent</li><li>CalculationContributionType.PercentBasedOnAgePlusService - AgePlusService, ContributionPercent</li><li>CalculationContributionType.DollarsBasedOnAge - Age, ContributionDollars</li><li>CalculationContributionType.DollarsBasedOnService - Service, ContributionDollars</li><li>CalculationContributionType.DollarsBasedOnAgePlusService - AgePlusService, ContributionDollars</li></ul>"
	)]
	public static double[,] BTRGetContributions(
		[ExcelArgument( "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, MultiplierBasedOnDollars = 2, ERMatchPercentBasedOnPercent = 3, ERMatchDollarsBasedOnDollars = 4." )]
		int matchType,
		[ExcelArgument( "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool isTrueup,
		[ExcelArgument( "Required.  Whether to provide match on catch-up contributions." )]
		bool isCatchupMatch,
		[ExcelArgument( "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool isPretaxOverflowToAftertax,
		[ExcelArgument( "Required.  Whether to provide match on after-tax contributions." )]
		bool isAftertaxMatch,
		[ExcelArgument( "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string matchParam,
		[ExcelArgument( "Required.  The ContributionType to use for calculations.  PercentBasedOnAge = 1, PercentBasedOnService = 2, PercentBasedOnAgePlusService = 3." )]
		int contributionType,
		[ExcelArgument( "Required.  Contribution parameters.  See matchParm for info." )]
		string contributionParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate contributions as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required.  Annual Pay rate as of start pay period." )]
		double rateOfPay,
		[ExcelArgument( "Required.  Pay period when pay increases." )]
		int payPeriodWhenPayIncreases,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double ratePayIncrease,
		[ExcelArgument( "Required.  Inflation rate (used to project limits)." )]
		double rateOfInflation,
		[ExcelArgument( "Required.  Investment rate of return." )]
		double rateOfReturn,
		[ExcelArgument( "Optional.  Pre-tax contribution as a % of pay." )]
		double pretaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double pretaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth contribution as a % of pay." )]
		double rothPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth contribution amount per pay period." )]
		double rothFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  After-tax contribution as a % of pay." )]
		double aftertaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double aftertaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Starting Pay period.  Defaults to 1." )]
		object? startPayPeriod = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax contributions." )]
		double ytdPretax = 0d,
		[ExcelArgument( "Optional.  YTD Roth contributions." )]
		double ytdRoth = 0d,
		[ExcelArgument( "Optional.  YTD after-tax contributions." )]
		double ytdAftertax = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax catch-up contributions." )]
		double ytdPretaxCatchup = 0d,
		[ExcelArgument( "Optional.  YTD Roth catch-up contributions." )]
		double ytdRothCatchup = 0d,
		[ExcelArgument( "Optional.  YTD employer match contributions." )]
		double ytdERMatch = 0d,
		[ExcelArgument( "Optional.  YTD employer contributions." )]
		double ytdERCont = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Service at BOY." )]
		double svcBOY = 0d,
		[ExcelArgument( "Optional.  Employer contribution annual limit. Defaults to 0, which means unlimited." )]
		double erContributionAnnualLimit = 0d,
		[ExcelArgument( "Optional.  DOC: Han, Is this retirement year?" )]
		bool isRetirement = false,
		[ExcelArgument( "Optional.  Frequency of match in a year." )]
		int matchFreq = 0,
		[ExcelArgument( "Optional.  Frequency of ER Contribution in a year." )]
		int erContFreq = 0,
		[ExcelArgument( "Optional.  Employee Pretax/Roth & Catchup Contributions is based on limited pay." )]
		object? pretaxRothPayIsLimited = null,
		[ExcelArgument( "Optional.  Use mid point contributions timing calculation method." )]
		bool midPointContribution = false,
		[ExcelArgument( "Optional.  Don't apply IRS limit." )]
		bool noLimit = false,
		[ExcelArgument( "Optional.  Pretax auto increase timing (month): Enter 4 if increase happens on 4/1." )]
		int increaseMonth = 0,
		[ExcelArgument( "Optional.  Pretax auto increase frequency per year." )]
		int increaseFreq = 0,
		[ExcelArgument( "Optional.  Pretax auto increase percentage." )]
		double increasePct = 0,
		[ExcelArgument( "Optional.  Pretax maximum rate." )]
		double maxRate = 0d,
		[ExcelArgument( "Optional.  Pay Limit." )]
		double limitPay = 0d,
		[ExcelArgument( "Optional.  Pretax limit." )]
		double limitPretax = 0d,
		[ExcelArgument( "Optional.  Contribution limit." )]
		double limitContribution = 0d,
		[ExcelArgument( "Optional.  Overflow to non qualified plan." )]
		bool isOverflowToNonQual = false,
		[ExcelArgument( "Optional.  Aftertax limit." )]
		double limitAftertax = 0d,
		[ExcelArgument( "Optional.  Overflow from pretax to catchup." )]
		object? isOverflowToCatchup = null,
		[ExcelArgument( "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double pretaxCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double pretaxCatchupFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth catchup contribution as a % of pay." )]
		double rothCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double rothCatchupFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Employer match is based on limited pay." )]
		object? matchPayIsLimited = null
	)
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );
		//var isRetirementArg = isRetirement.Check( nameof( isRetirement ), false );
		var pretaxRothPayIsLimitedArg = pretaxRothPayIsLimited.Check( nameof( pretaxRothPayIsLimited ), true );
		//var midPointContributionArg = midPointContribution.Check( nameof( midPointContribution ), false );
		//var noLimitArg = noLimit.Check( nameof( noLimit ), false );
		var isOverflowToCatchupArg = isOverflowToCatchup.Check( nameof( isOverflowToCatchup ), true );
		var matchPayIsLimitedArg = matchPayIsLimited.Check( nameof( matchPayIsLimited ), true );

		var contributions = Savings.GetContributions( ResultType.ArrayOfContributions, (CalculationMatchType)matchType, isTrueup, isCatchupMatch, isPretaxOverflowToAftertax, isAftertaxMatch, matchParam,
			(CalculationContributionType)contributionType, contributionParam, year, monthEnd, payPeriod, rateOfPay, payPeriodWhenPayIncreases, ratePayIncrease, rateOfInflation, rateOfReturn,
			pretaxPct, pretaxFlatPerPayPeriod, rothPct, rothFlatPerPayPeriod, aftertaxPct, aftertaxFlatPerPayPeriod, startPayPeriodArg,
			ytdPay, ytdPretax, ytdRoth, ytdAftertax, ytdPretaxCatchup, ytdRothCatchup, ytdERMatch, ytdERCont, ageBOY, svcBOY, erContributionAnnualLimit, isRetirement, matchFreq,
			erContFreq, pretaxRothPayIsLimitedArg, midPointContribution, noLimit, increaseMonth, increaseFreq, increasePct, maxRate,
			limitPay, limitPretax, limitContribution, isOverflowToNonQual, limitAftertax, isOverflowToCatchupArg, pretaxCatchupPct, pretaxCatchupFlatPerPayPeriod, rothCatchupPct, rothCatchupFlatPerPayPeriod, matchPayIsLimitedArg ) ?? 
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns 401(k) contributions/match." )]
	public static double[,] BTRGet401kContributions(
		[ExcelArgument( "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, MultiplierBasedOnDollars = 2, ERMatchPercentBasedOnPercent = 3, ERMatchDollarsBasedOnDollars = 4." )]
		int matchType,
		[ExcelArgument( "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool isTrueup,
		[ExcelArgument( "Required.  Whether to provide match on catch-up contributions." )]
		bool isCatchupMatch,
		[ExcelArgument( "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool isPretaxOverflowToAftertax,
		[ExcelArgument( "Required.  Whether to provide match on after-tax contributions." )]
		bool isAftertaxMatch,
		[ExcelArgument( "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string matchParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate contributions as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required.  Annual Pay rate as of start pay period." )]
		double rateOfPay,
		[ExcelArgument( "Required.  Pay period when pay increases." )]
		int payPeriodWhenPayIncreases,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double ratePayIncrease,
		[ExcelArgument( "Required.  Inflation rate (used to project limits)." )]
		double rateOfInflation,
		[ExcelArgument( "Required.  Investment rate of return." )]
		double rateOfReturn,
		[ExcelArgument( "Optional.  Pre-tax contribution as a % of pay." )]
		double PretaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double PretaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth contribution as a % of pay." )]
		double rothPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth contribution amount per pay period." )]
		double rothFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  After-tax contribution as a % of pay." )]
		double AftertaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double AftertaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Starting Pay period.  Defaults to 1." )]
		object? startPayPeriod = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax contributions." )]
		double ytdPretax = 0d,
		[ExcelArgument( "Optional.  YTD Roth contributions." )]
		double ytdRoth = 0d,
		[ExcelArgument( "Optional.  YTD after-tax contributions." )]
		double ytdAftertax = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax catch-up contributions." )]
		double ytdPretaxCatchup = 0d,
		[ExcelArgument( "Optional.  YTD Roth catch-up contributions." )]
		double ytdRothCatchup = 0d,
		[ExcelArgument( "Optional.  YTD employer match contributions." )]
		double ytdERMatch = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Is this retirement year?" )]
		bool isRetirement = false,
		[ExcelArgument( "Optional.  Frequency of match in a year." )]
		int matchFreq = 0,
		[ExcelArgument( "Optional.  Employee Pretax/Roth & Catchup Contributions is based on limited pay." )]
		object? pretaxRothPayIsLimited = null,
		[ExcelArgument( "Optional.  Use mid point contributions timing calculation method." )]
		bool midPointContribution = false,
		[ExcelArgument( "Optional.  Don't apply IRS limit." )]
		bool noLimit = false,
		[ExcelArgument( "Optional.  Pretax auto increase timing (month): Enter 4 if increase happens on 4/1." )]
		int increaseMonth = 0,
		[ExcelArgument( "Optional.  Pretax auto increase frequency per year." )]
		int increaseFreq = 0,
		[ExcelArgument( "Optional.  Pretax auto increase percentage." )]
		double increasePct = 0,
		[ExcelArgument( "Optional.  Pretax maximum rate." )]
		double maxRate = 0d,
		[ExcelArgument( "Optional.  Pay Limit." )]
		double limitPay = 0d,
		[ExcelArgument( "Optional.  Pretax limit." )]
		double limitPretax = 0d,
		[ExcelArgument( "Optional.  Contribution limit." )]
		double limitContribution = 0d,
		[ExcelArgument( "Optional.  Overflow to non qualified plan." )]
		bool isOverflowToNonQual = false,
		[ExcelArgument( "Optional.  Aftertax limit." )]
		double limitAftertax = 0d,
		[ExcelArgument( "Optional.  Overflow from pretax to catchup." )]
		object? isOverflowToCatchup = null,
		[ExcelArgument( "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double pretaxCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double pretaxCatchupFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth catchup contribution as a % of pay." )]
		double rothCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double rothCatchupFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Employer match is based on limited pay." )]
		object? matchPayIsLimited = null
	)
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );
		//var isRetirementArg = isRetirement.Check( nameof( isRetirement ), false );
		var pretaxRothPayIsLimitedArg = pretaxRothPayIsLimited.Check( nameof( pretaxRothPayIsLimited ), true );
		//var midPointContributionArg = midPointContribution.Check( nameof( midPointContribution ), false );
		//var noLimitArg = noLimit.Check( nameof( noLimit ), false );
		var isOverflowToCatchupArg = isOverflowToCatchup.Check( nameof( isOverflowToCatchup ), true );
		var matchPayIsLimitedArg = matchPayIsLimited.Check( nameof( matchPayIsLimited ), true );

		var contributions = Savings.GetContributions( ResultType.ArrayOf401kContributions, (CalculationMatchType)matchType, isTrueup, isCatchupMatch, isPretaxOverflowToAftertax, isAftertaxMatch, matchParam,
			null, "", year, monthEnd, payPeriod, rateOfPay, payPeriodWhenPayIncreases, ratePayIncrease, rateOfInflation, rateOfReturn,
			PretaxPct, PretaxFlatPerPayPeriod, rothPct, rothFlatPerPayPeriod, AftertaxPct, AftertaxFlatPerPayPeriod, startPayPeriodArg,
			ytdPay, ytdPretax, ytdRoth, ytdAftertax, ytdPretaxCatchup, ytdRothCatchup, ytdERMatch, 0d, ageBOY, 0d, 0d, isRetirement, matchFreq, 0,
			pretaxRothPayIsLimitedArg, midPointContribution, noLimit, increaseMonth, increaseFreq, increasePct, maxRate,
			limitPay, limitPretax, limitContribution, isOverflowToNonQual, limitAftertax, isOverflowToCatchupArg, pretaxCatchupPct, pretaxCatchupFlatPerPayPeriod, rothCatchupPct, rothCatchupFlatPerPayPeriod, matchPayIsLimitedArg ) ?? 
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer match % based on match parameters." )]
	public static double[] BTRGetMatchPct(
		[ExcelArgument( "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, ERMatchBasedOnPercent = 3" )]
		int matchType,
		[ExcelArgument( "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool isTrueup,
		[ExcelArgument( "Required.  Whether to provide match on catch-up contributions." )]
		bool isCatchupMatch,
		[ExcelArgument( "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool isPretaxOverflowToAftertax,
		[ExcelArgument( "Required.  Whether to provide match on after-tax contributions." )]
		bool isAftertaxMatch,
		[ExcelArgument( "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string matchParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate match as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required.  Annual Pay rate as of start pay period." )]
		double rateOfPay,
		[ExcelArgument( "Required.  Pay period when pay increases." )]
		int payPeriodWhenPayIncreases,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double rateOfPayIncrease,
		[ExcelArgument( "Required.  Inflation rate (used to project limits)." )]
		double rateOfInflation,
		[ExcelArgument( "Optional.  Pre-tax contribution as a % of pay." )]
		double PretaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double PretaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth contribution as a % of pay." )]
		double rothPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth contribution amount per pay period." )]
		double rothFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  After-tax contribution as a % of pay." )]
		double AftertaxPct = 0d,
		[ExcelArgument( "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double AftertaxFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Starting Pay period.  Defaults to 1." )]
		object? startPayPeriod = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax contributions." )]
		double ytdPretax = 0d,
		[ExcelArgument( "Optional.  YTD Roth contributions." )]
		double ytdRoth = 0d,
		[ExcelArgument( "Optional.  YTD after-tax contributions." )]
		double ytdAftertax = 0d,
		[ExcelArgument( "Optional.  YTD pre-tax catch-up contributions." )]
		double ytdPretaxCatchup = 0d,
		[ExcelArgument( "Optional.  YTD Roth catch-up contributions." )]
		double ytdRothCatchup = 0d,
		[ExcelArgument( "Optional.  YTD employer match contributions." )]
		double ytdERMatch = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Frequency of match in a year." )]
		int matchFreq = 0,
		[ExcelArgument( "Optional.  Aftertax limit." )]
		double limitAftertax = 0d,
		[ExcelArgument( "Optional.  Overflow from pretax to catchup." )]
		object? isOverflowToCatchup = null,
		[ExcelArgument( "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double pretaxCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double pretaxCatchupFlatPerPayPeriod = 0d,
		[ExcelArgument( "Optional.  Roth catchup contribution as a % of pay." )]
		double rothCatchupPct = 0d,
		[ExcelArgument( "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double rothCatchupFlatPerPayPeriod = 0d 
	)
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );
		var isOverflowToCatchupArg = isOverflowToCatchup.Check( nameof( isOverflowToCatchup ), true );

		var contributions = Savings.GetContributions( ResultType.MatchPct, (CalculationMatchType)matchType, isTrueup, isCatchupMatch, isPretaxOverflowToAftertax, isAftertaxMatch, matchParam,
			null, "", year, monthEnd, payPeriod, rateOfPay, payPeriodWhenPayIncreases, rateOfPayIncrease, rateOfInflation, 0d,
			PretaxPct, PretaxFlatPerPayPeriod, rothPct, rothFlatPerPayPeriod, AftertaxPct, AftertaxFlatPerPayPeriod, startPayPeriodArg,
			ytdPay, ytdPretax, ytdRoth, ytdAftertax, ytdPretaxCatchup, ytdRothCatchup, ytdERMatch, 0d, ageBOY, 0d, 0d, false, matchFreq, 0,
			true, false, false, 0, 0, 0d, 0d, 0d, 0d, 0d, false, limitAftertax, isOverflowToCatchupArg, pretaxCatchupPct, pretaxCatchupFlatPerPayPeriod, rothCatchupPct, rothCatchupFlatPerPayPeriod ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN ).GetRow( 0 );
		}

		return contributions.GetRow( 0 );
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer contributions." )]
	public static double[,] BTRGetERContribution(
		[ExcelArgument( "Required.  The ContributionType to use for calculations.  PercentBasedOnAge = 1, PercentBasedOnService = 2, PercentBasedOnAgePlusService = 3." )]
		int contributionType,
		[ExcelArgument( "Required.  Contribution parameters.  See matchParm for info." )]
		string contributionParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate contributions as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required.  Annual Pay rate as of start pay period." )]
		double rateOfPay,
		[ExcelArgument( "Required.  Pay period when pay increases." )]
		int payPeriodWhenPayIncreases,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double rateOfPayIncrease,
		[ExcelArgument( "Required.  Inflation rate (used to project limits)." )]
		double rateOfInflation,
		[ExcelArgument( "Required.  Investment rate of return." )]
		double rateOfReturn,
		[ExcelArgument( "Optional.  Starting Pay period.  Defaults to 1." )]
		object? startPayPeriod = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[ExcelArgument( "Optional.  YTD employer contributions." )]
		double ytdERCont = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Service at BOY." )]
		double svcBOY = 0d,
		[ExcelArgument( "Optional.  Employer contribution annual limit. Defaults to unlimited." )]
		double erContributionAnnualLimit = 0d,
		[ExcelArgument( "Optional.  Frequency of ER Contribution in a year." )]
		int erContFreq = 0,
		[ExcelArgument( "Optional.  Use mid point contributions timing calculation method." )]
		object? midPointContribution = null,
		[ExcelArgument( "Optional.  Don't apply IRS limit." )]
		object? noLimit = null,
		[ExcelArgument( "Optional.  Pay Limit." )]
		double limitPay = 0d,
		[ ExcelArgument( "Optional.  Contribution limit." ) ]
		double limitContribution = 0d,
		[ ExcelArgument( "Optional.  Overflow to non qualified plan." ) ]
		bool isOverflowToNonQual = false 
	)
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );
		var midPointContributionArg = midPointContribution.Check( nameof( midPointContribution ), false );
		var noLimitArg = noLimit.Check( nameof( noLimit ), false );

		var contributions = Savings.GetContributions( ResultType.ContributionPct, null, false, false, false, false, "", (CalculationContributionType)contributionType, contributionParam,
			year, monthEnd, payPeriod, rateOfPay, payPeriodWhenPayIncreases, rateOfPayIncrease, rateOfInflation, rateOfReturn, 0d, 0d, 0d, 0d, 0d, 0d,
			startPayPeriodArg, ytdPay, 0d, 0d, 0d, 0d, 0d, 0d, ytdERCont, ageBOY, svcBOY, erContributionAnnualLimit, false, 0, erContFreq, true, midPointContributionArg, noLimitArg,
			0, 0, 0d, 0d, limitPay, 0d, limitContribution, isOverflowToNonQual ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer contribution % based on contribution parameters." )]
	public static double[] BTRGetERContributionPct(
		[ExcelArgument( "Required.  The ContributionType to use for calculations.  PercentBasedOnAge = 1, PercentBasedOnService = 2, PercentBasedOnAgePlusService = 3." )]
		int contributionType,
		[ExcelArgument( "Required.  Contribution parameters.  See matchParm for info." )]
		string contributionParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate contributions as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[ExcelArgument( "Required.  Annual Pay rate as of start pay period." )]
		double rateOfPay,
		[ExcelArgument( "Required.  Pay period when pay increases." )]
		int payPeriodWhenPayIncreases,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double rateOfPayIncrease,
		[ExcelArgument( "Required.  Inflation rate (used to project limits)." )]
		double rateOfInflation,
		[ExcelArgument( "Required.  Investment rate of return." )]
		object? startPayPeriod = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[ExcelArgument( "Optional.  YTD employer contributions." )]
		double ytdERCont = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Service at BOY." )]
		double svcBOY = 0d,
		[ExcelArgument( "Optional.  Employer contribution annual limit. Defaults to unlimited." )]
		double erContributionAnnualLimit = 0d,
		[ExcelArgument( "Optional.  Frequency of ER Contribution in a year." )]
		int erContFreq = 0 
	)
	{
		var startPayPeriodArg = startPayPeriod.Check( nameof( startPayPeriod ), 1 );

		var contributions = Savings.GetContributions( ResultType.ContributionPct, null, false, false, false, false, "", (CalculationContributionType)contributionType, contributionParam,
			year, monthEnd, payPeriod, rateOfPay, payPeriodWhenPayIncreases, rateOfPayIncrease, rateOfInflation, 0d, 0d, 0d, 0d, 0d, 0d, 0d,
			startPayPeriodArg, ytdPay, 0d, 0d, 0d, 0d, 0d, 0d, ytdERCont, ageBOY, svcBOY, erContributionAnnualLimit, false, 0, erContFreq ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN ).GetRow( 0 );
		}

		return contributions.GetRow( 0 );
	}

	[DebugFunction]
	[ExcelFunction( Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns minimum employee contribution to get max mactch based on match parameters." )]
	public static double[,] BTRGetMinContributionsForMaxMatch(
		[ExcelArgument( "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, ERMatchBasedOnPercent = 3" )]
		int matchType,
		[ExcelArgument( "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string matchParam,
		[ExcelArgument( "Optional.  Number of Pay period in a year.  Default is 1." )]
		object? payPeriod = null,
		[ExcelArgument( "Optional.  Current Pay Period.  Default is 1." )]
		object? currPayPeriod = null )
	{
		var payPeriodArg = payPeriod.Check( nameof( payPeriod ), 1 );
		var currPayPeriodArg = currPayPeriod.Check( nameof( currPayPeriod ), 1 );

		var contributions = Savings.GetMinContributionsForMaxMatch( (CalculationMatchType)matchType, matchParam, payPeriodArg, currPayPeriodArg ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}
}