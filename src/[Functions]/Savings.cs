using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaSavings
{
	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns plan investment annual return based on Fund allocations and how Funds are mapped to assets classes." )]
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

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns plan asset class allocation based on Fund allocations and how Funds are mapped to assets classes.", IsMacroType = true )]
	public static double[] BTRAssetClassAllocation(
		[ExcelArgument( "Required.  Fund Table Name." )]
		string tableName,
		[ExcelArgument( "Required.  Plan type." )]
		int planType,
		[ExcelArgument( "Required.  Year of requested allocation." )]
		int year,
		[ExcelArgument( "Required.  Current fund allocations." )]
		double[ , ] fundAllocations,
		[KatExcelArgument( 
			AllowReference = true, 
			Description = "Optional.  Entered fund allocations (this will override current or future allocations).",
			Type = typeof( double[] )
		)]
		object? inputAllocations = null,
		[KatExcelArgument(
			Description = "Optional.  If provided allocation will be changed 100% to that target fund.",
			Type = typeof( string )
		)]
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

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns index of Target Fund based on retirement year and target fund group." )]
	public static string BTRTargetFundID(
		[ExcelArgument( "Fund Table Name." )]
		string tableName,
		[ExcelArgument( "Year of retirement." )]
		int yearRetirement,
		[ExcelArgument( "Target fund group ID." )]
		int targetFundGroup 
	) => Savings.TargetFundID( tableName, yearRetirement, targetFundGroup );
	
	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns an array with 3 values: balance at EOY, principle payments with investment return at EOY and principle payments with investment return at EOB." )]
	public static double[] BTRCalculateLoan(
		[ExcelArgument( "Required. Loan balance." )]
		double balance,
		[ExcelArgument( "Required. Number of Pay period in a year." )]
		int payPeriod,
		[KatExcelArgument( DisplayName = "paymentPerPayPeriod", Description = "Required. Payment per pay period." )]
		double paymentPPP,
		[ExcelArgument( "Required. Loan interest rate." )]
		double interestRate,
		[ExcelArgument( "Required. Calculate contributions as of the end of this month." )]
		int monthEnd,
		[ExcelArgument( "Required. Investment rate of return." )]
		double returnRate,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional. Starting pay period. Default to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? startPP = null,
		[KatExcelArgument(
			Description = "Optional. Use mid point contributions timing calculation method.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? midPointCont = null,
		[ExcelArgument( "Optional. Rounding, Defaulted to 0 decimals" )]
		int rounding = 0 )
	{
		var startPayPeriodArg = startPP.Check( nameof( startPP ), 1 );
		var midPointContributionArg = midPointCont.Check( nameof( midPointCont ), false );
		return Savings.CalculateLoan( balance, payPeriod, paymentPPP, interestRate, monthEnd, returnRate, startPayPeriodArg, midPointContributionArg, rounding );
	}

	[KatExcelFunction(
		CreateDebugFunction = true,
		Category = "Financial",
		Description = "DOC: Han, Cfgena replacement?  Returns 401k contributions/match.",
		Remarks = @"Both 'matchParam' and 'contributionParam' are a | delimited list of periods.  Each period is in the form of `M:P:P:P` where `M` is number of months for this period, and each `P` is a tier of a `,` seperated pair of decimal values.
The `matchType` and `contributionType` determine how each tier of numbers are processed.  Each tier's value pairs are described below:
1. `CalculationMatchType.MultiplierBasedOnPercent` - ContrbutionPercent, Multiplier
1. CalculationMatchType.MultiplierBasedOnDollars` - ContributionDollars, Multiplier
1. `CalculationMatchType.ERMatchPercentBasedOnPercent` - EEContributionPercent, ERContributionPercent
1. `CalculationMatchType.ERMatchDollarsBasedOnDollars` - ContributionDollars, ERContributionDollars
1. `CalculationContributionType.PercentBasedOnAge` - Age, ContributionPercent
1. `CalculationContributionType.PercentBasedOnService` - Service, ContributionPercent
1. `CalculationContributionType.PercentBasedOnAgePlusService` - AgePlusService, ContributionPercent
1. `CalculationContributionType.DollarsBasedOnAge` - Age, ContributionDollars
1. `CalculationContributionType.DollarsBasedOnService` - Service, ContributionDollars
1. `CalculationContributionType.DollarsBasedOnAgePlusService` - AgePlusService, ContributionDollars"
	)]
	public static double[,] BTRGetContributions(
		[KatExcelArgument( DisplayName = "matchType", Description = "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, MultiplierBasedOnDollars = 2, ERMatchPercentBasedOnPercent = 3, ERMatchDollarsBasedOnDollars = 4." )]
		int mTy,
		[KatExcelArgument( DisplayName = "isTrueUp", Description = "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool tUp,
		[KatExcelArgument( DisplayName = "isCatchUpMatch", Description = "Required.  Whether to provide match on catch-up contributions." )]
		bool mCU,
		[KatExcelArgument( DisplayName = "isPreTaxOverflowToAfterTax", Description = "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool pToA,
		[KatExcelArgument( DisplayName = "isMatchAfterTax", Description = "Required.  Whether to provide match on after-tax contributions." )]
		bool mA,
		[KatExcelArgument( DisplayName = "matchParam", Description = "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string mPm,
		[KatExcelArgument( DisplayName = "contributionType", Description = "Required.  The ContributionType to use for calculations.  PercentBasedOnAge = 1, PercentBasedOnService = 2, PercentBasedOnAgePlusService = 3." )]
		int cTy,
		[KatExcelArgument( DisplayName = "contributionParam", Description = "Required.  Contribution parameters.  See mPm for info." )]
		string cPm,
		[KatExcelArgument( DisplayName = "year", Description = "Required.  Calculation year." )]
		int y,
		[KatExcelArgument( DisplayName = "monthEnd", Description = "Required.  Calculate contributions as of the end of this month." )]
		double eom,
		[KatExcelArgument( DisplayName = "payPeriod", Description = "Required.  Number of Pay period in a year." )]
		int pyPd,
		[KatExcelArgument( DisplayName = "rateOfPay", Description = "Required.  Annual Pay rate as of start pay period." )]
		double rPy,
		[KatExcelArgument( DisplayName = "payPeriodWhenPayIncreases", Description = "Required.  Pay period when pay increases." )]
		int pyInPd,
		[KatExcelArgument( DisplayName = "ratePayIncrease", Description = "Required.  Pay increase rate." )]
		double rPyIn,
		[KatExcelArgument( DisplayName = "rateOfInflation", Description = "Required.  Inflation rate (used to project limits)." )]
		double rInf,
		[KatExcelArgument( DisplayName = "rateOfReturn", Description = "Required.  Investment rate of return." )]
		double ror,
		[KatExcelArgument( DisplayName = "preTaxPercentage", Description = "Optional.  Pre-tax contribution as a % of pay." )]
		double pPt = 0d,
		[KatExcelArgument( DisplayName = "preTaxFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double pF = 0d,
		[KatExcelArgument( DisplayName = "rothPercentage", Description = "Optional.  Roth contribution as a % of pay." )]
		double roPt = 0d,
		[KatExcelArgument( DisplayName = "rothFlatPerPayPeriod", Description = "Optional.  Flat $ Roth contribution amount per pay period." )]
		double roF = 0d,
		[KatExcelArgument( DisplayName = "afterTaxPercentage", Description = "Optional.  After-tax contribution as a % of pay." )]
		double aPt = 0d,
		[KatExcelArgument( DisplayName = "afterTaxFlatPerPayPeriod", Description = "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double aF = 0d,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional.  Starting Pay period.  Defaults to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? sPP = null,
		[KatExcelArgument( DisplayName = "ytdPay", Description = "Optional.  YTD savings pay." )]
		double yPy = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTax", Description = "Optional.  YTD pre-tax contributions." )]
		double yP = 0d,
		[KatExcelArgument( DisplayName = "ytdRoth", Description = "Optional.  YTD Roth contributions." )]
		double yRo = 0d,
		[KatExcelArgument( DisplayName = "ytdAfterTax", Description = "Optional.  YTD after-tax contributions." )]
		double yA = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTaxCatchUp", Description = "Optional.  YTD pre-tax catch-up contributions." )]
		double yPCU = 0d,
		[KatExcelArgument( DisplayName = "ytdRothCatchUp", Description = "Optional.  YTD Roth catch-up contributions." )]
		double yRoCU = 0d,
		[KatExcelArgument( DisplayName = "ytdErMatch", Description = "Optional.  YTD employer match contributions." )]
		double yErM = 0d,
		[KatExcelArgument( DisplayName = "ytdErContribution", Description = "Optional.  YTD employer contributions." )]
		double yErC = 0d,
		[KatExcelArgument( DisplayName = "ageBOY", Description = "Optional.  Age at BOY." )]
		double ageBY = 0d,
		[KatExcelArgument( DisplayName = "svcBOY", Description = "Optional.  Service at BOY." )]
		double svcBY = 0d,
		[KatExcelArgument( DisplayName = "erContributionAnnualLimit", Description = "Optional.  Employer contribution annual limit. Defaults to 0, which means unlimited." )]
		double erCAL = 0d,
		[KatExcelArgument( DisplayName = "isRetirementYear", Description = "Optional.  DOC: Han, Is this retirement year?" )]
		bool retY = false,
		[KatExcelArgument( DisplayName = "matchFrequency", Description = "Optional.  Frequency of match in a year." )]
		int mFr = 0,
		[KatExcelArgument( DisplayName = "erContributionFrequency", Description = "Optional.  Frequency of ER Contribution in a year." )]
		int erCFr = 0,
		[KatExcelArgument( 
			DisplayName = "preTaxRothPayIsLimited", 
			Description = "Optional.  Employee Pretax/Roth & Catchup Contributions is based on limited pay.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? pRoPyL = null,
		[KatExcelArgument( DisplayName = "midPointContribution", Description = "Optional.  Use mid point contributions timing calculation method." )]
		bool midC = false,
		[KatExcelArgument( DisplayName = "noLimit", Description = "Optional.  Don't apply IRS limit." )]
		bool noL = false,
		[KatExcelArgument( DisplayName = "increaseMonth", Description = "Optional.  Pretax auto increase timing (month): Enter 4 if increase happens on 4/1." )]
		int inMon = 0,
		[KatExcelArgument( DisplayName = "increaseFrequency", Description = "Optional.  Pretax auto increase frequency per year." )]
		int inFr = 0,
		[KatExcelArgument( DisplayName = "increasePercentage", Description = "Optional.  Pretax auto increase percentage." )]
		double inPt = 0,
		[KatExcelArgument( DisplayName = "rateMax", Description = "Optional.  Pretax maximum rate." )]
		double rMax = 0d,
		[KatExcelArgument( DisplayName = "limitPay", Description = "Optional.  Pay Limit." )]
		double lPy = 0d,
		[KatExcelArgument( DisplayName = "limitPreTax", Description = "Optional.  Pretax limit." )]
		double lP = 0d,
		[KatExcelArgument( DisplayName = "limitContribution", Description = "Optional.  Contribution limit." )]
		double lC = 0d,
		[KatExcelArgument( DisplayName = "isOverflowToNonQual", Description = "Optional.  Overflow to non qualified plan." )]
		bool toNQ = false,
		[KatExcelArgument( DisplayName = "limitAfterTax", Description = "Optional.  Aftertax limit." )]
		double lA = 0d,
		[KatExcelArgument( 
			DisplayName = "isOverflowToCatchUp", 
			Description = "Optional.  Overflow from preTax to catchup.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? toCU = null,
		[KatExcelArgument( DisplayName = "preTaxCatchUpPercentage", Description = "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double pCUPt = 0d,
		[KatExcelArgument( DisplayName = "preTaxCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double pCUF = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpPercentage", Description = "Optional.  Roth catchup contribution as a % of pay." )]
		double roCUPt = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double roCUF = 0d,
		[KatExcelArgument( 
			DisplayName = "matchPayIsLimited", 
			Description = "Optional.  Employer match is based on limited pay.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? mPyL = null
	)
	{
		var startPayPeriodArg = sPP.Check( nameof( sPP ), 1 );
		//var isRetirementArg = retirement.Check( nameof( retirement ), false );
		var preTaxRothPayIsLimitedArg = pRoPyL.Check( nameof( pRoPyL ), true );
		//var midPointContributionArg = midPointCont.Check( nameof( midPointCont ), false );
		//var noLimitArg = noLim.Check( nameof( noLim ), false );
		var isOverflowToCatchupArg = toCU.Check( nameof( toCU ), true );
		var matchPayIsLimitedArg = mPyL.Check( nameof( mPyL ), true );

		var contributions = Savings.GetContributions( ResultType.ArrayOfContributions, (CalculationMatchType)mTy, tUp, mCU, pToA, mA, mPm,
			(CalculationContributionType)cTy, cPm, y, eom, pyPd, rPy, pyInPd, rPyIn, rInf, ror,
			pPt, pF, roPt, roF, aPt, aF, startPayPeriodArg,
			yPy, yP, yRo, yA, yPCU, yRoCU, yErM, yErC, ageBY, svcBY, erCAL, retY, mFr,
			erCFr, preTaxRothPayIsLimitedArg, midC, noL, inMon, inFr, inPt, rMax,
			lPy, lP, lC, toNQ, lA, isOverflowToCatchupArg, pCUPt, pCUF, roCUPt, roCUF, matchPayIsLimitedArg ) ?? 
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns 401(k) contributions/match." )]
	public static double[,] BTRGet401kContributions(
		[KatExcelArgument( DisplayName = "matchType", Description = "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, MultiplierBasedOnDollars = 2, ERMatchPercentBasedOnPercent = 3, ERMatchDollarsBasedOnDollars = 4." )]
		int mTy,
		[KatExcelArgument( DisplayName = "isTrueUp", Description = "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool tUp,
		[KatExcelArgument( DisplayName = "isCatchUpMatch", Description = "Required.  Whether to provide match on catch-up contributions." )]
		bool mCU,
		[KatExcelArgument( DisplayName = "isPreTaxOverflowToAfterTax", Description = "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool pToA,
		[KatExcelArgument( DisplayName = "isAfterTaxMatch", Description = "Required.  Whether to provide match on after-tax contributions." )]
		bool mA,
		[KatExcelArgument( DisplayName = "matchParam", Description = "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string mPm,
		[KatExcelArgument( DisplayName = "year", Description = "Required.  Calculation year." )]
		int y,
		[KatExcelArgument( DisplayName = "monthEnd", Description = "Required.  Calculate contributions as of the end of this month." )]
		double eom,
		[KatExcelArgument( DisplayName = "payPeriod", Description = "Required.  Number of Pay period in a year." )]
		int pyPd,
		[KatExcelArgument( DisplayName = "rateOfPay", Description = "Required.  Annual Pay rate as of start pay period." )]
		double rPy,
		[KatExcelArgument( DisplayName = "payPeriodWhenPayIncreases", Description = "Required.  Pay period when pay increases." )]
		int pyInPd,
		[KatExcelArgument( DisplayName = "ratePayIncrease", Description = "Required.  Pay increase rate." )]
		double rPyIn,
		[KatExcelArgument( DisplayName = "rateOfInflation", Description = "Required.  Inflation rate (used to project limits)." )]
		double rInfl,
		[KatExcelArgument( DisplayName = "rateOfReturn", Description = "Required.  Investment rate of return." )]
		double ror,
		[KatExcelArgument( DisplayName = "preTaxPercentage", Description = "Optional.  Pre-tax contribution as a % of pay." )]
		double pPt = 0d,
		[KatExcelArgument( DisplayName = "preTaxFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double pF = 0d,
		[KatExcelArgument( DisplayName = "rothPercentage", Description = "Optional.  Roth contribution as a % of pay." )]
		double roPt = 0d,
		[KatExcelArgument( DisplayName = "rothFlatPerPayPeriod", Description = "Optional.  Flat $ Roth contribution amount per pay period." )]
		double roF = 0d,
		[KatExcelArgument( DisplayName = "afterTaxPercentage", Description = "Optional.  After-tax contribution as a % of pay." )]
		double aPt = 0d,
		[KatExcelArgument( DisplayName = "afterTaxFlatPerPayPeriod", Description = "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double aF = 0d,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional.  Starting Pay period.  Defaults to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? sPP = null,
		[KatExcelArgument( DisplayName = "ytdPay", Description = "Optional.  YTD savings pay." )]
		double yPy = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTax", Description = "Optional.  YTD pre-tax contributions." )]
		double yP = 0d,
		[KatExcelArgument( DisplayName = "ytdRoth", Description = "Optional.  YTD Roth contributions." )]
		double yRo = 0d,
		[KatExcelArgument( DisplayName = "ytdAfterTax", Description = "Optional.  YTD after-tax contributions." )]
		double yA = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTaxCatchUp", Description = "Optional.  YTD pre-tax catch-up contributions." )]
		double yPCU = 0d,
		[KatExcelArgument( DisplayName = "ytdRothCatchUp", Description = "Optional.  YTD Roth catch-up contributions." )]
		double yRoCU = 0d,
		[KatExcelArgument( DisplayName = "ytdErMatch", Description = "Optional.  YTD employer match contributions." )]
		double yErM = 0d,
		[KatExcelArgument( DisplayName = "ageBOY", Description = "Optional.  Age at BOY." )]
		double ageBY = 0d,
		[KatExcelArgument( DisplayName = "isRetirementYear", Description = "Optional.  Is this retirement year?" )]
		bool retY = false,
		[KatExcelArgument( DisplayName = "matchFrequency", Description = "Optional.  Frequency of match in a year." )]
		int mFr = 0,
		[KatExcelArgument( 
			DisplayName = "preTaxRothPayIsLimited", 
			Description = "Optional.  Employee Pretax/Roth & Catchup Contributions is based on limited pay.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? pRoPyL = null,
		[KatExcelArgument( DisplayName = "midPointContribution", Description = "Optional.  Use mid point contributions timing calculation method." )]
		bool midC = false,
		[KatExcelArgument( DisplayName = "noLimit", Description = "Optional.  Don't apply IRS limit." )]
		bool noL = false,
		[KatExcelArgument( DisplayName = "increaseMonth", Description = "Optional.  Pretax auto increase timing (month): Enter 4 if increase happens on 4/1." )]
		int incMon = 0,
		[KatExcelArgument( DisplayName = "increaseFrequency", Description = "Optional.  Pretax auto increase frequency per year." )]
		int inFr = 0,
		[KatExcelArgument( DisplayName = "increasePercentage", Description = "Optional.  Pretax auto increase percentage." )]
		double inPt = 0,
		[KatExcelArgument( DisplayName = "rateMax", Description = "Optional.  Pretax maximum rate." )]
		double rMax = 0d,
		[KatExcelArgument( DisplayName = "limitPay", Description = "Optional.  Pay Limit." )]
		double lPy = 0d,
		[KatExcelArgument( DisplayName = "limitPreTax", Description = "Optional.  Pretax limit." )]
		double lP = 0d,
		[KatExcelArgument( DisplayName = "limitContribution", Description = "Optional.  Contribution limit." )]
		double lC = 0d,
		[KatExcelArgument( DisplayName = "isOverflowToNonQual", Description = "Optional.  Overflow to non qualified plan." )]
		bool toNQ = false,
		[KatExcelArgument( DisplayName = "limitAfterTax", Description = "Optional.  Aftertax limit." )]
		double lA = 0d,
		[KatExcelArgument( 
			DisplayName = "isOverFlowToCatchUp", 
			Description = "Optional.  Overflow from preTax to catchup.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? toCU = null,
		[KatExcelArgument( DisplayName = "preTaxCatchUpPercentage", Description = "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double pCUPt = 0d,
		[KatExcelArgument( DisplayName = "preTaxCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double pCUF = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpPercentage", Description = "Optional.  Roth catchup contribution as a % of pay." )]
		double roCUPt = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double roCUF = 0d,
		[KatExcelArgument( 
			DisplayName = "matchPayIsLimited", 
			Description = "Optional.  Employer match is based on limited pay.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? mPyL = null
	)
	{
		var startPayPeriodArg = sPP.Check( nameof( sPP ), 1 );
		//var isRetirementArg = retirement.Check( nameof( retirement ), false );
		var preTaxRothPayIsLimitedArg = pRoPyL.Check( nameof( pRoPyL ), true );
		//var midPointContributionArg = midPointCont.Check( nameof( midPointCont ), false );
		//var noLimitArg = noLim.Check( nameof( noLim ), false );
		var isOverflowToCatchupArg = toCU.Check( nameof( toCU ), true );
		var matchPayIsLimitedArg = mPyL.Check( nameof( mPyL ), true );

		var contributions = Savings.GetContributions( ResultType.ArrayOf401kContributions, (CalculationMatchType)mTy, tUp, mCU, pToA, mA, mPm,
			null, "", y, eom, pyPd, rPy, pyInPd, rPyIn, rInfl, ror,
			pPt, pF, roPt, roF, aPt, aF, startPayPeriodArg,
			yPy, yP, yRo, yA, yPCU, yRoCU, yErM, 0d, ageBY, 0d, 0d, retY, mFr, 0,
			preTaxRothPayIsLimitedArg, midC, noL, incMon, inFr, inPt, rMax,
			lPy, lP, lC, toNQ, lA, isOverflowToCatchupArg, pCUPt, pCUF, roCUPt, roCUF, matchPayIsLimitedArg ) ?? 
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer match % based on match parameters." )]
	public static double[] BTRGetMatchPct(
		[KatExcelArgument( DisplayName = "matchType", Description = "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, ERMatchBasedOnPercent = 3" )]
		int mType,
		[KatExcelArgument( DisplayName = "isTrueUp", Description = "Required.  Whether to credit true-up math at the end of the year when employee hits contribution limit." )]
		bool trueUp,
		[KatExcelArgument( DisplayName = "isMatchCatchUp", Description = "Required.  Whether to provide match on catch-up contributions." )]
		bool mCUp,
		[KatExcelArgument( DisplayName = "isPreTaxOverflowToAfterTax", Description = "Required.  Whether to allow pre-tax contributions over limit to overflow to after-tax contributions." )]
		bool preToAft,
		[KatExcelArgument( DisplayName = "isAfterTaxMatch", Description = "Required.  Whether to provide match on after-tax contributions." )]
		bool mAft,
		[KatExcelArgument( DisplayName = "matchParam", Description = "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string mParam,
		[KatExcelArgument( DisplayName = "year", Description = "Required.  Calculation year." )]
		int yr,
		[ExcelArgument( "Required.  Calculate match as of the end of this month." )]
		double monthEnd,
		[KatExcelArgument( DisplayName = "payPeriod", Description = "Required.  Number of Pay period in a year." )]
		int payPd,
		[KatExcelArgument( DisplayName = "rateOfPay", Description = "Required.  Annual Pay rate as of start pay period." )]
		double rPay,
		[KatExcelArgument( DisplayName = "payPeriodWhenPayIncreases", Description = "Required.  Pay period when pay increases." )]
		int payIncPd,
		[KatExcelArgument( DisplayName = "rateOfPayIncrease", Description = "Required.  Pay increase rate." )]
		double rPayInc,
		[KatExcelArgument( DisplayName = "rateOfInflation", Description = "Required.  Inflation rate (used to project limits)." )]
		double rInfl,
		[KatExcelArgument( DisplayName = "preTaxPercentage", Description = "Optional.  Pre-tax contribution as a % of pay." )]
		double prePct = 0d,
		[KatExcelArgument( DisplayName = "preTaxFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax contribution amount per pay period." )]
		double preFPPP = 0d,
		[KatExcelArgument( DisplayName = "rothPercentage", Description = "Optional.  Roth contribution as a % of pay." )]
		double rothPct = 0d,
		[KatExcelArgument( DisplayName = "rothFlatPerPayPeriod", Description = "Optional.  Flat $ Roth contribution amount per pay period." )]
		double rothFPPP = 0d,
		[KatExcelArgument( DisplayName = "afterTaxPercentage", Description = "Optional.  After-tax contribution as a % of pay." )]
		double aftPct = 0d,
		[KatExcelArgument( DisplayName = "afterTaxFlatPerPayPeriod", Description = "Optional.  Flat $ after-tax contribution amount per pay period." )]
		double aftFPPP = 0d,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional.  Starting Pay period.  Defaults to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? stPP = null,
		[KatExcelArgument( DisplayName = "ytdPay", Description = "Optional.  YTD savings pay." )]
		double yrPay = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTax", Description = "Optional.  YTD pre-tax contributions." )]
		double yrPre = 0d,
		[KatExcelArgument( DisplayName = "ytdRoth", Description = "Optional.  YTD Roth contributions." )]
		double yrRoth = 0d,
		[KatExcelArgument( DisplayName = "ytdAfterTax", Description = "Optional.  YTD after-tax contributions." )]
		double yrAft = 0d,
		[KatExcelArgument( DisplayName = "ytdPreTaxCatchUp", Description = "Optional.  YTD pre-tax catch-up contributions." )]
		double yrPreCUp = 0d,
		[KatExcelArgument( DisplayName = "ytdRothCatchUp", Description = "Optional.  YTD Roth catch-up contributions." )]
		double yrRothCUp = 0d,
		[KatExcelArgument( DisplayName = "ytdErMatch", Description = "Optional.  YTD employer match contributions." )]
		double yrErM = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[KatExcelArgument( DisplayName = "matchFrequency", Description = "Optional.  Frequency of match in a year." )]
		int mFreq = 0,
		[KatExcelArgument( DisplayName = "limitAfterTax", Description = "Optional.  Aftertax limit." )]
		double lAft = 0d,
		[KatExcelArgument( 
			DisplayName = "isOverflowToCatchUp", 
			Description = "Optional.  Overflow from preTax to catchup.",
			Type = typeof( bool ),
			Default = "true"
		 )]
		object? toCUp = null,
		[KatExcelArgument( DisplayName = "preTaxCatchUpPercentage", Description = "Optional.  Pre-tax catchup contribution as a % of pay." )]
		double preCUpPct = 0d,
		[KatExcelArgument( DisplayName = "preTaxCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ pre-tax catchup contribution amount per pay period." )]
		double preCUpFPPP = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpPercentage", Description = "Optional.  Roth catchup contribution as a % of pay." )]
		double rothCUpPct = 0d,
		[KatExcelArgument( DisplayName = "rothCatchUpFlatPerPayPeriod", Description = "Optional.  Flat $ Roth catchup contribution amount per pay period." )]
		double rothCUpFPPP = 0d 
	)
	{
		var startPayPeriodArg = stPP.Check( nameof( stPP ), 1 );
		var isOverflowToCatchupArg = toCUp.Check( nameof( toCUp ), true );

		var contributions = Savings.GetContributions( ResultType.MatchPct, (CalculationMatchType)mType, trueUp, mCUp, preToAft, mAft, mParam,
			null, "", yr, monthEnd, payPd, rPay, payIncPd, rPayInc, rInfl, 0d,
			prePct, preFPPP, rothPct, rothFPPP, aftPct, aftFPPP, startPayPeriodArg,
			yrPay, yrPre, yrRoth, yrAft, yrPreCUp, yrRothCUp, yrErM, 0d, ageBOY, 0d, 0d, false, mFreq, 0,
			true, false, false, 0, 0, 0d, 0d, 0d, 0d, 0d, false, lAft, isOverflowToCatchupArg, preCUpPct, preCUpFPPP, rothCUpPct, rothCUpFPPP ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN ).GetRow( 0 );
		}

		return contributions.GetRow( 0 );
	}

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer contributions." )]
	public static double[,] BTRGetERContribution(
		[KatExcelArgument( DisplayName = "contributionType", Description = "Required.  The ContributionType to use for calculations.  PercentBasedOnAge = 1, PercentBasedOnService = 2, PercentBasedOnAgePlusService = 3." )]
		int contType,
		[KatExcelArgument( DisplayName = "contributionParam", Description = "Required.  Contribution parameters.  See matchParm for info." )]
		string contParam,
		[ExcelArgument( "Required.  Calculation year." )]
		int year,
		[ExcelArgument( "Required.  Calculate contributions as of the end of this month." )]
		double monthEnd,
		[ExcelArgument( "Required.  Number of Pay period in a year." )]
		int payPeriod,
		[KatExcelArgument( DisplayName = "rateOfPay", Description = "Required.  Annual Pay rate as of start pay period." )]
		double ratePay,
		[KatExcelArgument( DisplayName = "payPeriodWhenPayIncreases", Description = "Required.  Pay period when pay increases." )]
		int payIncPeriod,
		[KatExcelArgument( DisplayName = "rateOfPayIncrease", Description = "Required.  Pay increase rate." )]
		double payIncRate,
		[KatExcelArgument( DisplayName = "rateOfInflation", Description = "Required.  Inflation rate (used to project limits)." )]
		double rateInfl,
		[KatExcelArgument( DisplayName = "rateOfReturn", Description = "Required.  Investment rate of return." )]
		double rateReturn,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional.  Starting Pay period.  Defaults to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? startPP = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[KatExcelArgument( DisplayName = "ytdErContribution", Description = "Optional.  YTD employer contributions." )]
		double ytdErCont = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Service at BOY." )]
		double svcBOY = 0d,
		[KatExcelArgument( DisplayName = "erContributionAnnualLimit", Description = "Optional.  Employer contribution annual limit. Defaults to unlimited." )]
		double erContAnnualLimit = 0d,
		[KatExcelArgument( DisplayName = "erContributionFrequency", Description = "Optional.  Frequency of ER Contribution in a year." )]
		int erContFreq = 0,
		[KatExcelArgument( 
			DisplayName = "midPointContribution", 
			Description = "Optional.  Use mid point contributions timing calculation method.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? midPointCont = null,
		[KatExcelArgument( 
			DisplayName = "noLimit", 
			Description = "Optional.  Don't apply IRS limit.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? noLim = null,
		[KatExcelArgument( DisplayName = "limitPay", Description = "Optional.  Pay Limit." )]
		double limPay = 0d,
		[ KatExcelArgument( DisplayName = "limitContribution", Description = "Optional.  Contribution limit." ) ]
		double limitCont = 0d,
		[ KatExcelArgument( DisplayName = "isOverflowToNonQual", Description = "Optional.  Overflow to non qualified plan." ) ]
		bool overflowToNonQual = false 
	)
	{
		var startPayPeriodArg = startPP.Check( nameof( startPP ), 1 );
		var midPointContributionArg = midPointCont.Check( nameof( midPointCont ), false );
		var noLimitArg = noLim.Check( nameof( noLim ), false );

		var contributions = Savings.GetContributions( ResultType.ContributionPct, null, false, false, false, false, "", (CalculationContributionType)contType, contParam,
			year, monthEnd, payPeriod, ratePay, payIncPeriod, payIncRate, rateInfl, rateReturn, 0d, 0d, 0d, 0d, 0d, 0d,
			startPayPeriodArg, ytdPay, 0d, 0d, 0d, 0d, 0d, 0d, ytdErCont, ageBOY, svcBOY, erContAnnualLimit, false, 0, erContFreq, true, midPointContributionArg, noLimitArg,
			0, 0, 0d, 0d, limPay, 0d, limitCont, overflowToNonQual ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN );
		}

		return contributions;
	}

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns employer contribution % based on contribution parameters." )]
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
		[KatExcelArgument( DisplayName = "rateOfPay", Description = "Required.  Annual Pay rate as of start pay period." )]
		double payRate,
		[KatExcelArgument( DisplayName = "payPeriodWhenPayIncreases", Description = "Required.  Pay period when pay increases." )]
		int payIncPeriod,
		[ExcelArgument( "Required.  Pay increase rate." )]
		double rateOfPayIncrease,
		[KatExcelArgument( DisplayName = "rateOfInflation", Description = "Required.  Inflation rate (used to project limits)." )]
		double inflRate,
		[KatExcelArgument( 
			DisplayName = "startPayPeriod", 
			Description = "Optional.  Starting pay period.  Defaults to 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? startPP = null,
		[ExcelArgument( "Optional.  YTD savings pay." )]
		double ytdPay = 0d,
		[KatExcelArgument( DisplayName = "ytdErContribution", Description = "Optional.  YTD employer contributions." )]
		double ytdErCont = 0d,
		[ExcelArgument( "Optional.  Age at BOY." )]
		double ageBOY = 0d,
		[ExcelArgument( "Optional.  Service at BOY." )]
		double svcBOY = 0d,
		[ExcelArgument( "Optional.  Employer contribution annual limit. Defaults to unlimited." )]
		double erContributionAnnualLimit = 0d,
		[KatExcelArgument( DisplayName = "erContributionFrequency", Description = "Optional.  Frequency of ER Contribution in a year." )]
		int erContFreq = 0 
	)
	{
		var startPayPeriodArg = startPP.Check( nameof( startPP ), 1 );

		var contributions = Savings.GetContributions( ResultType.ContributionPct, null, false, false, false, false, "", (CalculationContributionType)contributionType, contributionParam,
			year, monthEnd, payPeriod, payRate, payIncPeriod, rateOfPayIncrease, inflRate, 0d, 0d, 0d, 0d, 0d, 0d, 0d,
			startPayPeriodArg, ytdPay, 0d, 0d, 0d, 0d, 0d, 0d, ytdErCont, ageBOY, svcBOY, erContributionAnnualLimit, false, 0, erContFreq ) ??
			throw new NullReferenceException( "Contributions is null." );

		// Until I get possible workaround from SpreadsheetGear, I want add-in and RBLe to behave same
		if ( contributions.Any( r => double.IsNaN( r ) ) )
		{
			return contributions.Select( r => double.NaN ).GetRow( 0 );
		}

		return contributions.GetRow( 0 );
	}

	[KatExcelFunction( CreateDebugFunction = true, Category = "Financial", Description = "DOC: Han, Cfgena replacement?  Returns minimum employee contribution to get max mactch based on match parameters." )]
	public static double[,] BTRGetMinContributionsForMaxMatch(
		[ExcelArgument( "Required.  The MatchType to use for calculations.  MultiplierBasedOnPercent = 1, ERMatchBasedOnPercent = 3" )]
		int matchType,
		[ExcelArgument( "Required.  | delimited list of periods.  Each period is in the form of M:P:P where M is number of months for this period, and each P is a tier of a , seperated pair of decimal values." )]
		string matchParam,
		[KatExcelArgument(
			Description = "Optional.  Number of Pay period in a year.  Default is 1.",
			Type = typeof( int ),
			Default = "1"
		)]
		object? payPeriod = null,
		[KatExcelArgument(
			Description = "Optional.  Current Pay Period.  Default is 1.",
			Type = typeof( int ),
			Default = "1"
		)]
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