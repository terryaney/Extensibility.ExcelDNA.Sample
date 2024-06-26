using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaAnnuity
{
	[ExcelFunction( Category = "Financial", Description = "List all mortality tables available for BTR Annuity functions" )]
	public static string BTRMortalityTableNames() => string.Join( ", ", Annuity.MortalityTableNames() );
	
	[KatExcelFunction( 
		CreateDebugFunction = true,
		Category = "Financial", 
		Description = "Replacement function for the Cfgena.xla!PPASingleLife() function.  Returns a decimal value equal to the selected single life annuity factor by the PPA method.",
		Remarks = "If you defer a temporary or certain annuity to an age earlier than the individual's current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity."
	)]
	public static double BTRPPASingleLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The interest rate to use for years 0 - 4." )]
		double intRate1,
		[ExcelArgument( "Required.  The interest rate to use for years 5 - 19." )]
		double intRate2,
		[ExcelArgument( "Required.  The interest rate to use for years 20 and after." )]
		double intRate3,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemp = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemp' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemp = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[KatExcelArgument( 
			Description = "The age or years for certain period with no effect prior to payment start time.  Certain period is: when howTemp=A and whenTemp < whenDefer, =0, else when howCertain=A, =value - whenDefer, otherwise =value.  Default is 0.",
			Summary = "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemp' is A and 'whenTemp' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0."
		)]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortTableAdj = 0,
		[KatExcelArgument( 
			Description = "A multiplier to adjust the mortality rates (qx). Entering 0 or blank will not change the rates. Any other value is a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." 
		)]
		object? mortSizeAdj = null,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[ExcelArgument( "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0
		// [ExcelArgument( "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely." )]
		// object? dynImpStopYr = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTempArg = howTemp.Check( nameof( howTemp ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );

		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var mortSizeAdjArg = mortSizeAdj.Check( nameof( mortSizeAdj ), 1d );

		return Annuity.PPASingleLife( 
			mortTable, intRate1, intRate2, intRate3, age, 
			howDeferArg, whenDefer, howTempArg, whenTemp, howCertainArg, whenCertain, 
			pmtsPerYrArg,
			mortTableAdj, mortSizeAdjArg, mortImp, impEffYrArg, yob 
		);
	}

	[KatExcelFunction( 
		CreateDebugFunction = true, 
		Category = "Financial", 
		Description = "Replacement function for the Cfgena.xla!SingleLife() function.  Returns a decimal value equal to the selected single life annuity factor.",
		Remarks = "If you defer a temporary or certain annuity to an age earlier than the individual's current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity."
	)]
	public static double BTRSingleLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double intRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( Name = "howDefer", Description = "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( Name = "howTemp", Description = "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemp = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemp' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemp = 0,
		[ExcelArgument( Name = "howCertain", Description = "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[KatExcelArgument( 
			Description = "The age or years for certain period with no effect prior to payment start time.  Certain period is: when howTemp=A and whenTemp < whenDefer, =0, else when howCertain=A, =value - whenDefer, otherwise =value.  Default is 0.",
			Summary = "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemp' is A and 'whenTemp' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0."
		)]
		double whenCertain = 0,
		[ExcelArgument( Name = "pmtsPerYr", Description = "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortTableAdj = 0,
		[KatExcelArgument( 
			Description = "A multiplier to adjust the mortality rates (qx). Entering 0 or blank will not change the rates. Any other value is a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mortSizeAdj = null,
		[KatExcelArgument( DisplayName = "mortalityImprovement", Description = "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is blank or '0' then the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTempArg = howTemp.Check( nameof( howTemp ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mortSizeAdjArg = mortSizeAdj.Check( nameof( mortSizeAdj ), 1d );

		return Annuity.SingleLife( 
			mortTable, intRate, age, 
			howDeferArg, whenDefer, howTempArg, whenTemp, howCertainArg, whenCertain, 
			pmtsPerYrArg, mortTableAdj, mortSizeAdjArg, mortImp, impEffYrArg, yob, dynImpStopYrArg 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!SingleLifeDeferPBGC() function.  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeDeferPBGC(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double intRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "Interest for first 7 deferral years, 0.06 is 6%. (Not allowed to be less than 4%.)" )]
		object? intRate7 = null,
		[ExcelArgument( "Interest for next 8 deferral years, 0.05 is 5%. (Not allowed to be less than 4%.)" )]
		object? intRate8 = null,
		[ExcelArgument( "Interest for remaining deferral years, 0.04 is 4%. (Not allowed to be less than 4%.)" )]
		object? intRateR = null,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortTableAdj = 0
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var interestRate7Arg = intRate7.Check( nameof( intRate7 ), 0.04d );
		var interestRate8Arg = intRate8.Check( nameof( intRate8 ), 0.04d );
		var interestRateRArg = intRateR.Check( nameof( intRateR ), 0.04d );

		var howTemp = "Y";
		var whenTemp = 0d;
		var howCertain = "Y";
		var whenCertain = 0d;
		var mortImp = 0;
		var impEffYr = -1;
		var yob = 0;
		var dynImpStopYr = 9999;
		var mortSizeAdj = 1d;

		// set interest rate array
		var intRates = new double[ 6, 2 ];
		intRates[ 0, 0 ] = Math.Min( 7, ( howDeferArg == "Y" ) ? whenDefer : whenDefer - age );
		intRates[ 0, 1 ] = intRates[ 0, 0 ] == 0 ? intRate : interestRate7Arg;
		intRates[ 1, 0 ] = Math.Min( 8, ( howDeferArg == "Y" ? whenDefer : whenDefer - age ) - intRates[ 0, 0 ] );
		intRates[ 1, 1 ] = intRates[ 1, 0 ] == 0 ? intRate : interestRate8Arg;
		intRates[ 2, 0 ] = Math.Max( 0, ( howDeferArg == "Y" ? whenDefer : whenDefer - age ) - intRates[ 0, 0 ] - intRates[ 1, 0 ] );
		intRates[ 2, 1 ] = intRates[ 2, 0 ] == 0 ? intRate : interestRateRArg;
		intRates[ 3, 0 ] = 0;
		intRates[ 3, 1 ] = intRate;

		return Annuity.SingleLife( 
			mortTable, intRates, age, 
			howDeferArg, whenDefer, howTemp, whenTemp, howCertain, whenCertain, 
			pmtsPerYrArg, mortTableAdj, mortSizeAdj, mortImp, impEffYr, yob, dynImpStopYr 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!???() function (DOC: Han, which function?).  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeWithRates(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  A 6x2 array representing interest rates.  The first column refers to the period for which the corresponding interest rate in the second column applies." )]
		double[ , ] intRates,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemp = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemp' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemp = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[KatExcelArgument( 
			Description = "The age or years for certain period with no effect prior to payment start time.  Certain period is: when howTemp=A and whenTemp < whenDefer, =0, else when howCertain=A, =value - whenDefer, otherwise =value.  Default is 0.",
			Summary = "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemp' is A and 'whenTemp' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0."	
		)]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortTableAdj = 0,
		[KatExcelArgument( 
			Description = "A multiplier to adjust the mortality rates (qx). Entering 0 or blank will not change the rates. Any other value is a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mortSizeAdj = null,
		[KatExcelArgument( DisplayName = "mortalityImprovement", Description = "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is blank or '0' then the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTempArg = howTemp.Check( nameof( howTemp ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mortSizeAdjArg = mortSizeAdj.Check( nameof( mortSizeAdj ), 1d );

		return Annuity.SingleLife( 
			mortTable, intRates, age, 
			howDeferArg, whenDefer, howTempArg, whenTemp, howCertainArg, whenCertain, 
			pmtsPerYrArg, mortTableAdj, mortSizeAdjArg, mortImp, impEffYrArg, yob, dynImpStopYrArg 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!SingleLife() function.  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeComm(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double intRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( Name = "typeCF", Description = "Commutation Function (CF) type, please see CFGENA help for details." )]
		object? typeCF = null,
		[ExcelArgument( Name = "pmtsPerYr", Description = "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortTableAdj = 0,
		[KatExcelArgument( 
			Description = "A multiplier to adjust the mortality rates (qx). Entering 0 or blank will not change the rates. Any other value is a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mortSizeAdj = null,
		[KatExcelArgument( DisplayName = "mortalityImprovement", Description = "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is blank or '0' then the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null
	)
	{
		var typeCFArg = typeCF.Check( nameof( typeCF ), "" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mortSizeAdjArg = mortSizeAdj.Check( nameof( mortSizeAdj ), 1d );

		return Annuity.SingleLifeComm( 
			mortTable, intRate, age, typeCFArg, pmtsPerYrArg,
			mortTableAdj, mortSizeAdjArg, mortImp, impEffYrArg, yob, dynImpStopYrArg 
		);
	}

	[KatExcelFunction(
		CreateDebugFunction = true, 
		Category = "Financial",
		Description = "Replacement function for the Cfgena.xla!PPAJointLife() function.  Returns a decimal value equal to the selected joint life annuity factor by the PPA method.",
		Remarks = "<p>If you defer a temporary or certain annuity to an age earlier than the individual's current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity.</p><p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'intRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberAge' is less than 1 or greater than 120.</li><li>'spouseAge' is less than 1 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortImp' is not 0, 1, 2, 31 or 32.</li><li>'mortImp' is 1 or 31 and 'uniBlending' is 2 or 'yob' is 0 or 'spYob' is 0.</li><li>'continuingPercentage' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'uniBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'mTableAdj' or 'fTableAdj' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRPPAJointLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string spMortTable,
		[ExcelArgument( "Required.  The interest rate to use for years 0 - 4." )]
		double intRate1,
		[ExcelArgument( "Required.  The interest rate to use for years 5 - 19." )]
		double intRate2,
		[ExcelArgument( "Required.  The interest rate to use for years 20 and after." )]
		double intRate3,
		[ExcelArgument( "Required.  The current member age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "Required.  The spouse age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double spAge,
		[ExcelArgument( "The options for annuity factors. 'C' for contingent, 'S' for survivor, 'P' for popup, 'D' for double popup and 'J' for joint life factor only.  Default value is 'C'." )]
		object? annuityOption = null,
		[ExcelArgument( "The fraction of contingent/survivor amount to the primary amount; 0-1.  Default value is 0.5." )]
		object? jointFraction = null,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemp = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemp' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemp = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[KatExcelArgument( 
			Description = "The age or years for certain period with no effect prior to payment start time.  Certain period is: when howTemp=A and whenTemp < whenDefer, =0, else when howCertain=A, =value - whenDefer, otherwise =value.  Default is 0.",
			Summary = "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemp' is A and 'whenTemp' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0."
		)]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[KatExcelArgument( DisplayName = "maleTableAdjustment", Description = "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int mTableAdj = 0,
		[KatExcelArgument( DisplayName = "femailTableAdjustment", Description = "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int fTableAdj = 0,
		[KatExcelArgument( 
			DisplayName = "maleSizeAdjustment",
			Description = "A multiplier to adjust the male mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mSizeAdj = null,
		[KatExcelArgument( 
			DisplayName = "femaleSizeAdjustment",
			Description = "A multiplier to adjust the female mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? fSizeAdj = null,
		[KatExcelArgument( DisplayName = "maleUnisexPercentage", Description = "The unisex blending percentage applied to the male mortality table." )]
		object? mUniPct = null,
		[KatExcelArgument( DisplayName = "unisexBlending", Description = "The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int uniBlending = 0,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( DisplayName = "spouseYearOfBirth", Description = "The year of birth for the spouse if dynamic generational is enabled 'mortImp' is 1." )]
		int spYob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is blank or '0' then the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTempArg = howTemp.Check( nameof( howTemp ), "Y" );
		var annuityOptionArg = annuityOption.Check( nameof( annuityOption ), "C" );
		var jointFractionArg = jointFraction.Check( nameof( jointFraction ), 0.5 );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mSizeAdjArg = mSizeAdj.Check( nameof( mSizeAdj ), 1d );
		var fSizeAdjArg = fSizeAdj.Check( nameof( fSizeAdj ), 1d );
		var maleUnisexPercentageArg = mUniPct.Check( nameof( mUniPct ), 1d );

		return Annuity.PPAJointLife( 
			mortTable, spMortTable, intRate1, intRate2, intRate3, age, spAge, annuityOptionArg, jointFractionArg, howDeferArg, whenDefer, howTempArg, whenTemp, howCertainArg, whenCertain, pmtsPerYrArg,
			mTableAdj, fTableAdj, mSizeAdjArg, fSizeAdjArg, maleUnisexPercentageArg, uniBlending, mortImp, impEffYrArg, yob, spYob, dynImpStopYrArg 
		);
	}

	[KatExcelFunction(
		CreateDebugFunction = true, 
		Category = "Financial",
		Description = "Replacement function for the Cfgena.xla!JointLife() function.  Returns decimal value equal to the selected joint life annuity factor.",
		Remarks = "<p>If you defer a temporary or certain annuity to an age earlier than the individual's current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity.</p><p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'intRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberAge' is less than 1 or greater than 120.</li><li>'spouseAge' is less than 1 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortImp' is not 0, 1, 2, 31 or 32.</li><li>'mortImp' is 1 or 31 and 'uniBlending' is 2 or 'yob' is 0 or 'spYob' is 0.</li><li>'jointFraction' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'uniBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'mTableAdj' or 'fTableAdj' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRJointLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortTable,
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string spMortTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double intRate,
		[ExcelArgument( "Required.  The current member age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "Required.  The current spouse age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double spAge,
		[ExcelArgument( "The options for annuity factors. 'C' for contingent, 'S' for survivor, 'P' for popup, 'D' for double popup and 'J' for joint life factor only.  Default value is 'C'." )]
		object? annuityOption = null,
		[ExcelArgument( "The fraction of contingent/survivor amount to the primary amount; 0-1.  Default value is 0.5." )]
		object? jointFraction = null,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemp = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemp' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemp = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[KatExcelArgument( 
			Description = "The age or years for certain period with no effect prior to payment start time.  Certain period is: when howTemp=A and whenTemp < whenDefer, =0, else when howCertain=A, =value - whenDefer, otherwise =value.  Default is 0.",
			Summary = "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemp' is A and 'whenTemp' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0."
		)]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? pmtsPerYr = null,
		[KatExcelArgument( DisplayName = "maleTableAdjustment", Description = "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int mTableAdj = 0,
		[KatExcelArgument( DisplayName = "femaleTableAdjustment", Description = "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int fTableAdj = 0,
		[KatExcelArgument( 
			DisplayName = "maleSizeAdjustment",
			Description = "A multiplier to adjust the male mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mSizeAdj = null,
		[KatExcelArgument( 
			DisplayName = "femaleSizeAdjustment",
			Description = "A multiplier to adjust the female mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? fSizeAdj = null,
		[KatExcelArgument( DisplayName = "maleUnisexPercentage", Description = "The unisex blending percentage applied to the male mortality table." )]
		object? mUniPct = null,
		[KatExcelArgument( DisplayName = "unisexBlending", Description = "The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int uniBlending = 0,
		[KatExcelArgument( DisplayName = "mortalityImprovement", Description = "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( DisplayName = "spouseYearOfBirth", Description = "The year of birth for the spouse if dynamic generational is enabled 'mortImp' is 1." )]
		int spYob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is blank or '0' then the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTempArg = howTemp.Check( nameof( howTemp ), "Y" );
		var annuityOptionArg = annuityOption.Check( nameof( annuityOption ), "C" );
		var jointFractionArg = jointFraction.Check( nameof( jointFraction ), 0.5 );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var pmtsPerYrArg = pmtsPerYr.Check( nameof( pmtsPerYr ), 12 );
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mSizeAdjArg = mSizeAdj.Check( nameof( mSizeAdj ), 1d );
		var fSizeAdjArg = fSizeAdj.Check( nameof( fSizeAdj ), 1d );
		var maleUnisexPercentageArg = mUniPct.Check( nameof( mUniPct ), 1d );

		return Annuity.JointLife( 
			mortTable, spMortTable, intRate, age, spAge, annuityOptionArg, jointFractionArg, howDeferArg, whenDefer, howTempArg, whenTemp, howCertainArg, whenCertain, pmtsPerYrArg,
			mTableAdj, fTableAdj, mSizeAdjArg, fSizeAdjArg, maleUnisexPercentageArg, uniBlending, mortImp, impEffYrArg, yob, spYob, dynImpStopYrArg 
		);
	}

	[KatExcelFunction(
		CreateDebugFunction = true, 
		Category = "Financial",
		Description = "A replacement function for the Annbuck.xla!AnnBuck() function.  Returns a decimal value representing the selected life annuity factor.",
		Remarks = "<p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'intRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberSex' is not 1 or 2.</li><li>'spouseSex' is not 1 or 2.</li><li>'memberAge' is less than 0 or greater than 120.</li><li>'spouseAge' is less than 0 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortImp' is not 0, 1, 2, 11, 12, 21, 22, 31, 32, 41, 42, 51, or 52.</li><li>'mortImp' is 1, 11, 21, 31, 41, or 51 and 'uniBlending' is 2 or 'yob' is 0 or 'spYob' is 0.</li><li>'mortImp' is 21, 22, 51, or 52 and the static year of selected male or female mortality tables are less than 2014.</li><li>'continuingPercentage' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'paymentTiming' is not 1, 2, or 3.</li><li>'preRetirementMortality' is not 1, 2, 3, 4, 5, or 6.</li><li>'uniBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'mTableAdj' or 'fTableAdj' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRAnnBuck(
		[ExcelArgument( "Required.  A 6x2 array representing interest rates.  The first column refers to the period for which the corresponding interest rate in the second column applies." )]
		double[ , ] intRates,
		[ExcelArgument( "Required.  The 'PaymentTimingType' to use where 1 = Continuous Approximation, 2 = Beginning of the month, and 3 = End of the month." )]
		int pmtTiming,
		[ExcelArgument( "Required.  The member's age to calculate; must be a decimal number between 0 and 120, inclusive." )]
		double age,
		[ExcelArgument( "Required.  The spouse's age to calculate; must be a decimal number between 0 and 120, inclusive.  Use 0 when not calculating joint factors." )]
		double spAge,
		[ExcelArgument( "Required. The member's 'SexType' to use where 1 = Male and 2 = Female." )]
		int sex,
		[ExcelArgument( "Required. The spouse's 'SexType' to use where 1 = Male and 2 = Female." )]
		int spSex,
		[ExcelArgument( "Required.  The age that benefits commence. For immediate factors, enter 0 or value equal to 'memberAge'." )]
		double deferredAge,
		[ExcelArgument( "Required.  The number of years that payments are made. Enter 0 if there is no temporary period." )]
		double tempPeriod,
		[ExcelArgument( "Required.  The number of years for which payments are guaranteed upon death. Enter 0 if there is no guarantee period." )]
		double guaranteePeriod,
		[ExcelArgument( "Required.  The percentage that which payments will continue to the spouse upon death of the member." )]
		double continuingPercentage,
		[ExcelArgument( "Required.  The 'PreRetirementMortalityType' to use, where 1 = NoMortality, 2 = MemberRetirementAgeJointSurvivor, 3 = MemberOnly, 4 = MemberDeathNoGuarantee, 5 = MemberRetirementAgeNoGuarantee, and 6 = MemberRetirementAgeFull." )]
		int preRetirementMort,
		[KatExcelArgument( DisplayName = "maleMortalityTable", Description = "Required.  The number of the mortality table that you wish to use. For example, enter '214' for GAM83 Male." )]
		string mMortTable,
		[KatExcelArgument( DisplayName = "femaleMortalityTable", Description = "Required.  The number of the mortality table that you wish to use.  See 'mMortTable'." )]
		string fMortTable,
		[KatExcelArgument( DisplayName = "maleUnisexPercentage", Description = "Required.  The unisex blending percentage applied to the male mortality table." )]
		double mUniPct,
		[KatExcelArgument( DisplayName = "unisexBlending", Description = "Required.  The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int uniBlending,
		[KatExcelArgument( DisplayName = "maleTableAdjustment", Description = "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int mTableAdj = 0,
		[KatExcelArgument( DisplayName = "femaleTableAdjustment", Description = "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int fTableAdj = 0,
		[KatExcelArgument( 
			DisplayName = "mortalityImprovement",
			Description = "The 'MortalityImprovementType' to use.  See help link for allowed values.",
			Summary = "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 11 = DynamicScaleCPMA1, 12 = StaticScaleCPMA1, 21 = DynamicScaleCPMA, 22 = StaticScaleCPMA, 31 = DynamicScaleBB, 32 = StaticScaleBB, 41 = DynamicScaleCPMB1, 42 = StaticScaleCPMB1, 51 = DynamicScaleCPMB, and 52 = StaticScaleCPMB."
		)]
		int mortImp = 0,
		[KatExcelArgument( DisplayName = "improvementEffectiveYear", Description = "The year of calculation if 'mortImp' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? impEffYr = null,
		[KatExcelArgument( DisplayName = "memberYearOfBirth", Description = "The year of birth for the member if dynamic generational is enabled 'mortImp' is 1." )]
		int yob = 0,
		[KatExcelArgument( DisplayName = "spouseYearOfBirth", Description = "The year of birth for the spouse if dynamic generational is enabled 'mortImp' is 1." )]
		int spYob = 0,
		[KatExcelArgument( 
			DisplayName ="dynamicImprovementStopYear",
			Description = "The stop year if dynamic generational is enabled and you wish to stop the generational projection at a future year. This caps the projection factors exponent. If 0 or blank, then it is assumed that the improvements continue indefinitely.",
			Summary = "The stop year if dynamic generational is enabled ('mortImp' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a '0' then it is assumed that the improvements continue indefinitely."
		)]
		object? dynImpStopYr = null,
		[KatExcelArgument( 
			DisplayName = "maleSizeAdjustment",
			Description = "A multiplier to adjust the male mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? mSizeAdj = null,
		[KatExcelArgument( 
			DisplayName = "femaleSizeAdjustment",
			Description = "A multiplier to adjust the female mortality rates (qx). Entering 0 or blank will not change the rates. Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table.",
			Summary = "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table."
		)]
		object? fSizeAdj = null 
	)
	{
		var impEffYrArg = impEffYr.Check( nameof( impEffYr ), -1 );
		var dynImpStopYrArg = dynImpStopYr.Check( nameof( dynImpStopYr ), 9999 );
		var mSizeAdjArg = mSizeAdj.Check( nameof( mSizeAdj ), 1d );
		var fSizeAdjArg = fSizeAdj.Check( nameof( fSizeAdj ), 1d );

		return Annuity.CanadaAnnuity( intRates, pmtTiming, age, spAge, sex, spSex, deferredAge, tempPeriod, guaranteePeriod, continuingPercentage, preRetirementMort, mMortTable, fMortTable, mUniPct, uniBlending, mTableAdj, fTableAdj, mortImp, impEffYrArg, yob, spYob, dynImpStopYrArg, mSizeAdjArg, fSizeAdjArg );
	}
}