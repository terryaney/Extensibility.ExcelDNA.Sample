using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaAnnuity
{
	[ExcelFunction( Category = "Financial", Description = "List all mortality tables available for BTR Annuity functions" )]
	public static string BTRMortalityTableNames() => string.Join( ", ", Annuity.MortalityTableNames() );
	

	[DebugFunction]
	[ExcelFunctionDoc( 
		Category = "Financial", 
		Description = "Replacement function for the Cfgena.xla!PPASingleLife() function.  Returns a decimal value equal to the selected single life annuity factor by the PPA method.",
		Remarks = "If you defer a temporary or certain annuity to an age earlier than the individual’s current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity."
	)]
	public static double BTRPPASingleLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortalityTable,
		[ExcelArgument( "Required.  The interest rate to use for years 0 - 4." )]
		double interestRate1,
		[ExcelArgument( "Required.  The interest rate to use for years 5 - 19." )]
		double interestRate2,
		[ExcelArgument( "Required.  The interest rate to use for years 20 and after." )]
		double interestRate3,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemporary = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemporary' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemporary = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[ExcelArgument( "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemporary' is A and 'whenTemporary' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0." )]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortalityTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? mortalitySizeAdjustment = null,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0
		// [ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		// object? dynamicImprovementStopYear = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTemporaryArg = howTemporary.Check( nameof( howTemporary ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );

		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var mortalitySizeAdjustmentArg = mortalitySizeAdjustment.Check( nameof( mortalitySizeAdjustment ), 1d );

		return Annuity.PPASingleLife( 
			mortalityTable, interestRate1, interestRate2, interestRate3, age, 
			howDeferArg, whenDefer, howTemporaryArg, whenTemporary, howCertainArg, whenCertain, 
			paymentsPerYearArg,
			mortalityTableAdjustment, mortalitySizeAdjustmentArg, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth 
		);
	}

	[DebugFunction]
	[ExcelFunctionDoc( 
		Category = "Financial", 
		Description = "Replacement function for the Cfgena.xla!SingleLife() function.  Returns a decimal value equal to the selected single life annuity factor.",
		Remarks = "If you defer a temporary or certain annuity to an age earlier than the individual’s current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity."
	)]
	public static double BTRSingleLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortalityTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double interestRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( Name = "howDefer", Description = "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( Name = "howTemporary", Description = "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemporary = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemporary' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemporary = 0,
		[ExcelArgument( Name = "howCertain", Description = "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[ExcelArgument( "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemporary' is A and 'whenTemporary' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0." )]
		double whenCertain = 0,
		[ExcelArgument( Name = "paymentsPerYear", Description = "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortalityTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? mortalitySizeAdjustment = null,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTemporaryArg = howTemporary.Check( nameof( howTemporary ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var mortalitySizeAdjustmentArg = mortalitySizeAdjustment.Check( nameof( mortalitySizeAdjustment ), 1d );

		return Annuity.SingleLife( 
			mortalityTable, interestRate, age, 
			howDeferArg, whenDefer, howTemporaryArg, whenTemporary, howCertainArg, whenCertain, 
			paymentsPerYearArg, mortalityTableAdjustment, mortalitySizeAdjustmentArg, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, dynamicImprovementStopYearArg 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!SingleLifeDeferPBGC() function.  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeDeferPBGC(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortalityTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double interestRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "Interest for first 7 deferral years, 0.06 is 6%. (Not allowed to be less than 4%.)" )]
		object? interestRate7 = null,
		[ExcelArgument( "Interest for next 8 deferral years, 0.05 is 5%. (Not allowed to be less than 4%.)" )]
		object? interestRate8 = null,
		[ExcelArgument( "Interest for remaining deferral years, 0.04 is 4%. (Not allowed to be less than 4%.)" )]
		object? interestRateR = null,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortalityTableAdjustment = 0
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var interestRate7Arg = interestRate7.Check( nameof( interestRate7 ), 0.04d );
		var interestRate8Arg = interestRate8.Check( nameof( interestRate8 ), 0.04d );
		var interestRateRArg = interestRateR.Check( nameof( interestRateR ), 0.04d );

		var howTemporary = "Y";
		var whenTemporary = 0d;
		var howCertain = "Y";
		var whenCertain = 0d;
		var mortalityImprovement = 0;
		var improvementEffectiveYear = -1;
		var memberYearOfBirth = 0;
		var dynamicImprovementStopYear = 9999;
		var mortalitySizeAdjustment = 1d;

		// set interest rate array
		var interestRates = new double[ 6, 2 ];
		interestRates[ 0, 0 ] = Math.Min( 7, ( howDeferArg == "Y" ) ? whenDefer : whenDefer - age );
		interestRates[ 0, 1 ] = interestRates[ 0, 0 ] == 0 ? interestRate : interestRate7Arg;
		interestRates[ 1, 0 ] = Math.Min( 8, ( howDeferArg == "Y" ? whenDefer : whenDefer - age ) - interestRates[ 0, 0 ] );
		interestRates[ 1, 1 ] = interestRates[ 1, 0 ] == 0 ? interestRate : interestRate8Arg;
		interestRates[ 2, 0 ] = Math.Max( 0, ( howDeferArg == "Y" ? whenDefer : whenDefer - age ) - interestRates[ 0, 0 ] - interestRates[ 1, 0 ] );
		interestRates[ 2, 1 ] = interestRates[ 2, 0 ] == 0 ? interestRate : interestRateRArg;
		interestRates[ 3, 0 ] = 0;
		interestRates[ 3, 1 ] = interestRate;

		return Annuity.SingleLife( 
			mortalityTable, interestRates, age, 
			howDeferArg, whenDefer, howTemporary, whenTemporary, howCertain, whenCertain, 
			paymentsPerYearArg, mortalityTableAdjustment, mortalitySizeAdjustment, mortalityImprovement, improvementEffectiveYear, memberYearOfBirth, dynamicImprovementStopYear 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!???() function (DOC: Han, which function?).  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeWithRates(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortalityTable,
		[ExcelArgument( "Required.  A 6x2 array representing interest rates.  The first column refers to the period for which the corresponding interest rate in the second column applies." )]
		double[ , ] interestRates,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemporary = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemporary' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemporary = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[ExcelArgument( "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemporary' is A and 'whenTemporary' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0." )]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortalityTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? mortalitySizeAdjustment = null,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTemporaryArg = howTemporary.Check( nameof( howTemporary ), "Y" );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var mortalitySizeAdjustmentArg = mortalitySizeAdjustment.Check( nameof( mortalitySizeAdjustment ), 1d );

		return Annuity.SingleLife( 
			mortalityTable, interestRates, age, 
			howDeferArg, whenDefer, howTemporaryArg, whenTemporary, howCertainArg, whenCertain, 
			paymentsPerYearArg, mortalityTableAdjustment, mortalitySizeAdjustmentArg, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, dynamicImprovementStopYearArg 
		);
	}

	[ExcelFunction( Category = "Financial", Description = "Replacement function for the Cfgena.xla!SingleLife() function.  Returns a decimal value equal to the selected single life annuity factor." )]
	public static double BTRSingleLifeComm(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string mortalityTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double interestRate,
		[ExcelArgument( "Required.  The current age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( Name = "typeCF", Description = "Commutation Function (CF) type, please see CFGENA help for details." )]
		object? typeCF = null,
		[ExcelArgument( Name = "paymentsPerYear", Description = "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the mortality table (not the age). This is done before unisex blending." )]
		int mortalityTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? mortalitySizeAdjustment = null,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null
	)
	{
		var typeCFArg = typeCF.Check( nameof( typeCF ), "" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var mortalitySizeAdjustmentArg = mortalitySizeAdjustment.Check( nameof( mortalitySizeAdjustment ), 1d );

		return Annuity.SingleLifeComm( 
			mortalityTable, interestRate, age, typeCFArg, paymentsPerYearArg,
			mortalityTableAdjustment, mortalitySizeAdjustmentArg, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, dynamicImprovementStopYearArg 
		);
	}

	[DebugFunction]
	[ExcelFunctionDoc(
		Category = "Financial",
		Description = "Replacement function for the Cfgena.xla!PPAJointLife() function.  Returns a decimal value equal to the selected joint life annuity factor by the PPA method.",
		Remarks = "<p>If you defer a temporary or certain annuity to an age earlier than the individual’s current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity.</p><p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'interestRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberAge' is less than 1 or greater than 120.</li><li>'spouseAge' is less than 1 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortalityImprovement' is not 0, 1, 2, 31 or 32.</li><li>'mortalityImprovement' is 1 or 31 and 'unisexBlending' is 2 or 'memberYearOfBirth' is 0 or 'spouseYearOfBirth' is 0.</li><li>'continuingPercentage' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'unisexBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'maleTableAdjustment' or 'femaleTableAdjustment' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRPPAJointLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string memberMortalityTable,
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string spouseMortalityTable,
		[ExcelArgument( "Required.  The interest rate to use for years 0 - 4." )]
		double interestRate1,
		[ExcelArgument( "Required.  The interest rate to use for years 5 - 19." )]
		double interestRate2,
		[ExcelArgument( "Required.  The interest rate to use for years 20 and after." )]
		double interestRate3,
		[ExcelArgument( "Required.  The current member age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "Required.  The spouse age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double spouseAge,
		[ExcelArgument( "The options for annuity factors. 'C' for contingent, 'S' for survivor, 'P' for popup, 'D' for double popup and 'J' for joint life factor only.  Default value is 'C'." )]
		object? annuityOption = null,
		[ExcelArgument( "The fraction of contingent/survivor amount to the primary amount; 0-1.  Default value is 0.5." )]
		object? jointFraction = null,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemporary = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemporary' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemporary = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[ExcelArgument( "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemporary' is A and 'whenTemporary' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0." )]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int maleTableAdjustment = 0,
		[ExcelArgument( "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int femaleTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? maleSizeAdjustment = null,
		[ExcelArgument( "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? femaleSizeAdjustment = null,
		[ExcelArgument( "The unisex blending percentage applied to the male mortality table." )]
		object? maleUnisexPercentage = null,
		[ExcelArgument( "The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int unisexBlending = 0,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The year of birth for the spouse if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int spouseYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTemporaryArg = howTemporary.Check( nameof( howTemporary ), "Y" );
		var annuityOptionArg = annuityOption.Check( nameof( annuityOption ), "C" );
		var jointFractionArg = jointFraction.Check( nameof( jointFraction ), 0.5 );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var maleSizeAdjustmentArg = maleSizeAdjustment.Check( nameof( maleSizeAdjustment ), 1d );
		var femaleSizeAdjustmentArg = femaleSizeAdjustment.Check( nameof( femaleSizeAdjustment ), 1d );
		var maleUnisexPercentageArg = maleUnisexPercentage.Check( nameof( maleUnisexPercentage ), 1d );

		return Annuity.PPAJointLife( 
			memberMortalityTable, spouseMortalityTable, interestRate1, interestRate2, interestRate3, age, spouseAge, annuityOptionArg, jointFractionArg, howDeferArg, whenDefer, howTemporaryArg, whenTemporary, howCertainArg, whenCertain, paymentsPerYearArg,
			maleTableAdjustment, femaleTableAdjustment, maleSizeAdjustmentArg, femaleSizeAdjustmentArg, maleUnisexPercentageArg, unisexBlending, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, spouseYearOfBirth, dynamicImprovementStopYearArg 
		);
	}

	[DebugFunction]
	[ExcelFunctionDoc(
		Category = "Financial",
		Description = "Replacement function for the Cfgena.xla!JointLife() function.  Returns decimal value equal to the selected joint life annuity factor.",
		Remarks = "<p>If you defer a temporary or certain annuity to an age earlier than the individual’s current age, the result is calculated only for the remainder of the annuity. For example, a 10-year temporary annuity deferred to age 65 for a 70-year old means that there are only 5 years remaining in the annuity and, thus, the result is equivalent to an immediate 5-year temporary annuity.</p><p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'interestRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberAge' is less than 1 or greater than 120.</li><li>'spouseAge' is less than 1 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortalityImprovement' is not 0, 1, 2, 31 or 32.</li><li>'mortalityImprovement' is 1 or 31 and 'unisexBlending' is 2 or 'memberYearOfBirth' is 0 or 'spouseYearOfBirth' is 0.</li><li>'jointFraction' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'unisexBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'maleTableAdjustment' or 'femaleTableAdjustment' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRJointLife(
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string memberMortalityTable,
		[ExcelArgument( "Required.  The name of the mortality table that you wish to use." )]
		string spouseMortalityTable,
		[ExcelArgument( "Required.  The interest rate to use." )]
		double interestRate,
		[ExcelArgument( "Required.  The current member age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double age,
		[ExcelArgument( "Required.  The current spouse age (see Remarks) to calculate the factor of. Value is typically 20 to 110 (matching mortality table ages). Interpolation is performed if value is not integral." )]
		double spouseAge,
		[ExcelArgument( "The options for annuity factors. 'C' for contingent, 'S' for survivor, 'P' for popup, 'D' for double popup and 'J' for joint life factor only.  Default value is 'C'." )]
		object? annuityOption = null,
		[ExcelArgument( "The fraction of contingent/survivor amount to the primary amount; 0-1.  Default value is 0.5." )]
		object? jointFraction = null,
		[ExcelArgument( "The method used to determine deferrment age.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howDefer = null,
		[ExcelArgument( "The age or years for deferred payment; 0-110.  When 'howDefer' is Y, deferred age is 'age' + value, otherwise value.  Default value is 0." )]
		double whenDefer = 0,
		[ExcelArgument( "The method used to determine temporary period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howTemporary = null,
		[ExcelArgument( "The age or years for temporaryPeriod; 0-120.  When 'howTemporary' is Y, temporary period age is 'age' + value, otherwise value.  Default value is 121, which means 'for life'." )]
		double whenTemporary = 0,
		[ExcelArgument( "The method used to determine certain period.  'A' for age, 'Y' for years.  Default value is 'Y'." )]
		object? howCertain = null,
		[ExcelArgument( "The age or years for certain period; 0-110.  This has no effect prior to payment start time.  When 'howTemporary' is A and 'whenTemporary' is less than deferred age, certain period is 0, else when 'howCertain' is A, certain period is value - deferred age, otherwise value.  Default value is 0." )]
		double whenCertain = 0,
		[ExcelArgument( "The frequency of payments per year (positive for beginning of period or negative for end of period payments).  Default value is 12." )]
		object? paymentsPerYear = null,
		[ExcelArgument( "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int maleTableAdjustment = 0,
		[ExcelArgument( "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int femaleTableAdjustment = 0,
		[ExcelArgument( "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? maleSizeAdjustment = null,
		[ExcelArgument( "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? femaleSizeAdjustment = null,
		[ExcelArgument( "The unisex blending percentage applied to the male mortality table." )]
		object? maleUnisexPercentage = null,
		[ExcelArgument( "The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int unisexBlending = 0,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 31 = DynamicScaleBB, and 32 = StaticScaleBB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The year of birth for the spouse if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int spouseYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null
	)
	{
		var howDeferArg = howDefer.Check( nameof( howDefer ), "Y" );
		var howTemporaryArg = howTemporary.Check( nameof( howTemporary ), "Y" );
		var annuityOptionArg = annuityOption.Check( nameof( annuityOption ), "C" );
		var jointFractionArg = jointFraction.Check( nameof( jointFraction ), 0.5 );
		var howCertainArg = howCertain.Check( nameof( howCertain ), "Y" );
		var paymentsPerYearArg = paymentsPerYear.Check( nameof( paymentsPerYear ), 12 );
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var maleSizeAdjustmentArg = maleSizeAdjustment.Check( nameof( maleSizeAdjustment ), 1d );
		var femaleSizeAdjustmentArg = femaleSizeAdjustment.Check( nameof( femaleSizeAdjustment ), 1d );
		var maleUnisexPercentageArg = maleUnisexPercentage.Check( nameof( maleUnisexPercentage ), 1d );

		return Annuity.JointLife( 
			memberMortalityTable, spouseMortalityTable, interestRate, age, spouseAge, annuityOptionArg, jointFractionArg, howDeferArg, whenDefer, howTemporaryArg, whenTemporary, howCertainArg, whenCertain, paymentsPerYearArg,
			maleTableAdjustment, femaleTableAdjustment, maleSizeAdjustmentArg, femaleSizeAdjustmentArg, maleUnisexPercentageArg, unisexBlending, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, spouseYearOfBirth, dynamicImprovementStopYearArg 
		);
	}

	[DebugFunction]
	[ExcelFunctionDoc(
		Category = "Financial",
		Description = "A replacement function for the Annbuck.xla!AnnBuck() function.  Returns a decimal value representing the selected life annuity factor.",
		Remarks = "<p>Non-integer values for 'memberAge', 'spouseAge', 'deferredAge', 'temporaryPeriod', and 'guaranteePeriod' can be used.  The factor will then be interpolated.</p><p>An ArgumentOutOfRangeException can be thrown if any of the following conditions occur:</p><ul><li>'interestRates' durations contain any negative or decimal numbers or the sum of the durations greater than 120.</li><li>'memberSex' is not 1 or 2.</li><li>'spouseSex' is not 1 or 2.</li><li>'memberAge' is less than 0 or greater than 120.</li><li>'spouseAge' is less than 0 or greater than 120.</li><li>'deferredAge' is less than 0 or greater than 120 or less than 'memberAge' (when deferredAge > 0).</li><li>'mortalityImprovement' is not 0, 1, 2, 11, 12, 21, 22, 31, 32, 41, 42, 51, or 52.</li><li>'mortalityImprovement' is 1, 11, 21, 31, 41, or 51 and 'unisexBlending' is 2 or 'memberYearOfBirth' is 0 or 'spouseYearOfBirth' is 0.</li><li>'mortalityImprovement' is 21, 22, 51, or 52 and the static year of selected male or female mortality tables are less than 2014.</li><li>'continuingPercentage' is less than 0.</li><li>'maleUnisexPercentage' is less than 0 or greater than 1.</li><li>'paymentTiming' is not 1, 2, or 3.</li><li>'preRetirementMortality' is not 1, 2, 3, 4, 5, or 6.</li><li>'unisexBlending' is not 0, 1, or 2.</li><li>'temporaryPeriod' is less than 0 or greater than Min( 120 - Max( 'memberAge', 'spouseAge' ), 120 - 'deferredAge' ).</li><li>'guaranteePeriod' is greater than 'temporaryPeriod' (when temporaryPeriod is greater than 0).</li><li>'maleTableAdjustment' or 'femaleTableAdjustment' is less than Max( -10, -age ) or greater than Min( 10, 120 - age ).</li></ul>"
	)]
	public static double BTRAnnBuck(
		[ExcelArgument( "Required.  A 6x2 array representing interest rates.  The first column refers to the period for which the corresponding interest rate in the second column applies." )]
		double[ , ] interestRates,
		[ExcelArgument( "Required.  The 'PaymentTimingType' to use where 1 = Continuous Approximation, 2 = Beginning of the month, and 3 = End of the month." )]
		int paymentTiming,
		[ExcelArgument( "Required.  The member's age to calculate; must be a decimal number between 0 and 120, inclusive." )]
		double memberAge,
		[ExcelArgument( "Required.  The spouse's age to calculate; must be a decimal number between 0 and 120, inclusive.  Use 0 when not calculating joint factors." )]
		double spouseAge,
		[ExcelArgument( "Required. The member's 'SexType' to use where 1 = Male and 2 = Female." )]
		int memberSex,
		[ExcelArgument( "Required. The spouse's 'SexType' to use where 1 = Male and 2 = Female." )]
		int spouseSex,
		[ExcelArgument( "Required.  The age that benefits commence. For immediate factors, enter 0 or value equal to 'memberAge'." )]
		double deferredAge,
		[ExcelArgument( "Required.  The number of years that payments are made. Enter 0 if there is no temporary period." )]
		double temporaryPeriod,
		[ExcelArgument( "Required.  The number of years for which payments are guaranteed upon death. Enter 0 if there is no guarantee period." )]
		double guaranteePeriod,
		[ExcelArgument( "Required.  The percentage that which payments will continue to the spouse upon death of the member." )]
		double continuingPercentage,
		[ExcelArgument( "Required.  The 'PreRetirementMortalityType' to use, where 1 = NoMortality, 2 = MemberRetirementAgeJointSurvivor, 3 = MemberOnly, 4 = MemberDeathNoGuarantee, 5 = MemberRetirementAgeNoGuarantee, and 6 = MemberRetirementAgeFull." )]
		int preRetirementMortality,
		[ExcelArgument( "Required.  The number of the mortality table that you wish to use. For example, enter '214' for GAM83 Male." )]
		string maleMortalityTable,
		[ExcelArgument( "Required.  The number of the mortality table that you wish to use.  See 'maleMortalityTable'." )]
		string femaleMortalityTable,
		[ExcelArgument( "Required.  The unisex blending percentage applied to the male mortality table." )]
		double maleUnisexPercentage,
		[ExcelArgument( "Required.  The 'UnisexBlendingType' to use, where 0 = Unisex off (sex distinct), 1 = Unisex blending by mortality rates, and 2 = Unisex blending by annuity factors." )]
		int unisexBlending,
		[ExcelArgument( "The adjustment years to apply as a shift to the male mortality table (not the age). This is done before unisex blending." )]
		int maleTableAdjustment = 0,
		[ExcelArgument( "The adjustment years to apply as a shift to the female mortality table (not the age). This is done before unisex blending." )]
		int femaleTableAdjustment = 0,
		[ExcelArgument( "The 'MortalityImprovementType' to use, where 0 = Disable, 1 = DynamicScaleAA, 2 = StaticScaleAA, 11 = DynamicScaleCPMA1, 12 = StaticScaleCPMA1, 21 = DynamicScaleCPMA, 22 = StaticScaleCPMA, 31 = DynamicScaleBB, 32 = StaticScaleBB, 41 = DynamicScaleCPMB1, 42 = StaticScaleCPMB1, 51 = DynamicScaleCPMB, and 52 = StaticScaleCPMB." )]
		int mortalityImprovement = 0,
		[ExcelArgument( "The year of calculation if 'mortalityImprovement' is 1 or 2. The year entered in this argument will determine the effective year of the projected mortality table." )]
		object? improvementEffectiveYear = null,
		[ExcelArgument( "The year of birth for the member if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int memberYearOfBirth = 0,
		[ExcelArgument( "The year of birth for the spouse if dynamic generational is enabled 'mortalityImprovement' is 1." )]
		int spouseYearOfBirth = 0,
		[ExcelArgument( "The stop year if dynamic generational is enabled ('mortalityImprovement' is 1) and you wish to stop the generational projection at a future year. This caps the exponent of the projection factors. If this parameter is not filled in or a „0‟ then it is assumed that the improvements continue indefinitely." )]
		object? dynamicImprovementStopYear = null,
		[ExcelArgument( "A multiplier to adjust the male mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? maleSizeAdjustment = null,
		[ExcelArgument( "A multiplier to adjust the female mortality rates (qx). Entering 0 or leaving it blank will multiply the rates by 1 (unchanged). Any other figure will serve as a multiplier. This action is done prior to applying unisex or generational improvements on the mortality table." )]
		object? femaleSizeAdjustment = null 
	)
	{
		var improvementEffectiveYearArg = improvementEffectiveYear.Check( nameof( improvementEffectiveYear ), -1 );
		var dynamicImprovementStopYearArg = dynamicImprovementStopYear.Check( nameof( dynamicImprovementStopYear ), 9999 );
		var maleSizeAdjustmentArg = maleSizeAdjustment.Check( nameof( maleSizeAdjustment ), 1d );
		var femaleSizeAdjustmentArg = femaleSizeAdjustment.Check( nameof( femaleSizeAdjustment ), 1d );

		return Annuity.CanadaAnnuity( interestRates, paymentTiming, memberAge, spouseAge, memberSex, spouseSex, deferredAge, temporaryPeriod, guaranteePeriod, continuingPercentage, preRetirementMortality, maleMortalityTable, femaleMortalityTable, maleUnisexPercentage, unisexBlending, maleTableAdjustment, femaleTableAdjustment, mortalityImprovement, improvementEffectiveYearArg, memberYearOfBirth, spouseYearOfBirth, dynamicImprovementStopYearArg, maleSizeAdjustmentArg, femaleSizeAdjustmentArg );
	}
}