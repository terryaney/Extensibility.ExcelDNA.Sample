using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaAnnualLimits
{
	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with '415DBLimit' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing maximum annual defined benefit payable at Social Security normal retirement age under §415(b)." )]
	public static int BTR415DBMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max415DefinedBenefit( year, rateProj );
	

	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with '415DCLimit' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing maximum Annual Addition under §415(c). Applies to total (including company) contributions." )]
	public static int BTR415DCMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max415DefinedContribution( year, rateProj );
	

	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with 'Max401(k)Contribution' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing maximum permitted salary deferral under §401(k)." )]
	public static int BTR401kMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max401k( year, rateProj );

	[DebugFunction]
	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with '401(a)(17)PayLimit' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing maximum compensation recognized under qualified pension or profit sharing plans (ignores EGTRRA limit changes). Post-2001 limits are updated each year with the annual CPI/W rate." )]
	public static int BTR401A17Max(
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max401a17( year, rateProj );

	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with 'HCECompLimit' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing annual compensation used to define highly compensated employee after 1996." )]
	public static int BTRHCELimit( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.MaxHighlyCompensatedLimit( year, rateProj );

	[ExcelFunction( Category = "Annual Limits", Description = "DOC: Han, Cfgena replacement?  Returns an integer value representing maximum permitted salary deferral under §403(b)." )]
	public static int BTR403bMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max403b( year, rateProj );

	[ExcelFunction( Category = "Annual Limits", Description = "DOC: Han, Cfgena replacement?  Returns an integer value representing maximum permitted salary deferral under §457(b)." )]
	public static int BTR457bMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max457b( year, rateProj );

	[ExcelFunction( Category = "Annual Limits", Description = "A replacement function for the Cfgena.xla!AnnualLimits() function call with 'Max401(k)Make-up' name parameter in favor of explicit function name (reduces parameter typo errors).  Returns an integer value representing maximum make-up contribution permitted for 50+ year-olds under §401(k)." )]
	public static int BTRCatchupMax( 
		[ExcelArgument( "The effective year for the limit, if this parameter is 0 then function will return current year unrounded limit." )] 
		int year,
		[ExcelArgument( "The increase rate at which to project the limit." )] 
		double rateProj 
	) => AnnualLimits.Max401kCatchUp( year, rateProj );
}