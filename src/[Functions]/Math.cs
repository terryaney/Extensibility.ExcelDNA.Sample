using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaMathUtitlies
{
	[ExcelFunction( Category = "Financial", Description = "Calculates the present value of a loan or an investment, based on a variable, multiple interest rates." )]
	public static double BTRPV(
		[ExcelArgument( "Required. The total number of periods to calculate.  For both interestRates and paymentAmounts, if the period is greater than number of elements present in each array, the array is padded to the end using the last element provided." )]
		int period,
		[ExcelArgument( "Required. The interest rates per period. For example, if you obtain an automobile loan at a 10 percent annual interest rate and make monthly payments, your interest rate per month is 10%/12, or 0.83%. You would enter 10%/12, or 0.83%, or 0.0083, into the formula as the rate." )]
		double[] interestRates,
		[ExcelArgument( "Required. The payment amounts made each period. Typically, each amount includes principal and interest but no other fees or taxes. For example, a monthly payment on a $10,000, four-year car loan at 12 percent is $263.33. You would enter -263.33 into the formula as one of the amounts." )]
		double[] paymentAmounts,
		[ExcelArgument( "Optional. The number 0 (at the end of the period) or 1 (at the beginning of the period) and indicates when payments are due.  The default is 0." )]
		int paymentTiming = 0 
	) => MathUtilities.PV( period, interestRates, paymentAmounts, paymentTiming );
	

	[ExcelFunction( Category = "Financial", Description = "Calculates elapsed days using 360 days per year." )]
	public static double BTRElapsed360(
		[ExcelArgument( "Required. The start date." )]
		DateTime dateStart,
		[ExcelArgument( "Required. The end date." )]
		DateTime dateEnd 
	) => MathUtilities.Elapsed360( dateStart, dateEnd );
	
	[ExcelFunction( Category = "Financial", Description = "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic." )]
	public static double BTRXIRR(
		[ExcelArgument( "Required. A series of cash flows that corresponds to a schedule of payments in dates. The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment. If the first value is a cost or payment, it must be a negative value. All succeeding payments are discounted based on a 365-day year. The series of values must contain at least one positive and one negative value." )]
		double[] values,
		[ExcelArgument( "Required. A schedule of payment dates that corresponds to the cash flow payments. Dates may occur in any order. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text." )]
		double[] dates,
		[ExcelArgument( "Optional. The default is 25." )]
		object? iterations = null,
		[ExcelArgument( "Optional. The default is 5.0." )]
		object? maxRate = null,
		[ExcelArgument( "Optional. The default is true." )]
		object? truncateTime = null 
	)
	{
		var iterationsArg = iterations.Check( nameof( iterations ), 25 );
		var maxRateArg = maxRate.Check( nameof( maxRate ), 5d );
		var truncateTimeArg = truncateTime.Check( nameof( truncateTime ), true );

		return MathUtilities.XIRR( values, dates.Select( d => DateTime.FromOADate( d ) ).ToArray(), iterationsArg, maxRateArg, truncateTimeArg );
	}
}