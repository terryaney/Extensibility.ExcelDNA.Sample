using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

internal static class FunctionExtensions
{
	/// <summary>
	/// The OptionalValues module contains methods used to handle default values passed from Excel/Excel-DNA into 
	/// C# allowing default values to be set correctly for any parameters that don't default to c# type defaults.
	/// </summary>
	/// <seealso href="https://excel-dna.net/docs/guides-basic/optional-parameters-and-default-values">Optional Parameters and Default Values</seealso>
	/// <example>
	/// This sample shows how to set default parameter values.
	/// <code>
	/// function foo( string s = "defA", int i = -1, double d = -1d, string sOk = null, int iOk = 0, double dOk = 0d )
	/// {
	///     /*
	///     By default, when using this function in Excel as =foo(), Excel-DNA would pass the following parameter values:
	///
	///     s = null,
	///     i = 0,
	///     d = 0d,
	///     sOk = null,
	///     iOk = 0,
	///     dOk = 0d
	///
	///     Since Excel-DNA already converts Excel parameters into specific types when the optional parameter is missing, you cannot
	///     have typed parameters for the parameters you need to fix.  See fooWithDefaults below.
	///     */
	/// }
	///
	/// function fooWithDefaults( object sArg = null, object iArg = null, object dArg = null, string sOk = null, int iOk = 0, double dOk = 0d )
	/// {
	///     /*
	///     Each parameter that had a default value different that the Default&lt;T> value (i.s. the *Ok params are valid because the
	///     default value assigned to each of them is the same as Default&lt;T>) had to be converted to an object type.  And when a parameter
	///     is of type object, you can only assign null as the default value.
	///
	///     Note that, appending an 'Arg' suffix is the suggested naming pattern, so when you create variables below, you fall back to the variable
	///     name as intended for subsequent coding.  After this conversion, you would then assign defaults the way you want.
	///     */
	///
	///     var s = OptionalValues.Check&lt;string>( sArg, "s", "defA" );
	///     var i = OptionalValues.Check&lt;int>( iArg, "i", -1 );
	///     var d = OptionalValues.Check&lt;double>( dArg, "d", -1d );
	/// }
	/// </code>
	/// </example>
	internal static T Check<T>( this object? arg, string argumentName, T defaultValue )
	{
		if ( arg == null || arg is ExcelMissing )
		{
			return defaultValue;
		}

		if ( arg is T t )
		{
			return t;
		}
		if ( ( typeof( T ) == typeof( int ) || typeof( T ) == typeof( int? ) ) && arg is double ) // when arg is object, all numbers are passed as double, when arg is int, DNA handles convert...			
		{
			return (T)Convert.ChangeType( arg, typeof( int ) );
		}
		if ( typeof( T ) == typeof( string ) )
		{
			return (T)Convert.ChangeType( arg, typeof( string ) );
		}
		// Didn't use this b/c I first have to convert object[,]->double[,], then I casted to an object, then casted back to a double[,]...too much work/overhead
		// and I couldn't figure out how to simply cast like return (T)Functions.RBLeMacros.CastArray<double>( arg );
		// UPDATE: Removed CastArray b/c it evolved into GetArray<T>.
		/*
		else if ( typeof( T ) == typeof( double[, ] ) && arg is object[, ] )
		{
			object d = Functions.RBLeMacros.CastArray<double>( arg );
			return (T)d;
		}
		*/
		else
		{
			throw new ArgumentException( "Invalid value passed in for this parameter.", argumentName );
		}
	}
}
