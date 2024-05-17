using System.Linq.Expressions;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.XmlMappingExport.GeneratedParser;

public class XmlMappingVisitor : ExpressionParseTreeVisitor
{
	public ParameterExpression Context { get; protected set; }

	public XmlMappingVisitor()
	{
		this.Context = Expression.Parameter( typeof( XmlContext ), "context" );
	}

	protected override Expression GetMethodExpression( string name, IEnumerable<Expression> parameters )
	{
		var method = Context.Type.GetMethods().FirstOrDefault( m => m.Name == name );

		if ( method != null )
		{
			var methodParameters = method.GetParameters();
			var parametersCount = parameters.Count();

			// Optional 'scopeDepth' parameter for MapOrdinal
			if ( name == "MapOrdinal" && parametersCount == 0 )
			{
				parameters = new[] { Expression.Constant( 1 ) };
				parametersCount++;
			}

			// Last parameter to mapping functions is a 'default' value for Excel usability.  It might be omitted so need to appened a null
			var parsedParameters = parametersCount == methodParameters.Length - 1
				? parameters.Append( GetNullExpression() )
				: parameters;

			return Expression.Call(
				Context,
				method,
				methodParameters.Zip(
					parsedParameters,
					( p, a ) => EnsureType( a, p.ParameterType ) ).ToArray() );
		}

		throw new NotImplementedException( string.Format( "Method '{0}' not supported", name ) );
	}

	protected Expression EnsureType( Expression value, Type type )
	{
		return ( value.Type != type )
			? ChangeType( value, type )
			: value;
	}

	protected override Expression ChangeType( Expression value, Type type )
	{
		return Expression.Convert(
			Expression.Call( Context, "ChangeType", null, Expression.Convert( value, typeof( object ) ), Expression.Constant( type ) ),
			type );
	}
}
