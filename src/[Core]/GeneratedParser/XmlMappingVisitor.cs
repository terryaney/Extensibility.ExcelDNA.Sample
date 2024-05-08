using System.Linq.Expressions;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

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
			return Expression.Call(
				Context,
				method,
				method.GetParameters().Zip(
					parameters,
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
