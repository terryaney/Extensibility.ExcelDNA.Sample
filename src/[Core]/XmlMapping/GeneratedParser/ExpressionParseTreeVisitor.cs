using System.Linq.Expressions;
using System.Xml.Linq;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.XmlMappingExport.GeneratedParser;

public abstract class ExpressionParseTreeVisitor : ParseTreeVisitor<Expression, string>
{
	protected override string GetIdentifier( string id ) => id;
	protected override IEnumerable<Expression> GetNodeExpressions( ParseNode node ) => new[] { GetNodeExpression( node ) };
	protected override Expression GetAssignmentExpression( ParseNode node ) => null!;
	protected override Expression GetGroupExpression( Expression expression ) => expression;
	protected override Expression GetNullExpression() => Expression.Constant( null );
	protected override Expression GetBooleanExpression( bool value ) => Expression.Constant( value );
	protected override Expression GetIntegerExpression( long value ) => Expression.Constant( value );
	protected override Expression GetNumberExpression( double value ) => Expression.Constant( value );

	protected override Expression GetAddExpression( ParseNode node )
	{
		Func<Expression, Expression, BinaryExpression> oper = null!;
		Expression expr = null!;
		var isConcat = false;

		foreach ( var n in node.Nodes )
		{
			if ( n.Token.Type == TokenType.PLUSMINUS )
			{
				switch ( n.Token.Text )
				{
					case "+": oper = Expression.Add; break;
					case "-": oper = Expression.Subtract; break;
					case "&": isConcat = true; break;
				}
			}
			else
			{
				var next = GetNodeExpression( n );

				expr = ( expr == null || ( oper == null && !isConcat ) )
					? next
					: isConcat
						? GetConcatExpression( expr, next )
						: GetBinaryExpression( oper!, expr, next );
			}
		}

		return expr;
	}

	private static Expression GetConcatExpression( params Expression[] expressions )
	{
		var method = typeof( string ).GetMethod( "Concat", new[] { typeof( string ), typeof( string ) } )!;

		return Expression.Call( method, expressions );
	}

	protected override Expression GetMultiplyExpression( ParseNode node )
	{
		Func<Expression, Expression, BinaryExpression> oper = null!;
		Expression expr = null!;

		foreach ( var n in node.Nodes )
		{
			if ( n.Token.Type == TokenType.MULTDIV )
			{
				switch ( n.Token.Text )
				{
					case "*": oper = Expression.Multiply; break;
					case "/": oper = Expression.Divide; break;
				}
			}
			else
			{
				var next = GetNodeExpression( n );

				expr = ( expr == null || oper == null )
					? next
					: GetBinaryExpression( oper, expr, next );
			}
		}

		return expr;
	}

	protected override Expression GetComparisonExpression( ParseNode node )
	{
		Func<Expression, Expression, BinaryExpression> oper = null!;
		Expression? expr = null;

		foreach ( var n in node.Nodes )
		{
			if ( n.Token.Type == TokenType.COMPARE )
			{
				switch ( n.Token.Text )
				{
					case "=":
					case "==": oper = Expression.Equal; break;
					case ">": oper = Expression.GreaterThan; break;
					case ">=": oper = Expression.GreaterThanOrEqual; break;
					case "<": oper = Expression.LessThan; break;
					case "<=": oper = Expression.LessThanOrEqual; break;
					case "!=":
					case "<>": oper = Expression.NotEqual; break;
				}
			}
			else
			{
				var next = GetNodeExpression( n );

				expr = ( expr == null )
					? next
					: GetBinaryExpression( oper, expr, next );
			}
		}

		return expr!;
	}

	protected virtual Expression GetBinaryExpression( Func<Expression, Expression, BinaryExpression> oper, Expression left, Expression right )
	{
		var leftTypeCode = left.Type.NullableTypeCode();
		var rightTypeCode = right.Type.NullableTypeCode();

		if ( leftTypeCode != rightTypeCode )
		{
			// Convert the left or right side to the "best" type
			if ( (int)leftTypeCode > (int)rightTypeCode && right.NodeType != ExpressionType.Constant ) right = ChangeType( right, left.Type );
			else left = ChangeType( left, right.Type );
		}

		if ( oper == Expression.Add && left.Type == typeof( string ) && right.Type == typeof( string ) )
		{
			return Expression.Add( left, right, typeof( string ).GetMethod( "Concat", new[] { typeof( string ), typeof( string ) } ) );
		}

		return oper( left, right );
	}

	protected virtual Expression ChangeType( Expression value, Type type ) =>
		Expression.Convert(
			Expression.Call(
				typeof( Convert ),
				"ChangeType",
				null,
				Expression.Convert( value, typeof( object ) ),
				Expression.Constant( type )
			),
			type
		);

	protected override Expression GetSignExpression( bool isNegative, Expression expression ) => isNegative ? Expression.Not( expression ) : expression;

	protected override Expression GetString( ParseNode node )
	{
		var textNode = ( node.Nodes.Count > 1 )
			? node.Nodes[ 1 ]
			: ( node.Nodes.Count > 0 ) ? node.Nodes[ 0 ] : node;

		return ( textNode != null )
			? Expression.Constant( textNode.Token.Text )
			: Expression.Constant( null );
	}

	protected override Expression GetConditionExpression( Expression[] args ) => Expression.Condition( args[ 0 ], args[ 1 ], args[ 2 ] );

	protected override Expression GetIdentifierExpression( string id ) => throw new Exception( $"Implement GetIdentifierExpression for {id}." );
	protected override Expression GetEvalExpression( string value ) => throw new Exception( "Implement GetEvalExpression" );
	protected override Expression GetDataExpression( string value ) => throw new Exception( "Implement GetDataExpression" );
	protected override Expression GetRangeExpression( IEnumerable<CellId> cellIds ) => throw new Exception( "Implement GetRangeExpression" );
}
