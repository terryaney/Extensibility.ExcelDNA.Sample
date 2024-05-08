namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

public abstract class ParseTreeVisitor<T, TIdentifier>
{
	public HashSet<TIdentifier> Dependencies { get; private set; } = new();

	protected Stack<string> methodStack = new();

	public T GetExpression( ParseTree tree ) => GetNodeExpression( tree );

	protected virtual T GetNodeExpression( ParseNode node )
	{
		return node.Token.Type switch
		{
			TokenType.Start => GetStartExpression( node ),
			TokenType.Assignment => GetAssignmentExpression( node ),
			TokenType.CompareExpr => GetComparisonExpression( node ),
			TokenType.AddExpr => GetAddExpression( node ),
			TokenType.MultExpr => GetMultiplyExpression( node ),
			TokenType.Atom => GetAtomExpression( node ),
			TokenType.Range => GetRangeExpression( node ),
			TokenType.TRUE => GetBooleanExpression( true ),
			TokenType.FALSE => GetBooleanExpression( false ),
			TokenType.NULL => GetNullExpression(),
			TokenType.String => GetString( node ),
			TokenType.INTEGER => GetIntegerExpression( node ),
			TokenType.NUMBER => GetNumberExpression( node ),
			TokenType.Sign => GetSignExpression( node ),
			TokenType.Identifier => GetIdentifierExpression( node ),
			TokenType.Method => GetMethodExpression( node ),
			TokenType.Constructor => GetConstructorExpression( node ),
			TokenType.Eval => GetEvalExpression( node ),
			TokenType.Data => GetDataExpression( node ),
			_ => GetString( node ),
		};
	}

	protected virtual T GetStartExpression( ParseNode node ) => GetNodeExpression( node.Nodes[ 0 ] );
	protected virtual T GetIntegerExpression( ParseNode node ) => GetIntegerExpression( long.Parse( node.Token.Text ) );
	protected virtual T GetNumberExpression( ParseNode node ) => GetNumberExpression( double.Parse( node.Token.Text ) );
	protected virtual T GetAtomExpression( ParseNode node )
	{
		return ( node.Nodes.Count > 2 && node.Nodes[ 0 ].Token.Type == TokenType.BROPEN )
			? GetGroupExpression( GetNodeExpression( node.Nodes[ 1 ] ) )
			: ( node.Nodes.Count > 0 ) ? GetNodeExpression( node.Nodes[ 0 ] ) : GetNullExpression();
	}

	protected virtual T GetSignExpression( ParseNode node )
	{
		if ( node.Nodes.Count < 1 ) return GetNullExpression();

		return ( node.Nodes[ 0 ].Token.Text == "-" )
			? GetSignExpression( true, GetNodeExpression( node.Nodes[ 1 ] ) )
			: GetSignExpression( false, GetNodeExpression( node.Nodes[ 0 ] ) );
	}

	protected string GetText( ParseNode node ) => string.Concat( node.Nodes.Select( n => n.Token.Text ).ToArray() );

	protected virtual T GetIdentifierExpression( ParseNode node )
	{
		var id = GetText( node );
		AddDependency( id );

		return GetIdentifierExpression( id );
	}

	protected virtual T GetEvalExpression( ParseNode node )
	{
		var id = GetText( node ).Replace( "{Eval ", "" ).Replace( "}", "" ).Replace( " ", "_" ).Replace( "-", "_" );
		AddDependency( id );

		return GetEvalExpression( id );
	}

	protected virtual T GetDataExpression( ParseNode node )
	{
		var id = GetText( node ).Replace( "{Data ", "Data." ).Replace( "}", "" ).Replace( " ", "_" ).Replace( "-", "_" );
		AddDependency( id );

		return GetDataExpression( id );
	}

	protected TIdentifier? AddDependency( string id )
	{
		var identifier = GetIdentifier( id );
		if ( identifier == null ) return default;

		if ( !Dependencies.Contains( identifier ) ) Dependencies.Add( identifier );

		return identifier;
	}

	protected abstract TIdentifier GetIdentifier( string id );

	protected virtual T GetRangeExpression( ParseNode node )
	{
		var cellIds = ( from n in node.Nodes
						where n.Token.Type == TokenType.CELLID
						select new CellId( n.Token.Text ) ).ToArray();

		return ( cellIds.Length > 1 )
			? GetRangeExpression( new Range( cellIds ).Cells )
			: GetIdentifierExpression( cellIds.First().Name );
	}

	protected virtual T GetMethodExpression( ParseNode node ) => GetMethodExpression( GetMethodName( node ), node );

	protected virtual T GetMethodExpression( string name, ParseNode node )
	{
		var lowerName = name.ToLower();

		methodStack.Push( name );

		try
		{
			var parameters = ( from n in node.Nodes
							   where n.Token.Type == TokenType.Params
							   from p in n.Nodes
							   where p.Token.Type != TokenType.COMMA
							   from e in GetNodeExpressions( p )
							   select e ).ToArray();

			if ( lowerName == "iif" || lowerName == "if" ) return GetConditionExpression( parameters );

			return GetMethodExpression( name, parameters );
		}
		finally
		{
			methodStack.Pop();
		}
	}

	protected virtual T GetConstructorExpression( ParseNode node )
	{
		var parameters = ( from n in node.Nodes
						   where n.Token.Type == TokenType.Params
						   from p in n.Nodes
						   where p.Token.Type != TokenType.COMMA
						   from e in GetNodeExpressions( p )
						   select e ).ToArray();

		var name = node.Nodes.First().Token.Text;

		return GetMethodExpression( name, parameters );
	}

	protected virtual string GetMethodName( ParseNode node )
	{
		return node.Nodes.First().Token.Text;
	}

	protected abstract IEnumerable<T> GetNodeExpressions( ParseNode node );
	protected abstract T GetAssignmentExpression( ParseNode node );
	protected abstract T GetGroupExpression( T expression );
	protected abstract T GetNullExpression();
	protected abstract T GetBooleanExpression( bool value );
	protected abstract T GetIntegerExpression( long value );
	protected abstract T GetNumberExpression( double value );
	protected abstract T GetComparisonExpression( ParseNode node );
	protected abstract T GetAddExpression( ParseNode node );
	protected abstract T GetMultiplyExpression( ParseNode node );
	protected abstract T GetSignExpression( bool isNegative, T expression );
	protected abstract T GetRangeExpression( IEnumerable<CellId> cellIds );
	protected abstract T GetString( ParseNode node );
	protected abstract T GetMethodExpression( string name, IEnumerable<T> parameters );
	protected abstract T GetConditionExpression( T[] args );
	protected abstract T GetIdentifierExpression( string id );
	protected abstract T GetEvalExpression( string eval );
	protected abstract T GetDataExpression( string data );
}
