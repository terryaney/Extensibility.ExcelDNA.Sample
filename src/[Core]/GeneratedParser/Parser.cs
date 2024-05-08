namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

public partial class Parser
{
	private readonly Scanner scanner;
	private ParseTree tree = null!;

	public Parser( Scanner scanner )
	{
		this.scanner = scanner;
	}

	public ParseTree Parse( string input )
	{
		tree = new ParseTree();
		return Parse( input, tree );
	}

	public ParseTree Parse( string input, ParseTree tree )
	{
		scanner.Init( input );

		this.tree = tree;
		ParseStart( tree );
		tree.Skipped = scanner.Skipped;

		return tree;
	}

	private void ParseStart( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Start ), "Start" );
		parent.Nodes.Add( node );



		tok = scanner.LookAhead( TokenType.ASSIGNID, TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN );
		if ( tok.Type == TokenType.ASSIGNID
			|| tok.Type == TokenType.SIGN
			|| tok.Type == TokenType.METHODNAME
			|| tok.Type == TokenType.CONSTRUCTOR
			|| tok.Type == TokenType.SHEETBEGIN
			|| tok.Type == TokenType.CELLID
			|| tok.Type == TokenType.IDENTIFIER
			|| tok.Type == TokenType.EVAL
			|| tok.Type == TokenType.DATA
			|| tok.Type == TokenType.TRUE
			|| tok.Type == TokenType.FALSE
			|| tok.Type == TokenType.NULL
			|| tok.Type == TokenType.INTEGER
			|| tok.Type == TokenType.NUMBER
			|| tok.Type == TokenType.QUOTEBEGIN
			|| tok.Type == TokenType.SNGQUOTEBEGIN
			|| tok.Type == TokenType.INVALID
			|| tok.Type == TokenType.BROPEN )
		{
			tok = scanner.LookAhead( TokenType.ASSIGNID, TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN );
			switch ( tok.Type )
			{
				case TokenType.ASSIGNID:
					ParseAssignment( node );
					break;
				case TokenType.SIGN:
				case TokenType.METHODNAME:
				case TokenType.CONSTRUCTOR:
				case TokenType.SHEETBEGIN:
				case TokenType.CELLID:
				case TokenType.IDENTIFIER:
				case TokenType.EVAL:
				case TokenType.DATA:
				case TokenType.TRUE:
				case TokenType.FALSE:
				case TokenType.NULL:
				case TokenType.INTEGER:
				case TokenType.NUMBER:
				case TokenType.QUOTEBEGIN:
				case TokenType.SNGQUOTEBEGIN:
				case TokenType.INVALID:
				case TokenType.BROPEN:
					ParseCompareExpr( node );
					break;
				default:
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found.", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					break;
			}
		}


		tok = scanner.Scan( TokenType.EOF );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.EOF )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.EOF.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseAssignment( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Assignment ), "Assignment" );
		parent.Nodes.Add( node );



		tok = scanner.Scan( TokenType.ASSIGNID );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.ASSIGNID )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.ASSIGNID.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.Scan( TokenType.EQUALS );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.EQUALS )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.EQUALS.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		ParseCompareExpr( node );

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseCompareExpr( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.CompareExpr ), "CompareExpr" );
		parent.Nodes.Add( node );



		ParseAddExpr( node );


		tok = scanner.LookAhead( TokenType.COMPARE );
		if ( tok.Type == TokenType.COMPARE )
		{


			tok = scanner.Scan( TokenType.COMPARE );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.COMPARE )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.COMPARE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}


			ParseAddExpr( node );
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseAddExpr( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.AddExpr ), "AddExpr" );
		parent.Nodes.Add( node );



		ParseMultExpr( node );


		tok = scanner.LookAhead( TokenType.PLUSMINUS );
		while ( tok.Type == TokenType.PLUSMINUS )
		{


			tok = scanner.Scan( TokenType.PLUSMINUS );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.PLUSMINUS )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.PLUSMINUS.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}


			ParseMultExpr( node );
			tok = scanner.LookAhead( TokenType.PLUSMINUS );
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseMultExpr( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.MultExpr ), "MultExpr" );
		parent.Nodes.Add( node );



		ParseSign( node );


		tok = scanner.LookAhead( TokenType.MULTDIV );
		while ( tok.Type == TokenType.MULTDIV )
		{


			tok = scanner.Scan( TokenType.MULTDIV );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.MULTDIV )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.MULTDIV.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}


			ParseSign( node );
			tok = scanner.LookAhead( TokenType.MULTDIV );
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseParams( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Params ), "Params" );
		parent.Nodes.Add( node );



		tok = scanner.LookAhead( TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN );
		if ( tok.Type == TokenType.SIGN
			|| tok.Type == TokenType.METHODNAME
			|| tok.Type == TokenType.CONSTRUCTOR
			|| tok.Type == TokenType.SHEETBEGIN
			|| tok.Type == TokenType.CELLID
			|| tok.Type == TokenType.IDENTIFIER
			|| tok.Type == TokenType.EVAL
			|| tok.Type == TokenType.DATA
			|| tok.Type == TokenType.TRUE
			|| tok.Type == TokenType.FALSE
			|| tok.Type == TokenType.NULL
			|| tok.Type == TokenType.INTEGER
			|| tok.Type == TokenType.NUMBER
			|| tok.Type == TokenType.QUOTEBEGIN
			|| tok.Type == TokenType.SNGQUOTEBEGIN
			|| tok.Type == TokenType.INVALID
			|| tok.Type == TokenType.BROPEN )
		{
			ParseCompareExpr( node );
		}


		tok = scanner.LookAhead( TokenType.COMMA );
		while ( tok.Type == TokenType.COMMA )
		{


			tok = scanner.Scan( TokenType.COMMA );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.COMMA )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.COMMA.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}


			ParseCompareExpr( node );
			tok = scanner.LookAhead( TokenType.COMMA );
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseConstructor( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Constructor ), "Constructor" );
		parent.Nodes.Add( node );



		tok = scanner.Scan( TokenType.CONSTRUCTOR );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.CONSTRUCTOR )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.CONSTRUCTOR.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.LookAhead( TokenType.BROPEN, TokenType.SBROPEN );
		switch ( tok.Type )
		{
			case TokenType.BROPEN:


				tok = scanner.Scan( TokenType.BROPEN );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.BROPEN )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BROPEN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.LookAhead( TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN, TokenType.COMMA );
				if ( tok.Type == TokenType.SIGN
					|| tok.Type == TokenType.METHODNAME
					|| tok.Type == TokenType.CONSTRUCTOR
					|| tok.Type == TokenType.SHEETBEGIN
					|| tok.Type == TokenType.CELLID
					|| tok.Type == TokenType.IDENTIFIER
					|| tok.Type == TokenType.EVAL
					|| tok.Type == TokenType.DATA
					|| tok.Type == TokenType.TRUE
					|| tok.Type == TokenType.FALSE
					|| tok.Type == TokenType.NULL
					|| tok.Type == TokenType.INTEGER
					|| tok.Type == TokenType.NUMBER
					|| tok.Type == TokenType.QUOTEBEGIN
					|| tok.Type == TokenType.SNGQUOTEBEGIN
					|| tok.Type == TokenType.INVALID
					|| tok.Type == TokenType.BROPEN
					|| tok.Type == TokenType.COMMA )
				{
					ParseParams( node );
				}


				tok = scanner.Scan( TokenType.BRCLOSE );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.BRCLOSE )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BRCLOSE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.SBROPEN:


				tok = scanner.Scan( TokenType.SBROPEN );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.SBROPEN )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SBROPEN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.LookAhead( TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN, TokenType.COMMA );
				if ( tok.Type == TokenType.SIGN
					|| tok.Type == TokenType.METHODNAME
					|| tok.Type == TokenType.CONSTRUCTOR
					|| tok.Type == TokenType.SHEETBEGIN
					|| tok.Type == TokenType.CELLID
					|| tok.Type == TokenType.IDENTIFIER
					|| tok.Type == TokenType.EVAL
					|| tok.Type == TokenType.DATA
					|| tok.Type == TokenType.TRUE
					|| tok.Type == TokenType.FALSE
					|| tok.Type == TokenType.NULL
					|| tok.Type == TokenType.INTEGER
					|| tok.Type == TokenType.NUMBER
					|| tok.Type == TokenType.QUOTEBEGIN
					|| tok.Type == TokenType.SNGQUOTEBEGIN
					|| tok.Type == TokenType.INVALID
					|| tok.Type == TokenType.BROPEN
					|| tok.Type == TokenType.COMMA )
				{
					ParseParams( node );
				}


				tok = scanner.Scan( TokenType.SBRCLOSE );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.SBRCLOSE )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SBRCLOSE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			default:
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found.", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				break;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseMethod( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Method ), "Method" );
		parent.Nodes.Add( node );



		tok = scanner.Scan( TokenType.METHODNAME );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.METHODNAME )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.METHODNAME.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.Scan( TokenType.BROPEN );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.BROPEN )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BROPEN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.LookAhead( TokenType.SIGN, TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN, TokenType.COMMA );
		if ( tok.Type == TokenType.SIGN
			|| tok.Type == TokenType.METHODNAME
			|| tok.Type == TokenType.CONSTRUCTOR
			|| tok.Type == TokenType.SHEETBEGIN
			|| tok.Type == TokenType.CELLID
			|| tok.Type == TokenType.IDENTIFIER
			|| tok.Type == TokenType.EVAL
			|| tok.Type == TokenType.DATA
			|| tok.Type == TokenType.TRUE
			|| tok.Type == TokenType.FALSE
			|| tok.Type == TokenType.NULL
			|| tok.Type == TokenType.INTEGER
			|| tok.Type == TokenType.NUMBER
			|| tok.Type == TokenType.QUOTEBEGIN
			|| tok.Type == TokenType.SNGQUOTEBEGIN
			|| tok.Type == TokenType.INVALID
			|| tok.Type == TokenType.BROPEN
			|| tok.Type == TokenType.COMMA )
		{
			ParseParams( node );
		}


		tok = scanner.Scan( TokenType.BRCLOSE );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.BRCLOSE )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BRCLOSE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseString( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.String ), "String" );
		parent.Nodes.Add( node );

		tok = scanner.LookAhead( TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN );
		switch ( tok.Type )
		{
			case TokenType.QUOTEBEGIN:


				tok = scanner.Scan( TokenType.QUOTEBEGIN );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.QUOTEBEGIN )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.QUOTEBEGIN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.Scan( TokenType.QUOTED );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.QUOTED )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.QUOTED.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.Scan( TokenType.QUOTEEND );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.QUOTEEND )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.QUOTEEND.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.SNGQUOTEBEGIN:


				tok = scanner.Scan( TokenType.SNGQUOTEBEGIN );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.SNGQUOTEBEGIN )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SNGQUOTEBEGIN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.Scan( TokenType.SNGQUOTED );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.SNGQUOTED )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SNGQUOTED.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				tok = scanner.Scan( TokenType.SNGQUOTEEND );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.SNGQUOTEEND )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SNGQUOTEEND.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			default:
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found.", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				break;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseSheet( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Sheet ), "Sheet" );
		parent.Nodes.Add( node );



		tok = scanner.Scan( TokenType.SHEETBEGIN );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.SHEETBEGIN )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SHEETBEGIN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.Scan( TokenType.SNGQUOTED );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.SNGQUOTED )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SNGQUOTED.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.Scan( TokenType.SHEETEND );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.SHEETEND )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SHEETEND.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseRange( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Range ), "Range" );
		parent.Nodes.Add( node );



		tok = scanner.LookAhead( TokenType.SHEETBEGIN );
		if ( tok.Type == TokenType.SHEETBEGIN )
		{
			ParseSheet( node );
		}

		tok = scanner.Scan( TokenType.CELLID );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.CELLID )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.CELLID.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}


		tok = scanner.LookAhead( TokenType.COLON );
		if ( tok.Type == TokenType.COLON )
		{


			tok = scanner.Scan( TokenType.COLON );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.COLON )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.COLON.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}


			tok = scanner.Scan( TokenType.CELLID );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.CELLID )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.CELLID.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseIdentifier( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Identifier ), "Identifier" );
		parent.Nodes.Add( node );

		tok = scanner.Scan( TokenType.IDENTIFIER );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.IDENTIFIER )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.IDENTIFIER.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseEval( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Eval ), "Eval" );
		parent.Nodes.Add( node );

		tok = scanner.Scan( TokenType.EVAL );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.EVAL )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.EVAL.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseData( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Data ), "Data" );
		parent.Nodes.Add( node );

		tok = scanner.Scan( TokenType.DATA );
		n = node.CreateNode( tok, tok.ToString() );
		node.Token.UpdateRange( tok );
		node.Nodes.Add( n );
		if ( tok.Type != TokenType.DATA )
		{
			tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.DATA.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
			return;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseSign( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Sign ), "Sign" );
		parent.Nodes.Add( node );



		tok = scanner.LookAhead( TokenType.SIGN );
		if ( tok.Type == TokenType.SIGN )
		{
			tok = scanner.Scan( TokenType.SIGN );
			n = node.CreateNode( tok, tok.ToString() );
			node.Token.UpdateRange( tok );
			node.Nodes.Add( n );
			if ( tok.Type != TokenType.SIGN )
			{
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.SIGN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				return;
			}
		}


		tok = scanner.LookAhead( TokenType.METHODNAME, TokenType.CONSTRUCTOR, TokenType.SHEETBEGIN, TokenType.CELLID, TokenType.IDENTIFIER, TokenType.EVAL, TokenType.DATA, TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN );
		switch ( tok.Type )
		{
			case TokenType.METHODNAME:
				ParseMethod( node );
				break;
			case TokenType.CONSTRUCTOR:
				ParseConstructor( node );
				break;
			case TokenType.SHEETBEGIN:
			case TokenType.CELLID:
				ParseRange( node );
				break;
			case TokenType.IDENTIFIER:
				ParseIdentifier( node );
				break;
			case TokenType.EVAL:
				ParseEval( node );
				break;
			case TokenType.DATA:
				ParseData( node );
				break;
			case TokenType.TRUE:
			case TokenType.FALSE:
			case TokenType.NULL:
			case TokenType.INTEGER:
			case TokenType.NUMBER:
			case TokenType.QUOTEBEGIN:
			case TokenType.SNGQUOTEBEGIN:
			case TokenType.INVALID:
			case TokenType.BROPEN:
				ParseAtom( node );
				break;
			default:
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found.", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				break;
		}

		parent.Token.UpdateRange( node.Token );
	}

	private void ParseAtom( ParseNode parent )
	{
		Token tok;
		ParseNode n;
		var node = parent.CreateNode( scanner.GetToken( TokenType.Atom ), "Atom" );
		parent.Nodes.Add( node );

		tok = scanner.LookAhead( TokenType.TRUE, TokenType.FALSE, TokenType.NULL, TokenType.INTEGER, TokenType.NUMBER, TokenType.QUOTEBEGIN, TokenType.SNGQUOTEBEGIN, TokenType.INVALID, TokenType.BROPEN );
		switch ( tok.Type )
		{
			case TokenType.TRUE:
				tok = scanner.Scan( TokenType.TRUE );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.TRUE )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.TRUE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.FALSE:
				tok = scanner.Scan( TokenType.FALSE );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.FALSE )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.FALSE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.NULL:
				tok = scanner.Scan( TokenType.NULL );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.NULL )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.NULL.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.INTEGER:
				tok = scanner.Scan( TokenType.INTEGER );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.INTEGER )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.INTEGER.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.NUMBER:
				tok = scanner.Scan( TokenType.NUMBER );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.NUMBER )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.NUMBER.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.QUOTEBEGIN:
			case TokenType.SNGQUOTEBEGIN:
				ParseString( node );
				break;
			case TokenType.INVALID:
				tok = scanner.Scan( TokenType.INVALID );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.INVALID )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.INVALID.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			case TokenType.BROPEN:


				tok = scanner.Scan( TokenType.BROPEN );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.BROPEN )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BROPEN.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}


				ParseCompareExpr( node );


				tok = scanner.Scan( TokenType.BRCLOSE );
				n = node.CreateNode( tok, tok.ToString() );
				node.Token.UpdateRange( tok );
				node.Nodes.Add( n );
				if ( tok.Type != TokenType.BRCLOSE )
				{
					tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found. Expected " + TokenType.BRCLOSE.ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length ) );
					return;
				}
				break;
			default:
				tree.Errors.Add( new ParseError( "Unexpected token '" + tok.Text.Replace( "\n", "" ) + "' found.", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length ) );
				break;
		}

		parent.Token.UpdateRange( node.Token );
	}
}
