using System.Text.RegularExpressions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.XmlMappingExport.GeneratedParser;

public partial class Scanner
{
	public string Input = null!;
	public int StartPos = 0;
	public int EndPos = 0;
	public int CurrentLine;
	public int CurrentColumn;
	public int CurrentPosition;
	public List<Token> Skipped; // tokens that were skipped
	public Dictionary<TokenType, Regex> Patterns;

	private Token? LookAheadToken;
	private readonly List<TokenType> Tokens;
	private readonly List<TokenType> SkipList; // tokens to be skipped

	public Scanner()
	{
		Regex regex;
		Patterns = new Dictionary<TokenType, Regex>();
		Tokens = new List<TokenType>();
		LookAheadToken = null;
		Skipped = new List<Token>();

		SkipList = new List<TokenType>
		{
			TokenType.WHITESPACE
		};

#pragma warning disable SYSLIB1045 // regex

		regex = new Regex( @"(\+|-|&)", RegexOptions.Compiled );
		Patterns.Add( TokenType.PLUSMINUS, regex );
		Tokens.Add( TokenType.PLUSMINUS );

		regex = new Regex( @"\*|/", RegexOptions.Compiled );
		Patterns.Add( TokenType.MULTDIV, regex );
		Tokens.Add( TokenType.MULTDIV );

		regex = new Regex( @"\(", RegexOptions.Compiled );
		Patterns.Add( TokenType.BROPEN, regex );
		Tokens.Add( TokenType.BROPEN );

		regex = new Regex( @"\)", RegexOptions.Compiled );
		Patterns.Add( TokenType.BRCLOSE, regex );
		Tokens.Add( TokenType.BRCLOSE );

		regex = new Regex( @"\{", RegexOptions.Compiled );
		Patterns.Add( TokenType.SBROPEN, regex );
		Tokens.Add( TokenType.SBROPEN );

		regex = new Regex( @"\}", RegexOptions.Compiled );
		Patterns.Add( TokenType.SBRCLOSE, regex );
		Tokens.Add( TokenType.SBRCLOSE );

		regex = new Regex( @",", RegexOptions.Compiled );
		Patterns.Add( TokenType.COMMA, regex );
		Tokens.Add( TokenType.COMMA );

		regex = new Regex( @"\.", RegexOptions.Compiled );
		Patterns.Add( TokenType.PERIOD, regex );
		Tokens.Add( TokenType.PERIOD );

		regex = new Regex( @"\:", RegexOptions.Compiled );
		Patterns.Add( TokenType.COLON, regex );
		Tokens.Add( TokenType.COLON );

		regex = new Regex( @"'(?=(''|[^'])*'!)", RegexOptions.Compiled );
		Patterns.Add( TokenType.SHEETBEGIN, regex );
		Tokens.Add( TokenType.SHEETBEGIN );

		regex = new Regex( @"'!", RegexOptions.Compiled );
		Patterns.Add( TokenType.SHEETEND, regex );
		Tokens.Add( TokenType.SHEETEND );

		regex = new Regex( @"@?\""(?=(\""\""|[^\""])*\"")", RegexOptions.Compiled );
		Patterns.Add( TokenType.QUOTEBEGIN, regex );
		Tokens.Add( TokenType.QUOTEBEGIN );

		regex = new Regex( @"\""", RegexOptions.Compiled );
		Patterns.Add( TokenType.QUOTEEND, regex );
		Tokens.Add( TokenType.QUOTEEND );

		regex = new Regex( @"(\""\""|[^\""])*", RegexOptions.Compiled );
		Patterns.Add( TokenType.QUOTED, regex );
		Tokens.Add( TokenType.QUOTED );

		regex = new Regex( @"'(?=[^']*'[^!]?)", RegexOptions.Compiled );
		Patterns.Add( TokenType.SNGQUOTEBEGIN, regex );
		Tokens.Add( TokenType.SNGQUOTEBEGIN );

		regex = new Regex( @"'(?=[^!]?)", RegexOptions.Compiled );
		Patterns.Add( TokenType.SNGQUOTEEND, regex );
		Tokens.Add( TokenType.SNGQUOTEEND );

		regex = new Regex( @"(''|[^'])*", RegexOptions.Compiled );
		Patterns.Add( TokenType.SNGQUOTED, regex );
		Tokens.Add( TokenType.SNGQUOTED );

		regex = new Regex( @"[Tt][Rr][Uu][Ee]", RegexOptions.Compiled );
		Patterns.Add( TokenType.TRUE, regex );
		Tokens.Add( TokenType.TRUE );

		regex = new Regex( @"[Ff][Aa][Ll][Ss][Ee]", RegexOptions.Compiled );
		Patterns.Add( TokenType.FALSE, regex );
		Tokens.Add( TokenType.FALSE );

		regex = new Regex( @"[Nn][Uu][Ll][Ll](\(\))?", RegexOptions.Compiled );
		Patterns.Add( TokenType.NULL, regex );
		Tokens.Add( TokenType.NULL );

		regex = new Regex( @"\$?[a-zA-Z]+\$?[0-9]+", RegexOptions.Compiled );
		Patterns.Add( TokenType.CELLID, regex );
		Tokens.Add( TokenType.CELLID );

		regex = new Regex( @"new\s[a-zA-Z_][a-zA-Z0-9_\.]*", RegexOptions.Compiled );
		Patterns.Add( TokenType.CONSTRUCTOR, regex );
		Tokens.Add( TokenType.CONSTRUCTOR );

		regex = new Regex( @"[a-zA-Z_][a-zA-Z0-9_\.]*[a-zA-Z0-9_](?=\()", RegexOptions.Compiled );
		Patterns.Add( TokenType.METHODNAME, regex );
		Tokens.Add( TokenType.METHODNAME );

		regex = new Regex( @"[a-zA-Z_][a-zA-Z0-9_\.]*[a-zA-Z0-9_]", RegexOptions.Compiled );
		Patterns.Add( TokenType.IDENTIFIER, regex );
		Tokens.Add( TokenType.IDENTIFIER );

		regex = new Regex( @"\{Eval\s[a-zA-Z0-9_\-\.\s]*?\}", RegexOptions.Compiled );
		Patterns.Add( TokenType.EVAL, regex );
		Tokens.Add( TokenType.EVAL );

		regex = new Regex( @"\{Data\s[a-zA-Z0-9_\-\.\s]*?\}", RegexOptions.Compiled );
		Patterns.Add( TokenType.DATA, regex );
		Tokens.Add( TokenType.DATA );

		regex = new Regex( @"[a-zA-Z_][a-zA-Z0-9_\.]*(?=\s?=\s?)", RegexOptions.Compiled );
		Patterns.Add( TokenType.ASSIGNID, regex );
		Tokens.Add( TokenType.ASSIGNID );

		regex = new Regex( @"(\+|-)?[0-9]+", RegexOptions.Compiled );
		Patterns.Add( TokenType.INTEGER, regex );
		Tokens.Add( TokenType.INTEGER );

		regex = new Regex( @"(\+|-)?[0-9]*\.[0-9]+", RegexOptions.Compiled );
		Patterns.Add( TokenType.NUMBER, regex );
		Tokens.Add( TokenType.NUMBER );

		regex = new Regex( @"==|<=|>=|>|<|=|!=|<>", RegexOptions.Compiled );
		Patterns.Add( TokenType.COMPARE, regex );
		Tokens.Add( TokenType.COMPARE );

		regex = new Regex( @"=", RegexOptions.Compiled );
		Patterns.Add( TokenType.EQUALS, regex );
		Tokens.Add( TokenType.EQUALS );

		regex = new Regex( @"\+|-", RegexOptions.Compiled );
		Patterns.Add( TokenType.SIGN, regex );
		Tokens.Add( TokenType.SIGN );

		regex = new Regex( @"#VALUE!", RegexOptions.Compiled );
		Patterns.Add( TokenType.INVALID, regex );
		Tokens.Add( TokenType.INVALID );

		regex = new Regex( @"^$", RegexOptions.Compiled );
		Patterns.Add( TokenType.EOF, regex );
		Tokens.Add( TokenType.EOF );

		regex = new Regex( @"\s+", RegexOptions.Compiled );
		Patterns.Add( TokenType.WHITESPACE, regex );
		Tokens.Add( TokenType.WHITESPACE );

#pragma warning restore SYSLIB1045 // regex

	}

	public void Init( string input )
	{
		this.Input = input;
		StartPos = 0;
		EndPos = 0;
		CurrentLine = 0;
		CurrentColumn = 0;
		CurrentPosition = 0;
		LookAheadToken = null;
	}

	public Token GetToken( TokenType type )
	{
		var t = new Token( this.StartPos, this.EndPos )
		{
			Type = type
		};
		return t;
	}

	/// <summary>
	/// executes a lookahead of the next token
	/// and will advance the scan on the input string
	/// </summary>
	/// <returns></returns>
	public Token Scan( params TokenType[] expectedtokens )
	{
		var tok = LookAhead( expectedtokens ); // temporarely retrieve the lookahead
		LookAheadToken = null; // reset lookahead token, so scanning will continue
		StartPos = tok.EndPos;
		EndPos = tok.EndPos; // set the tokenizer to the new scan position
		return tok;
	}

	/// <summary>
	/// returns token with longest best match
	/// </summary>
	/// <returns></returns>
	public Token LookAhead( params TokenType[] expectedtokens )
	{
		int i;
		var startpos = StartPos;
		List<TokenType> scantokens;

		// this prevents double scanning and matching
		// increased performance
		if ( LookAheadToken != null
			&& LookAheadToken.Type != TokenType._UNDETERMINED_
			&& LookAheadToken.Type != TokenType._NONE_ ) return LookAheadToken;

		// if no scantokens specified, then scan for all of them (= backward compatible)
		if ( expectedtokens.Length == 0 )
			scantokens = Tokens;
		else
		{
			scantokens = new List<TokenType>( expectedtokens );
			scantokens.AddRange( SkipList );
		}

		Token? tok;
		do
		{

			var len = -1;
			var index = (TokenType)int.MaxValue;
			var input = Input[ startpos.. ];

			tok = new Token( startpos, EndPos );

			for ( i = 0; i < scantokens.Count; i++ )
			{
				var r = Patterns[ scantokens[ i ] ];
				var m = r.Match( input );
				if ( m.Success && m.Index == 0 && ( ( m.Length > len ) || ( scantokens[ i ] < index && m.Length == len ) ) )
				{
					len = m.Length;
					index = scantokens[ i ];
				}
			}

			if ( index >= 0 && len >= 0 )
			{
				tok.EndPos = startpos + len;
				tok.Text = Input.Substring( tok.StartPos, len );
				tok.Type = index;
			}
			else if ( tok.StartPos < tok.EndPos - 1 )
			{
				tok.Text = Input.Substring( tok.StartPos, 1 );
			}

			if ( SkipList.Contains( tok.Type ) )
			{
				startpos = tok.EndPos;
				Skipped.Add( tok );
			}
			else
			{
				// only assign to non-skipped tokens
				tok.Skipped = Skipped; // assign prior skips to this token
				Skipped = new List<Token>(); //reset skips
			}
		}
		while ( SkipList.Contains( tok.Type ) );

		LookAheadToken = tok;
		return tok;
	}
}