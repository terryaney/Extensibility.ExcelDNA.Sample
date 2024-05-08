using System.Text;

namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

[Serializable]
public partial class ParseTree : ParseNode
{
	public ParseErrors Errors;

	public List<Token> Skipped = new();

	public ParseTree() : base( new Token(), "ParseTree" )
	{
		Token.Type = TokenType.Start;
		Token.Text = "Root";
		Errors = new ParseErrors();
	}

	public string PrintTree()
	{
		var sb = new StringBuilder();
		var indent = 0;
		PrintNode( sb, this, indent );
		return sb.ToString();
	}

	private static void PrintNode( StringBuilder sb, ParseNode node, int indent )
	{
		var space = "".PadLeft( indent, ' ' );

		sb.Append( space );
		sb.AppendLine( node.Text );

		foreach ( var n in node.Nodes )
			PrintNode( sb, n, indent + 2 );
	}

	/// <summary>
	/// this is the entry point for executing and evaluating the parse tree.
	/// </summary>
	/// <param name="paramlist">additional optional input parameters</param>
	/// <returns>the output of the evaluation function</returns>
	public object Eval( params object[] paramlist )
	{
		return Nodes[ 0 ].Eval( this, paramlist );
	}
}

[Serializable]
public class ParseErrors : List<ParseError>
{
}

[Serializable]
public class ParseError
{
	private string message = null!;
	private int code;
	private int line;
	private int col;
	private int pos;
	private int length;

	public int Code { get { return code; } }
	public int Line { get { return line; } }
	public int Column { get { return col; } }
	public int Position { get { return pos; } }
	public int Length { get { return length; } }
	public string Message { get { return message; } }

	// just for the sake of serialization
	public ParseError()
	{
	}

	public ParseError( string message, int code, ParseNode node ) : this( message, code, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length )
	{
	}

	public ParseError( string message, int code, int line, int col, int pos, int length )
	{
		this.message = message;
		this.code = code;
		this.line = line;
		this.col = col;
		this.pos = pos;
		this.length = length;
	}
}
