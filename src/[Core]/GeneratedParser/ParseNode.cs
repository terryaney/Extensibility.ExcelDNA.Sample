using System.Xml.Serialization;

namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

[Serializable]
[XmlInclude( typeof( ParseTree ) )]
public partial class ParseNode
{
	protected string text;
	protected List<ParseNode> nodes;

	public List<ParseNode> Nodes { get { return nodes; } }

	[XmlIgnore] // avoid circular references when serializing
	public ParseNode? Parent;
	public Token Token; // the token/rule

	[XmlIgnore] // skip redundant text (is part of Token)
	public string Text
	{ // text to display in parse tree
		get { return text; }
		set { text = value; }
	}

	public virtual ParseNode CreateNode( Token token, string text )
	{
		var node = new ParseNode( token, text )
		{
			Parent = this
		};
		return node;
	}

	protected ParseNode( Token token, string text )
	{
		this.Token = token;
		this.text = text;
		this.nodes = new List<ParseNode>();
	}

	protected object? GetValue( ParseTree tree, TokenType type, int index ) => GetValue( tree, type, ref index );

	protected object? GetValue( ParseTree tree, TokenType type, ref int index )
	{
		object? o = null;
		if ( index < 0 ) return o;

		// left to right
		foreach ( var node in nodes )
		{
			if ( node.Token.Type == type )
			{
				index--;
				if ( index < 0 )
				{
					o = node.Eval( tree );
					break;
				}
			}
		}
		return o;
	}

	/// <summary>
	/// this implements the evaluation functionality, cannot be used directly
	/// </summary>
	/// <param name="tree">the parsetree itself</param>
	/// <param name="paramlist">optional input parameters</param>
	/// <returns>a partial result of the evaluation</returns>
	internal object Eval( ParseTree tree, params object[] paramlist )
	{
		var value = Token.Type switch
		{
			TokenType.Start => EvalStart( tree, paramlist ),
			TokenType.Assignment => EvalAssignment( tree, paramlist ),
			TokenType.CompareExpr => EvalCompareExpr( tree, paramlist ),
			TokenType.AddExpr => EvalAddExpr( tree, paramlist ),
			TokenType.MultExpr => EvalMultExpr( tree, paramlist ),
			TokenType.Params => EvalParams( tree, paramlist ),
			TokenType.Constructor => EvalConstructor( tree, paramlist ),
			TokenType.Method => EvalMethod( tree, paramlist ),
			TokenType.String => EvalString( tree, paramlist ),
			TokenType.Sheet => EvalSheet( tree, paramlist ),
			TokenType.Range => EvalRange( tree, paramlist ),
			TokenType.Identifier => EvalIdentifier( tree, paramlist ),
			TokenType.Eval => EvalEval( tree, paramlist ),
			TokenType.Data => EvalData( tree, paramlist ),
			TokenType.Sign => EvalSign( tree, paramlist ),
			TokenType.Atom => EvalAtom( tree, paramlist ),
			_ => Token.Text,
		};
		return value;
	}

	protected virtual object EvalStart( ParseTree tree, params object[] paramlist ) => "Could not interpret input; no semantics implemented.";
	protected virtual object EvalAssignment( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalCompareExpr( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalAddExpr( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalMultExpr( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalParams( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalConstructor( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalMethod( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalString( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalSheet( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalRange( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalIdentifier( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalEval( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalData( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalSign( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
	protected virtual object EvalAtom( ParseTree tree, params object[] paramlist ) => throw new NotImplementedException();
}