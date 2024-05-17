using System.Xml.Serialization;

namespace KAT.Camelot.Extensibility.Excel.AddIn.XmlMappingExport.GeneratedParser;

public class Token
{
	private int startpos;
	private int endpos;
	private string text = null!;
	private object? value;

	// contains all prior skipped symbols
	private List<Token> skipped = new();

	public int StartPos
	{
		get { return startpos; }
		set { startpos = value; }
	}

	public int Length => endpos - startpos;

	public int EndPos
	{
		get { return endpos; }
		set { endpos = value; }
	}

	public string Text
	{
		get { return text; }
		set { text = value; }
	}

	public List<Token> Skipped
	{
		get { return skipped; }
		set { skipped = value; }
	}
	public object? Value
	{
		get { return value; }
		set { this.value = value; }
	}

	[XmlAttribute]
	public TokenType Type;

	public Token() : this( 0, 0 ) { }

	public Token( int start, int end )
	{
		Type = TokenType._UNDETERMINED_;
		startpos = start;
		endpos = end;
		Text = ""; // must initialize with empty string, may cause null reference exceptions otherwise
		Value = null;
	}

	public void UpdateRange( Token token )
	{
		if ( token.StartPos < startpos ) startpos = token.StartPos;
		if ( token.EndPos > endpos ) endpos = token.EndPos;
	}

	public override string ToString() => Text != null ? Type.ToString() + " '" + Text + "'" :  Type.ToString();	
}