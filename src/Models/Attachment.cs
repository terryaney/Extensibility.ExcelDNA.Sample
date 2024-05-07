namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class Attachment
{
	public Attachment( Color c, string f )
	{
		ItemColor = c;
		File = f;
	}
	public Color ItemColor { get; set; }
	public string File { get; set; }
}