namespace KAT.Camelot.Extensibility.Excel.AddIn;

/// <summary>
/// Class used to store the result of an InputBox.Show message.
/// </summary>
public class InputBoxResult
{
	public DialogResult ReturnCode;
	public string Text = null!;
}