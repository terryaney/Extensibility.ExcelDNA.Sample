namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

public class InteropValueException : ApplicationException
{
	public string ErrorText { get; set; }

	public InteropValueException( string message, string errorText ) : this( message, errorText, null ) { }
	public InteropValueException( string message, string errorText, Exception? innerException ) : base( message, innerException ) { ErrorText = errorText; }
}