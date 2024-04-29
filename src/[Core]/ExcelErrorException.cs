namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ExcelErrorException : ApplicationException
{
	public string ErrorText { get; set; }

	public ExcelErrorException( string message, string errorText ) : this( message, errorText, null ) { }
	public ExcelErrorException( string message, string errorText, Exception? innerException ) : base( message, innerException ) { ErrorText = errorText; }
}