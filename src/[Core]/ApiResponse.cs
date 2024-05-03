namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ApiResponse<T>
{
	public T? Response { get; init; }
	public ApiValidation[]? Validations { get; init; }
}