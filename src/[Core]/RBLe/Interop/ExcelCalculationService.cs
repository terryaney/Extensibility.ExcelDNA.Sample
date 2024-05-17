using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Domain.Services;
using KAT.Camelot.Domain.Telemetry;
using KAT.Camelot.RBLe.Core.Calculations;
using Microsoft.Extensions.Logging;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;

public class ExcelCalculationService : CalculationService<MSExcel.Workbook, MSExcel.Worksheet, MSExcel.Range, MSExcel.XlCVError>
{
	public ExcelCalculationService( IHttpClientFactory httpClientFactory, IEmailService emailService, ITextService textService, JwtInfo? jwtInfo, ILogger<CalculationSourceContext> logger ) :
		base( httpClientFactory, emailService, textService, jwtInfo, logger ) { }
}