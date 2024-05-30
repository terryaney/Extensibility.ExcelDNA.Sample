using ExcelDna.Integration;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Domain.Services;
using KAT.Camelot.Domain.Telemetry;
using KAT.Camelot.RBLe.Core.Calculations;
using Microsoft.Extensions.Logging;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;

class DnaCalculationService : CalculationService<DnaWorkbook, DnaWorksheet, ExcelReference, ExcelError>
{
	public DnaCalculationService( IHttpClientFactory httpClientFactory, IEmailService emailService, ITextService textService, JwtInfo? jwtInfo, ILogger<CalculationSourceContext> logger ) :
		base( httpClientFactory, emailService, textService, jwtInfo, logger ) { }
}