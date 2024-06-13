using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Infrastructure;
using Microsoft.Extensions.Localization;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ChartHelper : IGetLocalizedString
{
	private readonly IStringLocalizer<ICamelotMarker> localizer;

	public ChartHelper( IStringLocalizer<Infrastructure.ICamelotMarker> localizer )
	{
		this.localizer = localizer;
	}

	public string? GetLocalizedString( string? key, params object[] arguments ) 
	{
		if ( key == null )
		{
			return null;
		}

		var resourceText = localizer[ key ];
		return arguments.Length > 0 ? string.Format( resourceText, arguments ) : resourceText;
	}
}