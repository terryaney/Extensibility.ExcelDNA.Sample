using System.Diagnostics;
using System.Reflection;
using System.Text.Json.Nodes;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Kat_BlastEmail( IRibbonControl _ )
	{
		var ws = application.ActiveWorksheet();

		var excelReference = ( ws.RangeOrNull( "EmailList" ) ?? application.ActiveRange() ).GetReference();

		if ( excelReference.GetValue() == ExcelEmpty.Value )
		{
			MessageBox.Show( "To perform an email blast, you must select the first email address in the list.", "BTR Email Bast", MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
			return;
		}

		var owner = new NativeWindow();
		owner.AssignHandle( new IntPtr( application.Hwnd ) );

		var attachments = ws.RangeOrNull( "Attachments" )?.Offset[ 1, 0 ];
		var attachmentList = new List<string>();

		string? attachment;
		while ( !string.IsNullOrEmpty( attachment = attachments?.GetText() ) )
		{
			attachmentList.Add( attachment );
		}

		var requestInfo = new EmailBlastRequestInfo
		{
			AddressesPerEmail = ws.RangeOrNull<int?>( "AddressesPerEmail" ) ?? 1000,
			WaitPerBatch = ws.RangeOrNull<int?>( "WaitPerBatch" ) ?? 5,
			Bcc = ws.RangeOrNull<string>( "Bcc" ),
			From = ws.RangeOrNull<string>( "From" ),
			Subject = ws.RangeOrNull<string>( "Subject" ),
			Body = ws.RangeOrNull<string>( "Body" ),
			Attachments = attachmentList.ToArray()
		};

		RunRibbonTask( async () => 
		{
			var password = await AddIn.Settings.GetClearPasswordAsync();

			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				var bottomLeft = excelReference.End( DirectionType.Down );
				var topRight = excelReference.End( DirectionType.ToRight );

				var dataRange = new ExcelReference( excelReference.RowFirst, bottomLeft.RowFirst, excelReference.ColumnFirst, topRight.ColumnFirst, excelReference.SheetId );

				var data = dataRange.GetValueArray();

				var recipients =
					data.Rows.Select( r =>
						new JsonObject().AddPropertiesWithNulls( 
							r.Select( ( c, i ) => new JsonKeyProperty( "c" + i, c?.ToString() ) )
								.ToArray()
						)
					).ToJsonArray();

				using var emailBlast = new EmailBlast( GetWindowConfiguration( nameof( EmailBlast ) ) );

				var info = emailBlast.GetInfo(
					requestInfo,
					(string)recipients[ 0 ]![ "c0" ]!,
					(string)recipients.Last()![ "c0" ]!,
					recipients.Count,
					AddIn.Settings.KatUserName,
					password,
					owner
				);

				if ( info == null ) return;
				
				SaveWindowConfiguration( nameof( EmailBlast ), info.WindowConfiguration );

				RunRibbonTask( async () =>
				{
					try
					{
						await UpdateAddInCredentialsAsync( info.UserName, info.Password );

						SetStatusBar( "Submitting Email Blast Job..." );

						var request = new EmailBlastRequest
						{
							Recipients = recipients,
							AddressesPerEmail = info.AddressesPerEmail,
							WaitPerBatch = info.WaitPerBatch,
							From = info.From,
							Audit = info.Audit,
							Bcc = info.Bcc,
							Subject = info.Subject,
							Attachments = info.Attachments
						};
						
						var response = await apiService.EmailBlastAsync( info.Body, request, info.UserName, info.Password );

						if ( response.Validations != null )
						{
							ShowValidations( response.Validations );
							return;
						}

						ExcelAsyncUtil.QueueAsMacro( () =>
						{
							MessageBox.Show( "You KAT Email Blast job was successfully submitted.  You will be notified when it completes.", "KAT Email Blast", MessageBoxButtons.OK, MessageBoxIcon.Information );

							RunRibbonTask( async () =>
							{
								var validations = await apiService.WaitForEmailBlastAsync( response.Response!, info.UserName, info.Password );

								if ( validations != null )
								{
									ShowValidations( validations );
									return;
								}

								ExcelAsyncUtil.QueueAsMacro( () => MessageBox.Show( "The KAT Email Blast job completed successfully!", "KAT Email Blast", MessageBoxButtons.OK, MessageBoxIcon.Information ) );
							} );
						} );
					}
					catch ( Exception ex )
					{
						ClearStatusBar();
						ExcelAsyncUtil.QueueAsMacro( () => {
							MessageBox.Show( "Submitting Email Blast Job FAILED. " + ex.Message, "Email Blast", MessageBoxButtons.OK, MessageBoxIcon.Error );
						} );
					}
				} );
			} );
		} );
	}

	public void Kat_ShowLog( IRibbonControl? _ )
	{
		ExcelDna.Logging.LogDisplay.Show();
		auditShowLogBadgeCount = 0;
		ribbon.InvalidateControl( "katShowDiagnosticLog" );
	}

	public void Kat_OpenHelp( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void Kat_RefreshRibbon( IRibbonControl _ ) => Application_WorkbookActivate( application.ActiveWorkbook );

	public void Kat_HelpAbout( IRibbonControl _ )
	{
		// https://stackoverflow.com/a/14498889/166231 - discusses versioning in xll
		var fvi = FileVersionInfo.GetVersionInfo( AddIn.XllName );
		var versionParts = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyInformationalVersionAttribute>().First().InformationalVersion.Split( '.' );
		var version = versionParts.Length == 3 && versionParts[ 2 ].Contains( '+' )
			? string.Join( ".", versionParts.Take( 2 ).Concat( new [] { $"{versionParts[ 2 ].Split( '+' )[ 0 ]}+{versionParts[ 2 ].Split( '+' )[ 1 ][ ..6 ]}" } ) )
			: string.Join( ".", versionParts );

		MessageBox.Show(
			$"KAT Excel Add-In: {version}{Environment.NewLine}ExcelDna.AddIn Package: {fvi.FileVersion}",
			"KAT Excel Add-In",
			MessageBoxButtons.OK,
			MessageBoxIcon.Information
		);
	}
}