using System.Diagnostics;
using System.Text.Json.Nodes;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class EmailBlast : Form
{
	private readonly JsonObject windowConfiguration;
	private string? content;

	public EmailBlast( JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	private void EmailBlast_Load( object sender, EventArgs e )
	{
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	public EmailBlastInfo? GetInfo( EmailBlastRequestInfo requestInfo, string firstEmail, string lastEmail, int totalEmails, string? userName, string? password, NativeWindow owner )
	{
		emailAddress.Text = userName;
		this.password.Text = password;
		
		addressPerEmail.Text = requestInfo.AddressesPerEmail.ToString();
		waitMinutes.Text = requestInfo.WaitPerBatch.ToString();
		bcc.Text = requestInfo.Bcc;
		from.Text = requestInfo.From;
		subject.Text = requestInfo.Subject;
		body.Text = requestInfo.Body;
		attachments.Items.AddRange( requestInfo.Attachments.Select( a => new Attachment( Color.Black, a ) ).ToArray() );

		ok.Text = $"Send {totalEmails} Emails";
		lSendDetails.Text = $"Email blast will send email to {totalEmails} email address(es) starting with '{firstEmail}' and ending with '{lastEmail}';";

		var dialogResult = ShowDialog( owner );

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;
		windowConfiguration[ "height" ] = Size.Height;
		windowConfiguration[ "width" ] = Size.Width;

		return new()
		{
			WindowConfiguration = windowConfiguration,

			UserName = emailAddress.Text,
			Password = this.password.Text,

			AddressesPerEmail = int.Parse( addressPerEmail.Text ),
			WaitPerBatch = int.Parse( waitMinutes.Text ),
			Bcc = bcc.Text,
			From = from.Text,
			Subject = subject.Text,
			Body = body.Text,
			Audit = audit.Enabled ? audit.Text : null,
			Attachments = 
				attachments.Items.Cast<Attachment>().Select( a => new EmailBlastAttachment
				{
					Id = Guid.NewGuid().ToString(),
					Name = a.File
				}  ).ToArray()
		};
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		var isValid = true;
		errorProvider.Clear();

		if ( string.IsNullOrEmpty( emailAddress.Text ) )
		{
			errorProvider.SetError( emailAddress, "You must provide an Email Address to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( password.Text ) )
		{
			errorProvider.SetError( password, "You must provide a Password to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( addressPerEmail.Text ) )
		{
			errorProvider.SetError( addressPerEmail, "You must provide the number of addresses per email to continue." );
			isValid = false;
		}
		else if ( !int.TryParse( addressPerEmail.Text, out var _ ) )
		{
			errorProvider.SetError( addressPerEmail, "You must provide a numeric value to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( waitMinutes.Text ) )
		{
			errorProvider.SetError( waitMinutes, "You must provide the number of wait minutes between each email to continue." );
			isValid = false;
		}
		else if ( !int.TryParse( waitMinutes.Text, out var _ ) )
		{
			errorProvider.SetError( waitMinutes, "You must provide a numeric value to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( from.Text ) )
		{
			errorProvider.SetError( from, "You must provide a From email address to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( subject.Text ) )
		{
			errorProvider.SetError( subject, "You must provide an email subject to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( body.Text ) || !File.Exists( body.Text ) )
		{
			errorProvider.SetError( body, "You must provide an valid file that contains the html content for the email body to continue." );
			isValid = false;
		}
		else
		{
			var bodyContent = File.ReadAllText( body.Text );
			if ( bodyContent.IndexOf( "src=\"cid:" ) > -1 )
			{
				errorProvider.SetError( body, "Your email body cannot contain legacy <img src=\"cid:{id}\"/> format.  Update your .html file to correctly reference real files and resubmit your job." );
			}
		}

		var redraw = false;
		foreach ( Attachment item in attachments.Items )
		{
			item.ItemColor = Color.Black;
			var fileName = item.File.Split( '|' )[ 0 ];

			if ( !File.Exists( fileName ) )
			{
				item.ItemColor = Color.Red;
				redraw = true;
			}
		}

		if ( redraw )
		{
			errorProvider.SetError( add, "Some of the attachment files provided do not exist.  Check the path and try again." );
			isValid = false;
			attachments.Refresh();
		}

		if ( isValid )
		{
			DialogResult = DialogResult.OK;
			Close();
		}
	}

	private void InvalidateAudit()
	{
		// Not sure why I disabled if {c1} there...feel like I can still do so...
		// audit.Enabled = !( ( content ?? "" ).Contains( "{c1}" ) || subject.Text.Contains( "{c1}" ) );
		audit.Enabled = true;
	}

	private void Subject_TextChanged( object sender, EventArgs e ) => InvalidateAudit();
	
	private void Body_TextChanged( object sender, EventArgs e )
	{
		content = null;

		if ( File.Exists( body.Text ) )
		{
			content = File.ReadAllText( body.Text );
		}

		InvalidateAudit();
	}

	private void BodyFileNameSelect_Click( object sender, EventArgs e )
	{
		openFileDialog.Title = "Select Email Body File";
		openFileDialog.Multiselect = false;
		openFileDialog.Filter = "Html Files|*.htm;*.html";
		openFileDialog.FilterIndex = 0;

		if ( openFileDialog.ShowDialog() == DialogResult.OK )
		{
			body.Text = openFileDialog.FileName;
		}
	}

	private void Add_Click( object sender, EventArgs e )
	{
		openFileDialog.Title = "Select Email Attachment";
		openFileDialog.Multiselect = true;
		openFileDialog.Filter = "All Files|*.*";
		openFileDialog.FilterIndex = 0;

		if ( openFileDialog.ShowDialog() == DialogResult.OK )
		{
			foreach ( var f in openFileDialog.FileNames )
			{
				attachments.Items.Add( new Attachment( Color.Black, f ) );
			}
		}
	}

	private void Remove_Click( object sender, EventArgs e )
	{
		for ( var i = attachments.SelectedItems.Count - 1; i >= 0; i-- )
		{
			attachments.Items.Remove( attachments.SelectedItems[ i ]! );
		}
	}

	private void Attachments_DrawItem( object sender, DrawItemEventArgs e )
	{
		e.DrawBackground();
		e.DrawFocusRectangle();

		if ( e.Index == -1 ) return;

		var item = ( attachments.Items[ e.Index ] as Attachment )!;

		e.Graphics.DrawString(
			item.File,
			item.ItemColor == Color.Red
				? new Font( attachments.Font.FontFamily, attachments.Font.Size, FontStyle.Bold )
				: attachments.Font,
			new SolidBrush( item.ItemColor ),
			e.Bounds
		);
	}
}
