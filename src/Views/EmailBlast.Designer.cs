namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class EmailBlast
{
	/// <summary>
	/// Required designer variable.
	/// </summary>
	private System.ComponentModel.IContainer components = null;

	/// <summary>
	/// Clean up any resources being used.
	/// </summary>
	/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	protected override void Dispose( bool disposing )
	{
		if ( disposing && ( components != null ) )
		{
			components.Dispose();
		}
		base.Dispose( disposing );
	}

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        cancel = new Button();
        ok = new Button();
        errorProvider = new ErrorProvider(components);
        label1 = new Label();
        emailAddress = new TextBox();
        label2 = new Label();
        password = new TextBox();
        label3 = new Label();
        audit = new TextBox();
        label5 = new Label();
        bcc = new TextBox();
        label6 = new Label();
        from = new TextBox();
        label7 = new Label();
        waitMinutes = new TextBox();
        label9 = new Label();
        addressPerEmail = new TextBox();
        label10 = new Label();
        label11 = new Label();
        utilityTabsDivider = new Label();
        subject = new TextBox();
        label12 = new Label();
        label13 = new Label();
        openFileDialog = new OpenFileDialog();
        bodyFileNameSelect = new Button();
        body = new TextBox();
        label14 = new Label();
        attachments = new ListBox();
        add = new Button();
        remove = new Button();
        label4 = new Label();
        lSendDetails = new Label();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // cancel
        // 
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(317, 522);
        cancel.Margin = new Padding(4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 29;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // ok
        // 
        ok.Location = new Point(181, 522);
        ok.Margin = new Padding(4);
        ok.Name = "ok";
        ok.Size = new Size(130, 26);
        ok.TabIndex = 28;
        ok.Text = "Send 1000 Emails";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point);
        label1.Location = new Point(14, 10);
        label1.Margin = new Padding(4, 0, 4, 0);
        label1.Name = "label1";
        label1.Size = new Size(156, 13);
        label1.TabIndex = 0;
        label1.Text = "Authentication Information";
        // 
        // emailAddress
        // 
        emailAddress.Location = new Point(18, 51);
        emailAddress.Margin = new Padding(4);
        emailAddress.Name = "emailAddress";
        emailAddress.Size = new Size(201, 23);
        emailAddress.TabIndex = 2;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(14, 32);
        label2.Margin = new Padding(4, 0, 4, 0);
        label2.Name = "label2";
        label2.Size = new Size(65, 15);
        label2.TabIndex = 1;
        label2.Text = "&User Name";
        // 
        // password
        // 
        password.Location = new Point(225, 51);
        password.Margin = new Padding(4);
        password.Name = "password";
        password.PasswordChar = '*';
        password.Size = new Size(182, 23);
        password.TabIndex = 4;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(225, 32);
        label3.Margin = new Padding(4, 0, 4, 0);
        label3.Name = "label3";
        label3.Size = new Size(60, 15);
        label3.TabIndex = 3;
        label3.Text = "&Password:";
        // 
        // audit
        // 
        audit.Location = new Point(18, 180);
        audit.Margin = new Padding(4);
        audit.Name = "audit";
        audit.Size = new Size(389, 23);
        audit.TabIndex = 11;
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(14, 161);
        label5.Margin = new Padding(4, 0, 4, 0);
        label5.Name = "label5";
        label5.Size = new Size(135, 15);
        label5.TabIndex = 10;
        label5.Text = "&Audit Email Address(es):";
        // 
        // bcc
        // 
        bcc.Location = new Point(226, 223);
        bcc.Margin = new Padding(4);
        bcc.Name = "bcc";
        bcc.Size = new Size(181, 23);
        bcc.TabIndex = 15;
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(226, 204);
        label6.Margin = new Padding(4, 0, 4, 0);
        label6.Name = "label6";
        label6.Size = new Size(129, 15);
        label6.TabIndex = 14;
        label6.Text = "&BCC Email Address(es):";
        // 
        // from
        // 
        from.Location = new Point(18, 223);
        from.Margin = new Padding(4);
        from.Name = "from";
        from.Size = new Size(200, 23);
        from.TabIndex = 13;
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(14, 204);
        label7.Margin = new Padding(4, 0, 4, 0);
        label7.Name = "label7";
        label7.Size = new Size(38, 15);
        label7.TabIndex = 12;
        label7.Text = "&From:";
        // 
        // waitMinutes
        // 
        waitMinutes.Location = new Point(225, 130);
        waitMinutes.Margin = new Padding(4);
        waitMinutes.Name = "waitMinutes";
        waitMinutes.Size = new Size(182, 23);
        waitMinutes.TabIndex = 9;
        waitMinutes.Text = "5";
        // 
        // label9
        // 
        label9.AutoSize = true;
        label9.Location = new Point(226, 112);
        label9.Margin = new Padding(4, 0, 4, 0);
        label9.Name = "label9";
        label9.Size = new Size(80, 15);
        label9.TabIndex = 8;
        label9.Text = "&Wait Minutes:";
        // 
        // addressPerEmail
        // 
        addressPerEmail.Location = new Point(18, 130);
        addressPerEmail.Margin = new Padding(4);
        addressPerEmail.Name = "addressPerEmail";
        addressPerEmail.Size = new Size(201, 23);
        addressPerEmail.TabIndex = 7;
        addressPerEmail.Text = "1000";
        // 
        // label10
        // 
        label10.AutoSize = true;
        label10.Location = new Point(14, 112);
        label10.Margin = new Padding(4, 0, 4, 0);
        label10.Name = "label10";
        label10.Size = new Size(115, 15);
        label10.TabIndex = 6;
        label10.Text = "Addresses per Email:";
        // 
        // label11
        // 
        label11.AutoSize = true;
        label11.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point);
        label11.Location = new Point(14, 87);
        label11.Margin = new Padding(4, 0, 4, 0);
        label11.Name = "label11";
        label11.Size = new Size(106, 13);
        label11.TabIndex = 5;
        label11.Text = "Job Configuration";
        // 
        // utilityTabsDivider
        // 
        utilityTabsDivider.BorderStyle = BorderStyle.Fixed3D;
        utilityTabsDivider.Location = new Point(13, 443);
        utilityTabsDivider.Margin = new Padding(4, 0, 4, 0);
        utilityTabsDivider.Name = "utilityTabsDivider";
        utilityTabsDivider.Size = new Size(394, 2);
        utilityTabsDivider.TabIndex = 25;
        // 
        // subject
        // 
        subject.Location = new Point(18, 266);
        subject.Margin = new Padding(4);
        subject.Name = "subject";
        subject.Size = new Size(388, 23);
        subject.TabIndex = 17;
        subject.TextChanged += Subject_TextChanged;
        // 
        // label12
        // 
        label12.AutoSize = true;
        label12.Location = new Point(14, 247);
        label12.Margin = new Padding(4, 0, 4, 0);
        label12.Name = "label12";
        label12.Size = new Size(49, 15);
        label12.TabIndex = 16;
        label12.Text = "Sub&ject:";
        // 
        // label13
        // 
        label13.AutoSize = true;
        label13.Location = new Point(14, 290);
        label13.Margin = new Padding(4, 0, 4, 0);
        label13.Name = "label13";
        label13.Size = new Size(37, 15);
        label13.TabIndex = 18;
        label13.Text = "Bod&y:";
        // 
        // openFileDialog
        // 
        openFileDialog.Filter = "*.*|All Files";
        // 
        // bodyFileNameSelect
        // 
        bodyFileNameSelect.Location = new Point(367, 308);
        bodyFileNameSelect.Margin = new Padding(4);
        bodyFileNameSelect.Name = "bodyFileNameSelect";
        bodyFileNameSelect.Size = new Size(39, 21);
        bodyFileNameSelect.TabIndex = 20;
        bodyFileNameSelect.Text = "...";
        bodyFileNameSelect.UseVisualStyleBackColor = true;
        bodyFileNameSelect.Click += BodyFileNameSelect_Click;
        // 
        // body
        // 
        body.Location = new Point(18, 308);
        body.Margin = new Padding(4);
        body.Name = "body";
        body.Size = new Size(342, 23);
        body.TabIndex = 19;
        body.TextChanged += Body_TextChanged;
        // 
        // label14
        // 
        label14.AutoSize = true;
        label14.Location = new Point(14, 332);
        label14.Margin = new Padding(4, 0, 4, 0);
        label14.Name = "label14";
        label14.Size = new Size(78, 15);
        label14.TabIndex = 21;
        label14.Text = "A&ttachments:";
        // 
        // attachments
        // 
        attachments.DrawMode = DrawMode.OwnerDrawFixed;
        attachments.FormattingEnabled = true;
        attachments.Location = new Point(18, 350);
        attachments.Margin = new Padding(4);
        attachments.Name = "attachments";
        attachments.Size = new Size(309, 82);
        attachments.TabIndex = 22;
        attachments.DrawItem += Attachments_DrawItem;
        // 
        // add
        // 
        add.Location = new Point(333, 350);
        add.Margin = new Padding(4);
        add.Name = "add";
        add.Size = new Size(72, 26);
        add.TabIndex = 23;
        add.Text = "&Add";
        add.UseVisualStyleBackColor = true;
        add.Click += Add_Click;
        // 
        // remove
        // 
        remove.Location = new Point(333, 384);
        remove.Margin = new Padding(4);
        remove.Name = "remove";
        remove.Size = new Size(72, 26);
        remove.TabIndex = 24;
        remove.Text = "&Remove";
        remove.UseVisualStyleBackColor = true;
        remove.Click += Remove_Click;
        // 
        // label4
        // 
        label4.BorderStyle = BorderStyle.Fixed3D;
        label4.Location = new Point(13, 506);
        label4.Margin = new Padding(4, 0, 4, 0);
        label4.Name = "label4";
        label4.Size = new Size(394, 2);
        label4.TabIndex = 27;
        // 
        // lSendDetails
        // 
        lSendDetails.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point);
        lSendDetails.ForeColor = SystemColors.ControlText;
        lSendDetails.Location = new Point(18, 450);
        lSendDetails.Margin = new Padding(4, 0, 4, 0);
        lSendDetails.Name = "lSendDetails";
        lSendDetails.Size = new Size(387, 51);
        lSendDetails.TabIndex = 26;
        lSendDetails.Text = "Are you sure you want to send the current email to 1000 email address(es) starting with 'terry.aney@conduent.com' and ending with 'terry.aney@conduent.com'?";
        // 
        // EmailBlast
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(429, 561);
        Controls.Add(lSendDetails);
        Controls.Add(label4);
        Controls.Add(bodyFileNameSelect);
        Controls.Add(body);
        Controls.Add(remove);
        Controls.Add(add);
        Controls.Add(attachments);
        Controls.Add(label14);
        Controls.Add(label13);
        Controls.Add(subject);
        Controls.Add(label12);
        Controls.Add(utilityTabsDivider);
        Controls.Add(label11);
        Controls.Add(waitMinutes);
        Controls.Add(label9);
        Controls.Add(addressPerEmail);
        Controls.Add(label10);
        Controls.Add(from);
        Controls.Add(label7);
        Controls.Add(bcc);
        Controls.Add(label6);
        Controls.Add(audit);
        Controls.Add(label5);
        Controls.Add(password);
        Controls.Add(label3);
        Controls.Add(emailAddress);
        Controls.Add(label2);
        Controls.Add(label1);
        Controls.Add(ok);
        Controls.Add(cancel);
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(445, 600);
        MinimizeBox = false;
        MinimumSize = new Size(445, 600);
        Name = "EmailBlast";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Hide;
        Text = "KAT Email Blast Utility";
        Load += EmailBlast_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.Label label11;
	private System.Windows.Forms.TextBox waitMinutes;
	private System.Windows.Forms.Label label9;
	private System.Windows.Forms.TextBox addressPerEmail;
	private System.Windows.Forms.Label label10;
	private System.Windows.Forms.TextBox from;
	private System.Windows.Forms.Label label7;
	private System.Windows.Forms.TextBox bcc;
	private System.Windows.Forms.Label label6;
	private System.Windows.Forms.TextBox audit;
	private System.Windows.Forms.Label label5;
	private System.Windows.Forms.Label label2;
	private System.Windows.Forms.TextBox emailAddress;
	private System.Windows.Forms.Label label3;
	private System.Windows.Forms.TextBox password;
	private System.Windows.Forms.Label utilityTabsDivider;
	private System.Windows.Forms.TextBox subject;
	private System.Windows.Forms.Label label12;
	private System.Windows.Forms.Label label13;
	private System.Windows.Forms.OpenFileDialog openFileDialog;
	private Button bodyFileNameSelect;
	private TextBox body;
	private Button remove;
	private Button add;
	private ListBox attachments;
	private Label label14;
	private Label lSendDetails;
	private Label label4;
}