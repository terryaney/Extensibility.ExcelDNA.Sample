namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class NavigateToTable
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
		this.availableTables = new System.Windows.Forms.ListView();
		this.colName = ( (System.Windows.Forms.ColumnHeader)( new System.Windows.Forms.ColumnHeader() ) );
		this.colAddress = ( (System.Windows.Forms.ColumnHeader)( new System.Windows.Forms.ColumnHeader() ) );
		this.colDescription = ( (System.Windows.Forms.ColumnHeader)( new System.Windows.Forms.ColumnHeader() ) );
		this.cancel = new System.Windows.Forms.Button();
		this.ok = new System.Windows.Forms.Button();
		this.SuspendLayout();
		// 
		// availableTables
		// 
		this.availableTables.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom )
		| System.Windows.Forms.AnchorStyles.Left )
		| System.Windows.Forms.AnchorStyles.Right ) ) );
		this.availableTables.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
		this.availableTables.Columns.AddRange( new System.Windows.Forms.ColumnHeader[] {
			this.colName,
			this.colAddress,
			this.colDescription} );
		this.availableTables.FullRowSelect = true;
		this.availableTables.HideSelection = false;
		this.availableTables.Location = new System.Drawing.Point( 12, 12 );
		this.availableTables.MultiSelect = false;
		this.availableTables.Name = "availableTables";
		this.availableTables.Size = new System.Drawing.Size( 460, 178 );
		this.availableTables.TabIndex = 30;
		this.availableTables.UseCompatibleStateImageBehavior = false;
		this.availableTables.View = System.Windows.Forms.View.Details;
		this.availableTables.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler( this.AvailableTables_ColumnClick );
		this.availableTables.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler( this.AvailableTables_MouseDoubleClick );
		// 
		// colName
		// 
		this.colName.Tag = "colName";
		this.colName.Text = "Name";
		this.colName.Width = 150;
		// 
		// colAddress
		// 
		this.colAddress.Tag = "colAddress";
		this.colAddress.Text = "Address";
		this.colAddress.Width = 100;
		// 
		// colDescription
		// 
		this.colDescription.Tag = "colDescription";
		this.colDescription.Text = "Description";
		this.colDescription.Width = 400;
		// 
		// cancel
		// 
		this.cancel.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.cancel.Location = new System.Drawing.Point( 397, 201 );
		this.cancel.Name = "cancel";
		this.cancel.Size = new System.Drawing.Size( 75, 23 );
		this.cancel.TabIndex = 31;
		this.cancel.Text = "Cancel";
		this.cancel.UseVisualStyleBackColor = true;
		this.cancel.Click += new System.EventHandler( this.Cancel_Click );
		// 
		// ok
		// 
		this.ok.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.ok.Location = new System.Drawing.Point( 300, 201 );
		this.ok.Name = "ok";
		this.ok.Size = new System.Drawing.Size( 91, 23 );
		this.ok.TabIndex = 32;
		this.ok.Text = "Go";
		this.ok.UseVisualStyleBackColor = true;
		this.ok.Click += new System.EventHandler( this.Ok_Click );
		// 
		// NavigateToTable
		// 
		this.AcceptButton = this.ok;
		this.AutoScaleDimensions = new System.Drawing.SizeF( 8F, 17F );
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.cancel;
		this.ClientSize = new System.Drawing.Size( 484, 236 );
		this.Controls.Add( this.ok );
		this.Controls.Add( this.cancel );
		this.Controls.Add( this.availableTables );
		this.Font = new System.Drawing.Font( "Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ( (byte)( 0 ) ) );
		this.MaximizeBox = false;
		this.MaximumSize = new System.Drawing.Size( 2800, 2000 );
		this.MinimizeBox = false;
		this.MinimumSize = new System.Drawing.Size( 500, 275 );
		this.Name = "NavigateToTable";
		this.ShowIcon = false;
		this.ShowInTaskbar = false;
		this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
		this.Text = "Navigate To Table";
		this.ResumeLayout( false );
	}

	#endregion

	private System.Windows.Forms.ListView availableTables;
	private System.Windows.Forms.ColumnHeader colAddress;
	private System.Windows.Forms.ColumnHeader colName;
	private System.Windows.Forms.ColumnHeader colDescription;
	private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
}
