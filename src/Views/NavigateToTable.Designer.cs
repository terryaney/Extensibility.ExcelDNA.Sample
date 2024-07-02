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
        components = new System.ComponentModel.Container();
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NavigateToTable));
        imageList = new ImageList(components);
        availableTables = new TreeView();
        colName = new ColumnHeader();
        colAddress = new ColumnHeader();
        colDescription = new ColumnHeader();
        cancel = new Button();
        ok = new Button();
        SuspendLayout();
        // 
        // imageList
        // 
        imageList.ColorDepth = ColorDepth.Depth8Bit;
        imageList.ImageStream = (ImageListStreamer)resources.GetObject("imageList.ImageStream");
        imageList.TransparentColor = Color.Transparent;
        imageList.Images.SetKeyName(0, "Table");
        imageList.Images.SetKeyName(1, "Row");
        imageList.Images.SetKeyName(2, "Sheet");
        // 
        // availableTables
        // 
        availableTables.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        availableTables.BorderStyle = BorderStyle.FixedSingle;
        availableTables.Location = new Point(12, 12);
        availableTables.Name = "availableTables";
        availableTables.ShowNodeToolTips = true;
        availableTables.Size = new Size(460, 178);
        availableTables.TabIndex = 0;
        availableTables.NodeMouseDoubleClick += AvailableTables_NodeMouseDoubleClick;
        // 
        // colName
        // 
        colName.Tag = "colName";
        colName.Text = "Name";
        colName.Width = 150;
        // 
        // colAddress
        // 
        colAddress.Tag = "colAddress";
        colAddress.Text = "Address";
        colAddress.Width = 100;
        // 
        // colDescription
        // 
        colDescription.Tag = "colDescription";
        colDescription.Text = "Description";
        colDescription.Width = 400;
        // 
        // cancel
        // 
        cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(397, 201);
        cancel.Name = "cancel";
        cancel.Size = new Size(75, 23);
        cancel.TabIndex = 2;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        cancel.Click += Cancel_Click;
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        ok.Location = new Point(300, 201);
        ok.Name = "ok";
        ok.Size = new Size(91, 23);
        ok.TabIndex = 1;
        ok.Text = "Go";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // NavigateToTable
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(8F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(484, 236);
        Controls.Add(ok);
        Controls.Add(cancel);
        Controls.Add(availableTables);
        Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point);
        MaximizeBox = false;
        MaximumSize = new Size(2800, 2000);
        MinimizeBox = false;
        MinimumSize = new Size(500, 275);
        Name = "NavigateToTable";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        Text = "Navigate To Table";
        Load += NavigateToTable_Load;
        ResumeLayout(false);
    }

    #endregion

    private System.Windows.Forms.ImageList imageList;
	private System.Windows.Forms.TreeView availableTables;
	private System.Windows.Forms.ColumnHeader colAddress;
	private System.Windows.Forms.ColumnHeader colName;
	private System.Windows.Forms.ColumnHeader colDescription;
	private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
}
