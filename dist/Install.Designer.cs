namespace Excel.AddIn.Setup
{
	partial class Install
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
            this.logo = new System.Windows.Forms.PictureBox();
            this.step1Check = new System.Windows.Forms.PictureBox();
            this.step1Label = new System.Windows.Forms.LinkLabel();
            this.cancel = new System.Windows.Forms.Button();
            this.ok = new System.Windows.Forms.Button();
            this.step2Check = new System.Windows.Forms.PictureBox();
            this.step2Label = new System.Windows.Forms.Label();
            this.step3Label = new System.Windows.Forms.Label();
            this.step3Check = new System.Windows.Forms.PictureBox();
            this.step4Label = new System.Windows.Forms.Label();
            this.step4Check = new System.Windows.Forms.PictureBox();
            this.step5Label = new System.Windows.Forms.Label();
            this.step5Check = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.logo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.step1Check)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.step2Check)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.step3Check)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.step4Check)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.step5Check)).BeginInit();
            this.SuspendLayout();
            // 
            // logo
            // 
            this.logo.Location = new System.Drawing.Point(10, 10);
            this.logo.Name = "logo";
            this.logo.Size = new System.Drawing.Size(172, 176);
            this.logo.TabIndex = 0;
            this.logo.TabStop = false;
            // 
            // step1Check
            // 
            this.step1Check.Location = new System.Drawing.Point(188, 23);
            this.step1Check.Name = "step1Check";
            this.step1Check.Size = new System.Drawing.Size(24, 24);
            this.step1Check.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.step1Check.TabIndex = 1;
            this.step1Check.TabStop = false;
            // 
            // step1Label
            // 
            this.step1Label.AutoSize = true;
            this.step1Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.step1Label.LinkArea = new System.Windows.Forms.LinkArea(0, 0);
            this.step1Label.Location = new System.Drawing.Point(218, 27);
            this.step1Label.Name = "step1Label";
            this.step1Label.Size = new System.Drawing.Size(153, 17);
            this.step1Label.TabIndex = 2;
            this.step1Label.Text = "Detect .NET 7 Runtime";
            this.step1Label.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.step1Label_LinkClicked);
            // 
            // cancel
            // 
            this.cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancel.Location = new System.Drawing.Point(675, 183);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(113, 37);
            this.cancel.TabIndex = 3;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // ok
            // 
            this.ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ok.Location = new System.Drawing.Point(556, 183);
            this.ok.Name = "ok";
            this.ok.Size = new System.Drawing.Size(113, 37);
            this.ok.TabIndex = 4;
            this.ok.Text = "Install";
            this.ok.UseVisualStyleBackColor = true;
            this.ok.Click += new System.EventHandler(this.ok_Click);
            // 
            // step2Check
            // 
            this.step2Check.Location = new System.Drawing.Point(188, 53);
            this.step2Check.Name = "step2Check";
            this.step2Check.Size = new System.Drawing.Size(24, 24);
            this.step2Check.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.step2Check.TabIndex = 5;
            this.step2Check.TabStop = false;
            // 
            // step2Label
            // 
            this.step2Label.AutoSize = true;
            this.step2Label.Enabled = false;
            this.step2Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.step2Label.Location = new System.Drawing.Point(218, 57);
            this.step2Label.Name = "step2Label";
            this.step2Label.Size = new System.Drawing.Size(159, 17);
            this.step2Label.TabIndex = 6;
            this.step2Label.Text = "Detect Excel Install Path";
            // 
            // step3Label
            // 
            this.step3Label.AutoSize = true;
            this.step3Label.Enabled = false;
            this.step3Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.step3Label.Location = new System.Drawing.Point(218, 87);
            this.step3Label.Name = "step3Label";
            this.step3Label.Size = new System.Drawing.Size(144, 17);
            this.step3Label.TabIndex = 8;
            this.step3Label.Text = "Copy Installation Files";
            // 
            // step3Check
            // 
            this.step3Check.Location = new System.Drawing.Point(188, 83);
            this.step3Check.Name = "step3Check";
            this.step3Check.Size = new System.Drawing.Size(24, 24);
            this.step3Check.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.step3Check.TabIndex = 7;
            this.step3Check.TabStop = false;
            // 
            // step4Label
            // 
            this.step4Label.AutoSize = true;
            this.step4Label.Enabled = false;
            this.step4Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.step4Label.Location = new System.Drawing.Point(218, 117);
            this.step4Label.Name = "step4Label";
            this.step4Label.Size = new System.Drawing.Size(132, 17);
            this.step4Label.TabIndex = 10;
            this.step4Label.Text = "Apply User Settings";
            // 
            // step4Check
            // 
            this.step4Check.Location = new System.Drawing.Point(188, 113);
            this.step4Check.Name = "step4Check";
            this.step4Check.Size = new System.Drawing.Size(24, 24);
            this.step4Check.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.step4Check.TabIndex = 9;
            this.step4Check.TabStop = false;
            // 
            // step5Label
            // 
            this.step5Label.AutoSize = true;
            this.step5Label.Enabled = false;
            this.step5Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.step5Label.Location = new System.Drawing.Point(218, 147);
            this.step5Label.Name = "step5Label";
            this.step5Label.Size = new System.Drawing.Size(170, 17);
            this.step5Label.TabIndex = 12;
            this.step5Label.Text = "Register add-in with Excel";
            // 
            // step5Check
            // 
            this.step5Check.Location = new System.Drawing.Point(188, 143);
            this.step5Check.Name = "step5Check";
            this.step5Check.Size = new System.Drawing.Size(24, 24);
            this.step5Check.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.step5Check.TabIndex = 11;
            this.step5Check.TabStop = false;
            // 
            // Install
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 232);
            this.Controls.Add(this.step5Label);
            this.Controls.Add(this.step5Check);
            this.Controls.Add(this.step4Label);
            this.Controls.Add(this.step4Check);
            this.Controls.Add(this.step3Label);
            this.Controls.Add(this.step3Check);
            this.Controls.Add(this.step2Label);
            this.Controls.Add(this.step2Check);
            this.Controls.Add(this.ok);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.step1Label);
            this.Controls.Add(this.step1Check);
            this.Controls.Add(this.logo);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Install";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "KAT Tools Installation";
            this.Load += new System.EventHandler(this.Install_Load);
            ((System.ComponentModel.ISupportInitialize)(this.logo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.step1Check)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.step2Check)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.step3Check)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.step4Check)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.step5Check)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.PictureBox logo;
		private System.Windows.Forms.PictureBox step1Check;
		private System.Windows.Forms.LinkLabel step1Label;
		private System.Windows.Forms.Button cancel;
		private System.Windows.Forms.Button ok;
		private System.Windows.Forms.PictureBox step2Check;
		private System.Windows.Forms.Label step2Label;
		private System.Windows.Forms.Label step3Label;
		private System.Windows.Forms.PictureBox step3Check;
		private System.Windows.Forms.Label step4Label;
		private System.Windows.Forms.PictureBox step4Check;
		private System.Windows.Forms.Label step5Label;
		private System.Windows.Forms.PictureBox step5Check;
	}
}

