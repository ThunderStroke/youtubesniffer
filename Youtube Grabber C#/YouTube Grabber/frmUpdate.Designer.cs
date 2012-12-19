namespace YouTube_Grabber
{
    partial class frmUpdate
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmUpdate));
            this.picProgressBar = new System.Windows.Forms.PictureBox();
            this.btnUpdateNow = new System.Windows.Forms.Button();
            this.lblNoUpdate = new System.Windows.Forms.Label();
            this.picClose = new System.Windows.Forms.PictureBox();
            this.lblChecking = new System.Windows.Forms.Label();
            this.lblCritical = new System.Windows.Forms.LinkLabel();
            this.lblUpdating = new System.Windows.Forms.Label();
            this.lblHistory = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.picProgressBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picClose)).BeginInit();
            this.SuspendLayout();
            // 
            // picProgressBar
            // 
            this.picProgressBar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picProgressBar.Location = new System.Drawing.Point(12, 136);
            this.picProgressBar.Name = "picProgressBar";
            this.picProgressBar.Size = new System.Drawing.Size(253, 44);
            this.picProgressBar.TabIndex = 2;
            this.picProgressBar.TabStop = false;
            this.picProgressBar.Visible = false;
            // 
            // btnUpdateNow
            // 
            this.btnUpdateNow.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdateNow.Location = new System.Drawing.Point(12, 44);
            this.btnUpdateNow.Name = "btnUpdateNow";
            this.btnUpdateNow.Size = new System.Drawing.Size(253, 86);
            this.btnUpdateNow.TabIndex = 4;
            this.btnUpdateNow.Text = "New Version Found!\r\nUpdate Now";
            this.btnUpdateNow.UseVisualStyleBackColor = true;
            this.btnUpdateNow.Visible = false;
            this.btnUpdateNow.Click += new System.EventHandler(this.btnUpdateNow_Click);
            // 
            // lblNoUpdate
            // 
            this.lblNoUpdate.AutoSize = true;
            this.lblNoUpdate.BackColor = System.Drawing.Color.Transparent;
            this.lblNoUpdate.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNoUpdate.Location = new System.Drawing.Point(19, 20);
            this.lblNoUpdate.Name = "lblNoUpdate";
            this.lblNoUpdate.Size = new System.Drawing.Size(246, 21);
            this.lblNoUpdate.TabIndex = 3;
            this.lblNoUpdate.Text = "Your Version Is The Latest Version";
            this.lblNoUpdate.Visible = false;
            // 
            // picClose
            // 
            this.picClose.Cursor = System.Windows.Forms.Cursors.Hand;
            this.picClose.Location = new System.Drawing.Point(267, 4);
            this.picClose.Name = "picClose";
            this.picClose.Size = new System.Drawing.Size(8, 8);
            this.picClose.TabIndex = 5;
            this.picClose.TabStop = false;
            this.picClose.Click += new System.EventHandler(this.picClose_Click);
            this.picClose.MouseEnter += new System.EventHandler(this.picClose_MouseEnter);
            this.picClose.MouseLeave += new System.EventHandler(this.picClose_MouseLeave);
            // 
            // lblChecking
            // 
            this.lblChecking.AutoSize = true;
            this.lblChecking.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChecking.Location = new System.Drawing.Point(8, 1);
            this.lblChecking.Name = "lblChecking";
            this.lblChecking.Size = new System.Drawing.Size(78, 20);
            this.lblChecking.TabIndex = 6;
            this.lblChecking.Text = "Checking...";
            // 
            // lblCritical
            // 
            this.lblCritical.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCritical.Location = new System.Drawing.Point(19, 53);
            this.lblCritical.Name = "lblCritical";
            this.lblCritical.Size = new System.Drawing.Size(236, 44);
            this.lblCritical.TabIndex = 7;
            this.lblCritical.TabStop = true;
            this.lblCritical.Text = "This is a critical update, please download manually (Click me)";
            this.lblCritical.Visible = false;
            this.lblCritical.Click += new System.EventHandler(this.lblCritical_Click);
            // 
            // lblUpdating
            // 
            this.lblUpdating.AutoSize = true;
            this.lblUpdating.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUpdating.Location = new System.Drawing.Point(92, 1);
            this.lblUpdating.Name = "lblUpdating";
            this.lblUpdating.Size = new System.Drawing.Size(80, 20);
            this.lblUpdating.TabIndex = 8;
            this.lblUpdating.Text = "Updating...";
            this.lblUpdating.Visible = false;
            // 
            // lblHistory
            // 
            this.lblHistory.AutoSize = true;
            this.lblHistory.Location = new System.Drawing.Point(12, 133);
            this.lblHistory.Name = "lblHistory";
            this.lblHistory.Size = new System.Drawing.Size(112, 13);
            this.lblHistory.TabIndex = 9;
            this.lblHistory.TabStop = true;
            this.lblHistory.Text = "View Update History";
            this.lblHistory.Visible = false;
            this.lblHistory.Click += new System.EventHandler(this.lblHistory_Click);
            // 
            // frmUpdate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(281, 196);
            this.Controls.Add(this.lblHistory);
            this.Controls.Add(this.lblUpdating);
            this.Controls.Add(this.lblCritical);
            this.Controls.Add(this.lblChecking);
            this.Controls.Add(this.picClose);
            this.Controls.Add(this.btnUpdateNow);
            this.Controls.Add(this.lblNoUpdate);
            this.Controls.Add(this.picProgressBar);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmUpdate";
            this.Load += new System.EventHandler(this.frmUpdate_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.frmUpdate_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.frmUpdate_MouseMove);
            ((System.ComponentModel.ISupportInitialize)(this.picProgressBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picClose)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox picProgressBar;
        private System.Windows.Forms.Button btnUpdateNow;
        private System.Windows.Forms.Label lblNoUpdate;
        private System.Windows.Forms.PictureBox picClose;
        private System.Windows.Forms.Label lblChecking;
        private System.Windows.Forms.LinkLabel lblCritical;
        private System.Windows.Forms.Label lblUpdating;
        private System.Windows.Forms.LinkLabel lblHistory;
    }
}