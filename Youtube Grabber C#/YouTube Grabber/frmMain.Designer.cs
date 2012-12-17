namespace YouTube_Grabber
{
    partial class frmMain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.mnuPopUp = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuStop = new System.Windows.Forms.ToolStripMenuItem();
            this.startDownloadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tpAbout = new System.Windows.Forms.TabPage();
            this.lblAbout = new System.Windows.Forms.Label();
            this.lblGoogle = new System.Windows.Forms.LinkLabel();
            this.lblBlog = new System.Windows.Forms.LinkLabel();
            this.lblVersion = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tpDownload = new System.Windows.Forms.TabPage();
            this.lvwDonwloadList = new System.Windows.Forms.ListView();
            this.chFilename = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chSize = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chProgress = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chTimeElapsed = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chFilepath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chQuality = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tpMain = new System.Windows.Forms.TabPage();
            this.pnlScroll = new System.Windows.Forms.Panel();
            this.lblScroll = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtLinkInfo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbDownOpt = new System.Windows.Forms.ComboBox();
            this.btnDownload = new System.Windows.Forms.Button();
            this.txtUrl = new System.Windows.Forms.TextBox();
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.pnlStatus = new System.Windows.Forms.Panel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.picVideoPic = new System.Windows.Forms.PictureBox();
            this.tcMain = new System.Windows.Forms.TabControl();
            this.mnuPopUp.SuspendLayout();
            this.tpAbout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.tpDownload.SuspendLayout();
            this.tpMain.SuspendLayout();
            this.pnlScroll.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.pnlStatus.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picVideoPic)).BeginInit();
            this.tcMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnuPopUp
            // 
            this.mnuPopUp.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuStop,
            this.startDownloadToolStripMenuItem});
            this.mnuPopUp.Name = "mnuPopUp";
            this.mnuPopUp.Size = new System.Drawing.Size(156, 48);
            // 
            // mnuStop
            // 
            this.mnuStop.Name = "mnuStop";
            this.mnuStop.Size = new System.Drawing.Size(155, 22);
            this.mnuStop.Text = "Stop Download";
            this.mnuStop.Click += new System.EventHandler(this.mnuStop_Click);
            // 
            // startDownloadToolStripMenuItem
            // 
            this.startDownloadToolStripMenuItem.Name = "startDownloadToolStripMenuItem";
            this.startDownloadToolStripMenuItem.Size = new System.Drawing.Size(155, 22);
            this.startDownloadToolStripMenuItem.Text = "Start Download";
            this.startDownloadToolStripMenuItem.Click += new System.EventHandler(this.startDownloadToolStripMenuItem_Click);
            // 
            // tpAbout
            // 
            this.tpAbout.BackColor = System.Drawing.SystemColors.Control;
            this.tpAbout.Controls.Add(this.lblAbout);
            this.tpAbout.Controls.Add(this.lblGoogle);
            this.tpAbout.Controls.Add(this.lblBlog);
            this.tpAbout.Controls.Add(this.lblVersion);
            this.tpAbout.Controls.Add(this.pictureBox1);
            this.tpAbout.Location = new System.Drawing.Point(4, 24);
            this.tpAbout.Name = "tpAbout";
            this.tpAbout.Padding = new System.Windows.Forms.Padding(3);
            this.tpAbout.Size = new System.Drawing.Size(888, 391);
            this.tpAbout.TabIndex = 2;
            this.tpAbout.Text = "About";
            // 
            // lblAbout
            // 
            this.lblAbout.AutoSize = true;
            this.lblAbout.Location = new System.Drawing.Point(470, 136);
            this.lblAbout.Name = "lblAbout";
            this.lblAbout.Size = new System.Drawing.Size(275, 30);
            this.lblAbout.TabIndex = 9;
            this.lblAbout.Text = "The .net version of AG YouTube Video Grabber\r\nRewritten with C# 2010 by GaryNg @ " +
                "AG dev team.";
            // 
            // lblGoogle
            // 
            this.lblGoogle.AutoSize = true;
            this.lblGoogle.Location = new System.Drawing.Point(470, 95);
            this.lblGoogle.Name = "lblGoogle";
            this.lblGoogle.Size = new System.Drawing.Size(234, 30);
            this.lblGoogle.TabIndex = 8;
            this.lblGoogle.TabStop = true;
            this.lblGoogle.Text = "YouTube Sniffer Google Code Page:\r\nhttp://code.google.com/p/youtubesniffer/";
            this.lblGoogle.Click += new System.EventHandler(this.lblGoogle_Click);
            // 
            // lblBlog
            // 
            this.lblBlog.AutoSize = true;
            this.lblBlog.Location = new System.Drawing.Point(470, 67);
            this.lblBlog.Name = "lblBlog";
            this.lblBlog.Size = new System.Drawing.Size(253, 15);
            this.lblBlog.TabIndex = 7;
            this.lblBlog.TabStop = true;
            this.lblBlog.Text = "My Blog: http://garyngzhongbo.blogspot.com";
            this.lblBlog.Click += new System.EventHandler(this.lblBlog_Click);
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(470, 35);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(89, 15);
            this.lblVersion.TabIndex = 6;
            this.lblVersion.Text = "Current Version";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::YouTube_Grabber.Properties.Resources.GrabberLogo;
            this.pictureBox1.Location = new System.Drawing.Point(144, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(320, 320);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // tpDownload
            // 
            this.tpDownload.BackColor = System.Drawing.SystemColors.Control;
            this.tpDownload.Controls.Add(this.lvwDonwloadList);
            this.tpDownload.Location = new System.Drawing.Point(4, 24);
            this.tpDownload.Name = "tpDownload";
            this.tpDownload.Padding = new System.Windows.Forms.Padding(3);
            this.tpDownload.Size = new System.Drawing.Size(888, 391);
            this.tpDownload.TabIndex = 1;
            this.tpDownload.Text = "Download";
            // 
            // lvwDonwloadList
            // 
            this.lvwDonwloadList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chFilename,
            this.chSize,
            this.chProgress,
            this.chTimeElapsed,
            this.chStatus,
            this.chFilepath,
            this.chQuality});
            this.lvwDonwloadList.ContextMenuStrip = this.mnuPopUp;
            this.lvwDonwloadList.FullRowSelect = true;
            this.lvwDonwloadList.GridLines = true;
            this.lvwDonwloadList.Location = new System.Drawing.Point(6, 7);
            this.lvwDonwloadList.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.lvwDonwloadList.MultiSelect = false;
            this.lvwDonwloadList.Name = "lvwDonwloadList";
            this.lvwDonwloadList.Size = new System.Drawing.Size(876, 377);
            this.lvwDonwloadList.TabIndex = 2;
            this.lvwDonwloadList.UseCompatibleStateImageBehavior = false;
            this.lvwDonwloadList.View = System.Windows.Forms.View.Details;
            // 
            // chFilename
            // 
            this.chFilename.Text = "Filename";
            this.chFilename.Width = 306;
            // 
            // chSize
            // 
            this.chSize.Text = "Size";
            this.chSize.Width = 73;
            // 
            // chProgress
            // 
            this.chProgress.Text = "Progress";
            this.chProgress.Width = 69;
            // 
            // chTimeElapsed
            // 
            this.chTimeElapsed.Text = "Time Elapsed";
            this.chTimeElapsed.Width = 146;
            // 
            // chStatus
            // 
            this.chStatus.Text = "Status";
            this.chStatus.Width = 96;
            // 
            // chFilepath
            // 
            this.chFilepath.Text = "File Path";
            this.chFilepath.Width = 101;
            // 
            // chQuality
            // 
            this.chQuality.Text = "Quality";
            this.chQuality.Width = 102;
            // 
            // tpMain
            // 
            this.tpMain.BackColor = System.Drawing.Color.Transparent;
            this.tpMain.Controls.Add(this.pnlScroll);
            this.tpMain.Controls.Add(this.groupBox1);
            this.tpMain.Controls.Add(this.txtUrl);
            this.tpMain.Controls.Add(this.txtInfo);
            this.tpMain.Controls.Add(this.pnlStatus);
            this.tpMain.Controls.Add(this.picVideoPic);
            this.tpMain.Location = new System.Drawing.Point(4, 24);
            this.tpMain.Name = "tpMain";
            this.tpMain.Padding = new System.Windows.Forms.Padding(3);
            this.tpMain.Size = new System.Drawing.Size(888, 391);
            this.tpMain.TabIndex = 0;
            this.tpMain.Text = "Main";
            // 
            // pnlScroll
            // 
            this.pnlScroll.Controls.Add(this.lblScroll);
            this.pnlScroll.Location = new System.Drawing.Point(6, 41);
            this.pnlScroll.Name = "pnlScroll";
            this.pnlScroll.Size = new System.Drawing.Size(612, 50);
            this.pnlScroll.TabIndex = 1;
            // 
            // lblScroll
            // 
            this.lblScroll.AutoSize = true;
            this.lblScroll.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScroll.Location = new System.Drawing.Point(3, 13);
            this.lblScroll.Name = "lblScroll";
            this.lblScroll.Size = new System.Drawing.Size(215, 25);
            this.lblScroll.TabIndex = 13;
            this.lblScroll.Text = "YouTube Video Grabber";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtLinkInfo);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cbDownOpt);
            this.groupBox1.Controls.Add(this.btnDownload);
            this.groupBox1.Location = new System.Drawing.Point(624, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(245, 357);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Download Options";
            // 
            // txtLinkInfo
            // 
            this.txtLinkInfo.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLinkInfo.Location = new System.Drawing.Point(6, 73);
            this.txtLinkInfo.Multiline = true;
            this.txtLinkInfo.Name = "txtLinkInfo";
            this.txtLinkInfo.ReadOnly = true;
            this.txtLinkInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLinkInfo.Size = new System.Drawing.Size(233, 226);
            this.txtLinkInfo.TabIndex = 3;
            this.txtLinkInfo.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(236, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Please select a video quality to download:";
            // 
            // cbDownOpt
            // 
            this.cbDownOpt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbDownOpt.Enabled = false;
            this.cbDownOpt.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbDownOpt.FormattingEnabled = true;
            this.cbDownOpt.Location = new System.Drawing.Point(6, 44);
            this.cbDownOpt.Name = "cbDownOpt";
            this.cbDownOpt.Size = new System.Drawing.Size(233, 21);
            this.cbDownOpt.TabIndex = 1;
            this.cbDownOpt.SelectedIndexChanged += new System.EventHandler(this.cbDownOpt_SelectedIndexChanged);
            // 
            // btnDownload
            // 
            this.btnDownload.Enabled = false;
            this.btnDownload.Location = new System.Drawing.Point(6, 308);
            this.btnDownload.Name = "btnDownload";
            this.btnDownload.Size = new System.Drawing.Size(233, 43);
            this.btnDownload.TabIndex = 0;
            this.btnDownload.Text = "Start Download!";
            this.btnDownload.UseVisualStyleBackColor = true;
            this.btnDownload.Click += new System.EventHandler(this.btnDownload_Click);
            // 
            // txtUrl
            // 
            this.txtUrl.Location = new System.Drawing.Point(6, 12);
            this.txtUrl.Name = "txtUrl";
            this.txtUrl.Size = new System.Drawing.Size(612, 23);
            this.txtUrl.TabIndex = 0;
            this.txtUrl.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUrl_KeyDown);
            this.txtUrl.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtUrl_KeyUp);
            this.txtUrl.MouseEnter += new System.EventHandler(this.txtUrl_MouseEnter);
            // 
            // txtInfo
            // 
            this.txtInfo.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtInfo.Location = new System.Drawing.Point(332, 97);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.ReadOnly = true;
            this.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtInfo.Size = new System.Drawing.Size(286, 240);
            this.txtInfo.TabIndex = 17;
            this.txtInfo.TabStop = false;
            // 
            // pnlStatus
            // 
            this.pnlStatus.Controls.Add(this.lblStatus);
            this.pnlStatus.Location = new System.Drawing.Point(6, 343);
            this.pnlStatus.Name = "pnlStatus";
            this.pnlStatus.Size = new System.Drawing.Size(612, 26);
            this.pnlStatus.TabIndex = 18;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(11, 6);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 15);
            this.lblStatus.TabIndex = 16;
            // 
            // picVideoPic
            // 
            this.picVideoPic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picVideoPic.InitialImage = null;
            this.picVideoPic.Location = new System.Drawing.Point(6, 97);
            this.picVideoPic.Name = "picVideoPic";
            this.picVideoPic.Size = new System.Drawing.Size(320, 240);
            this.picVideoPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picVideoPic.TabIndex = 10;
            this.picVideoPic.TabStop = false;
            // 
            // tcMain
            // 
            this.tcMain.Controls.Add(this.tpMain);
            this.tcMain.Controls.Add(this.tpDownload);
            this.tcMain.Controls.Add(this.tpAbout);
            this.tcMain.Location = new System.Drawing.Point(8, 12);
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.Size = new System.Drawing.Size(896, 419);
            this.tcMain.TabIndex = 21;
            this.tcMain.TabStop = false;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(917, 439);
            this.Controls.Add(this.tcMain);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.Text = "YouTube Video Grabber .net";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.mnuPopUp.ResumeLayout(false);
            this.tpAbout.ResumeLayout(false);
            this.tpAbout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.tpDownload.ResumeLayout(false);
            this.tpMain.ResumeLayout(false);
            this.tpMain.PerformLayout();
            this.pnlScroll.ResumeLayout(false);
            this.pnlScroll.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.pnlStatus.ResumeLayout(false);
            this.pnlStatus.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picVideoPic)).EndInit();
            this.tcMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip mnuPopUp;
        private System.Windows.Forms.ToolStripMenuItem mnuStop;
        private System.Windows.Forms.TabPage tpAbout;
        private System.Windows.Forms.Label lblAbout;
        private System.Windows.Forms.LinkLabel lblGoogle;
        private System.Windows.Forms.LinkLabel lblBlog;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TabPage tpDownload;
        private System.Windows.Forms.ListView lvwDonwloadList;
        private System.Windows.Forms.ColumnHeader chFilename;
        private System.Windows.Forms.ColumnHeader chSize;
        private System.Windows.Forms.ColumnHeader chProgress;
        private System.Windows.Forms.ColumnHeader chTimeElapsed;
        private System.Windows.Forms.ColumnHeader chStatus;
        private System.Windows.Forms.ColumnHeader chFilepath;
        private System.Windows.Forms.ColumnHeader chQuality;
        private System.Windows.Forms.TabPage tpMain;
        private System.Windows.Forms.Panel pnlScroll;
        private System.Windows.Forms.Label lblScroll;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtLinkInfo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbDownOpt;
        private System.Windows.Forms.Button btnDownload;
        private System.Windows.Forms.TextBox txtUrl;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.Panel pnlStatus;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.PictureBox picVideoPic;
        private System.Windows.Forms.TabControl tcMain;
        private System.Windows.Forms.ToolStripMenuItem startDownloadToolStripMenuItem;

    }
}

