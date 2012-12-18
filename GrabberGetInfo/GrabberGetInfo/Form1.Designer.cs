namespace GrabberGetInfo
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtVer = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtHistory = new System.Windows.Forms.TextBox();
            this.btnDone = new System.Windows.Forms.Button();
            this.btnCore = new System.Windows.Forms.Button();
            this.btnNewUI = new System.Windows.Forms.Button();
            this.btnNewFunc = new System.Windows.Forms.Button();
            this.btnUIUpdated = new System.Windows.Forms.Button();
            this.btnBugFixed = new System.Windows.Forms.Button();
            this.btnFuncRemove = new System.Windows.Forms.Button();
            this.btnUIRemoved = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "Current Version :";
            // 
            // txtVer
            // 
            this.txtVer.Location = new System.Drawing.Point(12, 76);
            this.txtVer.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtVer.Name = "txtVer";
            this.txtVer.ReadOnly = true;
            this.txtVer.Size = new System.Drawing.Size(638, 25);
            this.txtVer.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 105);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(264, 18);
            this.label2.TabIndex = 2;
            this.label2.Text = "Current Version Download Link : ";
            // 
            // txtDL
            // 
            this.txtDL.Location = new System.Drawing.Point(12, 127);
            this.txtDL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtDL.Name = "txtDL";
            this.txtDL.Size = new System.Drawing.Size(638, 25);
            this.txtDL.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 205);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(264, 18);
            this.label3.TabIndex = 4;
            this.label3.Text = "Current Version Update History :";
            // 
            // txtHistory
            // 
            this.txtHistory.AcceptsTab = true;
            this.txtHistory.Location = new System.Drawing.Point(12, 227);
            this.txtHistory.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtHistory.Multiline = true;
            this.txtHistory.Name = "txtHistory";
            this.txtHistory.Size = new System.Drawing.Size(507, 222);
            this.txtHistory.TabIndex = 0;
            // 
            // btnDone
            // 
            this.btnDone.Location = new System.Drawing.Point(12, 463);
            this.btnDone.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnDone.Name = "btnDone";
            this.btnDone.Size = new System.Drawing.Size(504, 60);
            this.btnDone.TabIndex = 6;
            this.btnDone.Text = "Done (Write To XML + Write To Updates Log)";
            this.btnDone.UseVisualStyleBackColor = true;
            this.btnDone.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // btnCore
            // 
            this.btnCore.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCore.Location = new System.Drawing.Point(525, 211);
            this.btnCore.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCore.Name = "btnCore";
            this.btnCore.Size = new System.Drawing.Size(123, 28);
            this.btnCore.TabIndex = 8;
            this.btnCore.Text = "[Core Updated]";
            this.btnCore.UseVisualStyleBackColor = true;
            this.btnCore.Click += new System.EventHandler(this.btnCore_Click);
            // 
            // btnNewUI
            // 
            this.btnNewUI.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewUI.Location = new System.Drawing.Point(525, 243);
            this.btnNewUI.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnNewUI.Name = "btnNewUI";
            this.btnNewUI.Size = new System.Drawing.Size(123, 28);
            this.btnNewUI.TabIndex = 9;
            this.btnNewUI.Text = "[New UI]";
            this.btnNewUI.UseVisualStyleBackColor = true;
            this.btnNewUI.Click += new System.EventHandler(this.btnNewUI_Click);
            // 
            // btnNewFunc
            // 
            this.btnNewFunc.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewFunc.Location = new System.Drawing.Point(525, 275);
            this.btnNewFunc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnNewFunc.Name = "btnNewFunc";
            this.btnNewFunc.Size = new System.Drawing.Size(123, 28);
            this.btnNewFunc.TabIndex = 10;
            this.btnNewFunc.Text = "[New Function]";
            this.btnNewFunc.UseVisualStyleBackColor = true;
            this.btnNewFunc.Click += new System.EventHandler(this.btnNewFunc_Click);
            // 
            // btnUIUpdated
            // 
            this.btnUIUpdated.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUIUpdated.Location = new System.Drawing.Point(525, 307);
            this.btnUIUpdated.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnUIUpdated.Name = "btnUIUpdated";
            this.btnUIUpdated.Size = new System.Drawing.Size(123, 28);
            this.btnUIUpdated.TabIndex = 11;
            this.btnUIUpdated.Text = "[UI Updated]";
            this.btnUIUpdated.UseVisualStyleBackColor = true;
            this.btnUIUpdated.Click += new System.EventHandler(this.btnUIUpdated_Click);
            // 
            // btnBugFixed
            // 
            this.btnBugFixed.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBugFixed.Location = new System.Drawing.Point(525, 339);
            this.btnBugFixed.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBugFixed.Name = "btnBugFixed";
            this.btnBugFixed.Size = new System.Drawing.Size(123, 28);
            this.btnBugFixed.TabIndex = 12;
            this.btnBugFixed.Text = "[Bug Fixed]";
            this.btnBugFixed.UseVisualStyleBackColor = true;
            this.btnBugFixed.Click += new System.EventHandler(this.btnBugFixed_Click);
            // 
            // btnFuncRemove
            // 
            this.btnFuncRemove.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFuncRemove.Location = new System.Drawing.Point(525, 371);
            this.btnFuncRemove.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnFuncRemove.Name = "btnFuncRemove";
            this.btnFuncRemove.Size = new System.Drawing.Size(123, 37);
            this.btnFuncRemove.TabIndex = 13;
            this.btnFuncRemove.Text = "[Function Removed]";
            this.btnFuncRemove.UseVisualStyleBackColor = true;
            this.btnFuncRemove.Click += new System.EventHandler(this.btnFuncRemove_Click);
            // 
            // btnUIRemoved
            // 
            this.btnUIRemoved.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUIRemoved.Location = new System.Drawing.Point(525, 412);
            this.btnUIRemoved.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnUIRemoved.Name = "btnUIRemoved";
            this.btnUIRemoved.Size = new System.Drawing.Size(123, 37);
            this.btnUIRemoved.TabIndex = 14;
            this.btnUIRemoved.Text = "[UI Removed]";
            this.btnUIRemoved.UseVisualStyleBackColor = true;
            this.btnUIRemoved.Click += new System.EventHandler(this.btnUIRemoved_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(522, 191);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 18);
            this.label4.TabIndex = 15;
            this.label4.Text = "--Shortcuts--";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Maroon;
            this.label5.Location = new System.Drawing.Point(12, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(504, 36);
            this.label5.TabIndex = 16;
            this.label5.Text = "Please fill in those information (for the settings of updater)\r\n--GaryNg";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 156);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(104, 18);
            this.label6.TabIndex = 17;
            this.label6.Text = "Build Date :";
            // 
            // txtDate
            // 
            this.txtDate.Location = new System.Drawing.Point(11, 177);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(505, 25);
            this.txtDate.TabIndex = 18;
            this.txtDate.TextChanged += new System.EventHandler(this.txtDate_TextChanged);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExit.Location = new System.Drawing.Point(522, 463);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(126, 60);
            this.btnExit.TabIndex = 19;
            this.btnExit.Text = "Developer Testing (Don\'t Write Anything)";
            this.btnExit.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(666, 531);
            this.ControlBox = false;
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnUIRemoved);
            this.Controls.Add(this.btnFuncRemove);
            this.Controls.Add(this.btnBugFixed);
            this.Controls.Add(this.btnUIUpdated);
            this.Controls.Add(this.btnNewFunc);
            this.Controls.Add(this.btnNewUI);
            this.Controls.Add(this.btnCore);
            this.Controls.Add(this.btnDone);
            this.Controls.Add(this.txtHistory);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtDL);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtVer);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Consolas", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.Text = "Grabber Set Version Info";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtVer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDL;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtHistory;
        private System.Windows.Forms.Button btnDone;
        private System.Windows.Forms.Button btnCore;
        private System.Windows.Forms.Button btnNewUI;
        private System.Windows.Forms.Button btnNewFunc;
        private System.Windows.Forms.Button btnUIUpdated;
        private System.Windows.Forms.Button btnBugFixed;
        private System.Windows.Forms.Button btnFuncRemove;
        private System.Windows.Forms.Button btnUIRemoved;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.Button btnExit;
    }
}

