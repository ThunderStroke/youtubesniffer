using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.IO;
namespace YouTube_Grabber
{
    public partial class frmUpdate : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        delegate void ZeroParamHandler();
        Thread tDelayExit;

        clsUpdater clsCheckForUpdate;
        
        public frmUpdate()
        {
            InitializeComponent();
        }




        private void frmUpdate_Load(object sender, EventArgs e)
        {
            DrawClose(picClose, false);
            MoveToCenter(lblChecking);
            MoveFormToBottom();
            clsCheckForUpdate = new clsUpdater(picProgressBar);
            clsCheckForUpdate.InfoLoaded += new clsUpdater.ZeroParamEventHandler(clsCheckForUpdate_InfoLoaded);          

        }

        void MoveFormToBottom()
        {
            this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Left + Screen.PrimaryScreen.WorkingArea.Width - this.Size.Width, Screen.PrimaryScreen.WorkingArea.Bottom - this.Size.Height);
        }


        void CheckForUpdate()
        {
            lblChecking.Visible = false;
            bool bGotUpdate = clsCheckForUpdate.CheckForNewVersion;
            if (bGotUpdate == false)
            {
                MoveToCenter(lblNoUpdate);
                lblNoUpdate.Visible = true;
                tDelayExit = new Thread(DelayExit);
                tDelayExit.Start();
            }
            else if (bGotUpdate == true)
            {
                MoveToCenter(btnUpdateNow);
                btnUpdateNow.Text = "New Version Found! ( Version " + clsCheckForUpdate.NewVersion + " )" + Environment.NewLine + "Update Now!";
                lblHistory.Top = btnUpdateNow.Bottom + 10;
                lblHistory.Left = btnUpdateNow.Left;
                lblHistory.Visible = true;
                btnUpdateNow.Visible = true;
                if (clsCheckForUpdate.IsCritical == true)
                {
                    btnUpdateNow.Visible = false;
                    MoveToCenter(lblCritical);
                    lblCritical.Visible = true;
                }
            }
        }


        void DrawClose(PictureBox pBox, bool bClick)
        {
            Bitmap bClose = new Bitmap(pBox.Width, pBox.Height);
            Graphics gClose = Graphics.FromImage(bClose);
            Pen pClose;
            if (bClick)
            {
                pClose = new Pen(new SolidBrush(Color.Black), 2);
            }
            else
            {
                pClose = new Pen(new SolidBrush(Color.DarkGray), 2);    
            }
                
            gClose.DrawLine(pClose, Point.Empty, new Point(8, 8));
            gClose.DrawLine(pClose, new Point(0, 8), new Point(8, 0));
            if (bClick)
            {
                Rectangle rectClose = new Rectangle(new Point(1,1),new Size(pBox.Width - 1 ,pBox.Height - 1));
                gClose.DrawRectangle(Pens.White, rectClose);
            }
            pBox.Image = bClose;
        }
        void MoveToCenter(Control ctlToMove)
        {
            ctlToMove.Left = this.Width / 2 - ctlToMove.Width / 2;
            ctlToMove.Top = this.Height / 2 - ctlToMove.Height / 2;
        }

        void clsCheckForUpdate_InfoLoaded()
        {
            ZeroParamHandler zph = new ZeroParamHandler(CheckForUpdate);
            picProgressBar.Invoke(zph);
            
        }

        void DelayExit()
        {
            Thread.Sleep(3000);
            ZeroParamHandler zph = new ZeroParamHandler(CloseForm);
            this.Invoke(zph);
        }

        void CloseForm()
        {
            clsCheckForUpdate.Dispose();
            this.Close();
        }

        private void frmUpdate_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void picClose_Click(object sender, EventArgs e)
        {
            DrawClose(picClose, true);
            CloseForm();
        }

        private void picClose_MouseEnter(object sender, EventArgs e)
        {
            DrawClose(picClose, true);
        }

        private void picClose_MouseLeave(object sender, EventArgs e)
        {
            DrawClose(picClose, false);
        }

        private void lblCritical_Click(object sender, EventArgs e)
        {
            Process.Start(clsCheckForUpdate.ArticleLink);
            DelayExit();
        }

        private void btnUpdateNow_Click(object sender, EventArgs e)
        {
            btnUpdateNow.Visible = false;
            MoveToCenter(picProgressBar);
            lblUpdating.Left = picProgressBar.Left;
            lblUpdating.Top = picProgressBar.Top - 20;
            clsCheckForUpdate.DrawProgressBar(picProgressBar, 0);
            lblUpdating.Visible = true;
            picProgressBar.Visible = true;
            clsCheckForUpdate.StartDownloadUpdate();
            clsCheckForUpdate.DownloadUpdateCompleted += new AsyncCompletedEventHandler(clsCheckForUpdate_DownloadUpdateCompleted);


        }

        void clsCheckForUpdate_DownloadUpdateCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Cancelled == false)
            {
                MoveToCenter(lblUpdating);
                lblUpdating.Text = "Applying Update...";
                lblHistory.Visible = false;
                picProgressBar.Visible = false;
                clsCheckForUpdate.ApplyUpdate();
            }
            else
            { 
                File.Delete(clsCheckForUpdate.NewVersion + ".tmp"); 
            }
        }

        private void frmUpdate_MouseMove(object sender, MouseEventArgs e)
        {
            int intRight = this.Right - Screen.PrimaryScreen.WorkingArea.Right;
            int intBottom = this.Bottom - Screen.PrimaryScreen.WorkingArea.Bottom;
            if (intRight > 0 || intBottom > 0)
            {
                MoveFormToBottom();
            }
        }

        private void lblHistory_Click(object sender, EventArgs e)
        {
            MessageBox.Show(clsCheckForUpdate.UpdateHistory);
        }
    }
}
