using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;

namespace YouTube_Grabber
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        #region MainFormFunc
        delegate void SetTextHandler(Control t, string text);
        delegate void SetImageHandler(PictureBox pic, string url);
        delegate void ZeroParamHandler();
        delegate void EnDisableControlHandler(Control cnt, bool bControlState);
        delegate void SetControlPos(Control cntControl, int Left);
        delegate void ComboBoxAddItemhandler(ComboBox cbBox, string str);

        public struct LinkInfo
        {
            public string strVideoTitle;
            public string strVideoLink;
            public string strVideoType;
            public string strVideoSize;
            public string strQuality;
        }
        string strCurrentVideoTitle;

        Thread tScroll;
        List<LinkInfo> strDownloadLinks;

        clsMain clsDownload;
        int intCurrentSelectedDownloadLinks;
        Form frmUpdater;
        void ComboBoxAddItem(ComboBox cbBox, string str)
        {
            cbBox.Items.Add(str);
            //cbBox.SelectedIndex = 0;
        }

        void EnDisableControl(Control cnt, bool bControlState)
        {
            cnt.Enabled = bControlState;
        }

        void SetText(Control t, string text)
        {
            Monitor.Enter(this);
            t.Text = text;
            Monitor.Exit(this);
        }

        void SetImage(PictureBox pic, string url)
        {
            pic.ImageLocation = url;
        }


        void GotLostFocus(TextBox txtFocus, bool GotFocus)
        {
            if (GotFocus)
            {
                if (txtFocus.Text == "Please input the YouTube video link, press enter when your're done...")
                {
                    txtFocus.Text = "";
                    txtFocus.Font = new Font(txtFocus.Font, FontStyle.Regular);
                    txtFocus.ForeColor = Color.Black;
                }
            }
            else if (GotFocus == false)
            {
                if (txtFocus.Text == "")
                {
                    txtFocus.Text = "Please input the YouTube video link, press enter when your're done...";
                    txtFocus.Font = new Font(txtFocus.Font, FontStyle.Italic);
                    txtFocus.ForeColor = Color.Gray;
                }
            }
        }
        void ThreadScrollFunc()
        {
            SetControlPos scpControl = new SetControlPos(SetConPos);
            for (; ; )
            {
                object[] objLabel = { lblScroll, lblScroll.Left - 1 };
                lblScroll.Invoke(scpControl, objLabel);
                if (lblScroll.Left == pnlScroll.Left - lblScroll.Width)
                {
                    object[] objSetPos = { lblScroll, pnlScroll.Left + pnlScroll.Width };
                    lblScroll.Invoke(scpControl, objSetPos);
                }
                //
                Application.DoEvents();
                Thread.Sleep(10);
            }

        }

        void SetConPos(Control cntControl, int Left)
        {
            try
            {
                cntControl.Left = Left;
                string strClip = Clipboard.GetText();
                string strTxtUrl = txtUrl.Text;
                if (strClip == "") return;
                if (strClip.IndexOf("://") != -1)
                {
                    strClip = strClip.Remove(0, strClip.IndexOf("://") + 3);
                }

                if (strTxtUrl.IndexOf("://") != -1)
                {
                    strTxtUrl = strTxtUrl.Remove(0, strTxtUrl.IndexOf("://") + 3);
                }
                if (strClip.Substring(0, "www.youtube.com/watch?".Length) == "www.youtube.com/watch?")
                {
                    if (strTxtUrl != strClip)
                    {
                        this.Activate();
                        txtUrl.Text = Clipboard.GetText();
                        KeyEventArgs e = new KeyEventArgs(Keys.Enter);
                        txtUrl_GotFocus(this, e);
                        txtUrl_KeyUp(this, e);
                    }
                }
            }
            catch { }
            //if (Clipboard.GetText().Contains("www.youtube.com/watch?"))
            //{
            //    if (txtUrl.Text != Clipboard.GetText())
            //    {
            //        txtUrl.Text = Clipboard.GetText();
            //        KeyEventArgs e  =  new KeyEventArgs(Keys.Enter);
            //        txtUrl_GotFocus(this, e);
            //        txtUrl_KeyUp(this, e);
            //    }

            //}
        }

        void CleanUp()
        {
            txtUrl.Text = "";
            lblScroll.Text = "YouTube Video Grabber";
            lblScroll.Left = pnlScroll.Left;
            picVideoPic.ImageLocation = "";
            lblStatus.Text = "";
            txtInfo.Text = "";
            cbDownOpt.Items.Clear();
            cbDownOpt.Enabled = false;
            btnDownload.Enabled = false;
            intCurrentSelectedDownloadLinks = 0;
            strDownloadLinks.Clear();
            txtLinkInfo.Text = "";

        }

        #endregion
        #region DownloaderFunc
        List<LinkInfo> liDonwloadList;
        int intDownloadListCount;
        List<clsDownloader> DownloadQueue;
        struct DownloadInfo
        {
            public string strStatus;
            public string strFilePath;
            public int intDownloadPercentage;
            public string strTimeEplased;

        }
        List<DownloadInfo> diDownloadQueue;

        string CheckForDupFileName(string filename, string extension)
        {
            if (File.Exists(filename + extension))
            {
                int index = 1;
                string tmpfilename = filename;
                for (; ; )
                {
                    if (File.Exists(tmpfilename + "(" + index.ToString() + ")" + extension))
                    {
                        index++;
                        continue;
                    }
                    return tmpfilename + "(" + index.ToString() + ")";
                }
            }
            else
            {
                return filename;
            }
        }

        void UpdateListView(int index)
        {
            lvwDonwloadList.Items[index].SubItems[2].Text = diDownloadQueue[index].intDownloadPercentage.ToString() + "%";
            lvwDonwloadList.Items[index].SubItems[3].Text = diDownloadQueue[index].strTimeEplased;
            lvwDonwloadList.Items[index].SubItems[4].Text = diDownloadQueue[index].strStatus;
        }

        void frmMain_DownloadProgressChangedIndex(System.Net.DownloadProgressChangedEventArgs e, int Index, string time, string speed)
        {

            DownloadInfo diTmp = diDownloadQueue[Index];
            diTmp.intDownloadPercentage = e.ProgressPercentage;
            diTmp.strTimeEplased = time + " ( " + speed + "/s" + " ) ";
            diTmp.strStatus = "Downloading...";
            diDownloadQueue[Index] = diTmp;
            UpdateListView(Index);
            System.Diagnostics.Debug.Print(speed);


        }
        void frmMain_DownloadFileCompletedIndex(AsyncCompletedEventArgs e, int Index)
        {
            if (!e.Cancelled)
            {
                DownloadInfo diTmp = diDownloadQueue[Index];
                diTmp.strStatus = "Done";
                diDownloadQueue[Index] = diTmp;
                UpdateListView(Index);
            }

        }

        private void mnuStop_Click(object sender, EventArgs e)
        {
            if (lvwDonwloadList.SelectedItems.Count == 0)
            {
                return;
            }
            if (diDownloadQueue[lvwDonwloadList.SelectedIndices[0]].strStatus == "Done")
            {
                return;
            }

            DownloadQueue[lvwDonwloadList.SelectedIndices[0]].CancelAsync();
            DownloadInfo diTmp = diDownloadQueue[lvwDonwloadList.SelectedIndices[0]];
            //File.Delete(diTmp.strFilePath);
            diTmp.strStatus = "Canceled...";
            diTmp.intDownloadPercentage = 0;
            diDownloadQueue[lvwDonwloadList.SelectedIndices[0]] = diTmp;
            UpdateListView(lvwDonwloadList.SelectedIndices[0]);



        }
        private void startDownloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lvwDonwloadList.SelectedItems.Count == 0)
            {
                return;
            }
            if (diDownloadQueue[lvwDonwloadList.SelectedIndices[0]].strStatus == "Canceled...")
            {
                DownloadQueue[lvwDonwloadList.SelectedIndices[0]].DownloadFileAsync(new Uri(liDonwloadList[lvwDonwloadList.SelectedIndices[0]].strVideoLink), diDownloadQueue[lvwDonwloadList.SelectedIndices[0]].strFilePath); ;
                DownloadInfo diTmp = diDownloadQueue[lvwDonwloadList.SelectedIndices[0]];
                diTmp.strStatus = "Downloading...";
                diTmp.intDownloadPercentage = 0;
                diDownloadQueue[lvwDonwloadList.SelectedIndices[0]] = diTmp;
                UpdateListView(lvwDonwloadList.SelectedIndices[0]);
            }

        }

        #endregion
        #region MainFormEvent
        
        private void btnDownload_Click(object sender, EventArgs e)
        {
            
            liDonwloadList.Add(strDownloadLinks[intCurrentSelectedDownloadLinks]);
            string filename = liDonwloadList[intDownloadListCount].strVideoTitle;
            string extension = liDonwloadList[intDownloadListCount].strVideoType.Replace("*", "");

            SaveFileDialog sfdSave = new SaveFileDialog();
            sfdSave.DefaultExt = liDonwloadList[intDownloadListCount].strVideoType;
            sfdSave.Filter = liDonwloadList[intDownloadListCount].strVideoType + "|" + liDonwloadList[intDownloadListCount].strVideoType;

            if (sfdSave.ShowDialog() == DialogResult.OK)
            {

                DownloadInfo diTmp = new DownloadInfo();
                diTmp.intDownloadPercentage = 0;
                diTmp.strFilePath = sfdSave.FileName;
                diTmp.strTimeEplased = "00:00:00";
                diTmp.strStatus = "Pending...";
                diDownloadQueue.Add(new DownloadInfo());
                diDownloadQueue[intDownloadListCount] = diTmp;


                DownloadQueue.Add(new clsDownloader(intDownloadListCount));
                DownloadQueue[intDownloadListCount].DownloadFileCompletedIndex += new clsDownloader.OnDownloadFileCompletedIndexEventHandler(frmMain_DownloadFileCompletedIndex);
                DownloadQueue[intDownloadListCount].DownloadProgressChangedIndex += new clsDownloader.OnDownloadProgressChangedIndexEventHandler(frmMain_DownloadProgressChangedIndex);
                DownloadQueue[intDownloadListCount].DownloadFileAsync(new Uri(liDonwloadList[intDownloadListCount].strVideoLink), diTmp.strFilePath);

                ListViewItem lviDownload = new ListViewItem(filename + extension);
                lviDownload.SubItems.Add(liDonwloadList[intDownloadListCount].strVideoSize);
                lviDownload.SubItems.Add("0%");
                lviDownload.SubItems.Add(diTmp.strTimeEplased);
                lviDownload.SubItems.Add(diTmp.strStatus);
                lviDownload.SubItems.Add(diTmp.strFilePath);
                lviDownload.SubItems.Add(liDonwloadList[intDownloadListCount].strQuality);
                intDownloadListCount++;
                lvwDonwloadList.Items.Add(lviDownload);
            }
            else
            {
                return;
            }

            tcMain.SelectedIndex = 1;


            //string newFilename = CheckForDupFileName(filename, extension);

            //LinkInfo liTmp = liDonwloadList[intDownloadListCount];
            //liTmp.strVideoTitle = newFilename;
            //liDonwloadList[intDownloadListCount] = liTmp;
        }

        private void txtUrl_MouseEnter(object sender, EventArgs e)
        {
            ToolTip tt = new ToolTip();
            tt.SetToolTip(txtUrl, "Please input the Youtube video URL, press enter when you're done...");
        }
        private void cbDownOpt_SelectedIndexChanged(object sender, EventArgs e)
        {

            btnDownload.Enabled = true;
            intCurrentSelectedDownloadLinks = cbDownOpt.SelectedIndex;
            LinkInfo liTmp = strDownloadLinks[cbDownOpt.SelectedIndex];

            txtLinkInfo.Text = "[QUALITY:]" + Environment.NewLine + liTmp.strQuality + Environment.NewLine + Environment.NewLine;
            txtLinkInfo.Text += "[TYPE:]" + Environment.NewLine + liTmp.strVideoType + Environment.NewLine + Environment.NewLine;
            txtLinkInfo.Text += "[DOWNLOAD SIZE:]" + Environment.NewLine + liTmp.strVideoSize + Environment.NewLine + Environment.NewLine;
            txtLinkInfo.Text += "[DOWNLOAD LINK:]" + Environment.NewLine + liTmp.strVideoLink + Environment.NewLine + Environment.NewLine;

        }
        private void txtUrl_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtUrl.Text.Contains(Environment.NewLine))
            {
                txtUrl.Text = txtUrl.Text.Replace(Environment.NewLine, "");
                txtUrl.SelectionStart = txtUrl.Text.Length;
            }

            if (e.KeyCode == Keys.Enter && (txtUrl.Text == "" || !txtUrl.Text.Contains("www.youtube.com/watch?")))
            {
                CleanUp();
                lblStatus.Text = "Please enter a vaild url";
                txtUrl.Text = "";
                txtUrl_LostFocus(this, new EventArgs());
                this.Focus();
                return;
            }


            if (e.KeyCode == Keys.Enter)
            {
                txtUrl.Enabled = false;
                string strTextTmp = txtUrl.Text;
                CleanUp();
                picVideoPic.Focus();
                picVideoPic.Image = Properties.Resources.Loading;
                txtUrl.Text = strTextTmp;
                clsDownload = new clsMain(txtUrl.Text);
                clsDownload.Progress += new clsMain.ProgressEventHandler(clsDownload_Progress);
                clsDownload.DownloadLinksDone += new clsMain.ProgressEventHandler(clsDownload_DownloadLinksDone);
                clsDownload.VideoInfoReceived += new clsMain.VideoInfoEventHandler(clsDownload_VideoInfoReceived);
                clsDownload.Error += new clsMain.ErrorHandler(clsDownload_Error);
                clsDownload.Start();
            }

        }
        void clsDownload_Progress(clsMain.ProgressEventArgs Status)
        {
            SetTextHandler sth = new SetTextHandler(SetText);
            if (Status.Status == "Donwload links done")
            {
                EnDisableControlHandler edch = new EnDisableControlHandler(EnDisableControl);
                object[] objState = { txtUrl, true };
                txtUrl.Invoke(edch, objState);
                object[] objLinks = { lblStatus, "Done (" + Status.strDownloadLinks.Count.ToString() + " download links found)" };
                lblStatus.Invoke(sth, objLinks);
            }
            else
            {
                object[] objStatus = { lblStatus, Status.Status };
                lblStatus.Invoke(sth, objStatus);
            }
        }
        void clsDownload_Error(string err)
        {
            MessageBox.Show(err);
            SetTextHandler sth = new SetTextHandler(SetText);
            object[] objText = { txtUrl, "" };
            txtUrl.Invoke(sth, objText);
            ZeroParamHandler cuh = new ZeroParamHandler(CleanUp);
            this.Invoke(cuh);
            object[] objErr = { lblStatus, err };
            lblStatus.Invoke(sth, objErr);
            Environment.Exit(-1);

        }
        void clsDownload_VideoInfoReceived(clsMain.VideoInfoEventArgs VideoInfo)
        {
            SetTextHandler sth = new SetTextHandler(SetText);
            strCurrentVideoTitle = VideoInfo.strTitle;
            object[] text = { txtInfo, "[AUTHOR YOUTUBE CHANNEL:]" + Environment.NewLine + VideoInfo.strAuthor + Environment.NewLine + Environment.NewLine + "[VIDEO VIEWS:]" + Environment.NewLine + VideoInfo.strViews + Environment.NewLine + Environment.NewLine + "[LIKES:]" + Environment.NewLine + VideoInfo.strLikes + Environment.NewLine + Environment.NewLine + "[DISLIKES:]" + Environment.NewLine + VideoInfo.strDislikes + Environment.NewLine + Environment.NewLine + "[VIDEO DURATION:]" + Environment.NewLine + VideoInfo.strLength };
            txtInfo.Invoke(sth, text);
            SetImageHandler si = new SetImageHandler(SetImage);
            //object[] objLoading = { picVideoPic,Properties.Resources.Loading };
            //picVideoPic.Invoke(si, objLoading);

            object[] objImage = { picVideoPic, VideoInfo.strPic };
            picVideoPic.Invoke(si, objImage);
            SetControlPos scp = new SetControlPos(SetConPos);
            object[] objScroll = { lblScroll, pnlScroll.Left };
            lblScroll.Invoke(scp, objScroll);
            object[] strTitle = { lblScroll, VideoInfo.strTitle };
            lblScroll.Invoke(sth, strTitle);



        }
        void clsDownload_DownloadLinksDone(clsMain.ProgressEventArgs Status)
        {
            string DownloadLinks = "";
            SetTextHandler sth = new SetTextHandler(SetText);
            ComboBoxAddItemhandler cbaih = new ComboBoxAddItemhandler(ComboBoxAddItem);
            for (int i = 0; i < Status.strDownloadLinks.Count; i++)
            {
                //object[] objCombo = {cbDownOpt, Status.strDownloadLinks[i]};
                object[] objCombo = { cbDownOpt, Status.strQuality[i] + " ( " + Status.strType[i] + " ) [" + Status.strSize[i] + " ]" };
                cbDownOpt.Invoke(cbaih, objCombo);
                DownloadLinks += Status.strDownloadLinks[i] + "\t\t" + Status.strQuality[i] + "\t\t" + Status.strType[i] + Environment.NewLine + Environment.NewLine; //+ "\t\t\t"  +Status.strQuality[i] + Status.strType[i] + Environment.NewLine + Environment.NewLine;

                strDownloadLinks.Add(new LinkInfo());
                LinkInfo liTmp = strDownloadLinks[i];
                liTmp.strVideoTitle = strCurrentVideoTitle;
                liTmp.strVideoSize = Status.strSize[i];
                liTmp.strVideoType = Status.strType[i];
                liTmp.strQuality = Status.strQuality[i];
                liTmp.strVideoLink = Status.strDownloadLinks[i];
                strDownloadLinks[i] = liTmp;

            }
            EnDisableControlHandler edch = new EnDisableControlHandler(EnDisableControl);
            object[] objEnable = { cbDownOpt, true };
            cbDownOpt.Invoke(edch, objEnable);

        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            txtInfo.BackColor = this.BackColor;
            txtLinkInfo.BackColor = this.BackColor;
            tpMain.BackColor = this.BackColor;
            tpDownload.BackColor = this.BackColor;
            tpAbout.BackColor = this.BackColor;
            //txtDownloadInfo.BackColor = this.BackColor;

            strDownloadLinks = new List<LinkInfo>();
            liDonwloadList = new List<LinkInfo>();
            DownloadQueue = new List<clsDownloader>();
            diDownloadQueue = new List<DownloadInfo>();
            intDownloadListCount = 0;

            txtUrl.GotFocus += new EventHandler(txtUrl_GotFocus);
            txtUrl.LostFocus += new EventHandler(txtUrl_LostFocus);

            lblVersion.Text = "Current Build Version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Revision.ToString();

            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            tpMain.BackColor = Color.Transparent;
            tpDownload.BackColor = Color.Transparent;
            tpAbout.BackColor = Color.Transparent;

            tScroll = new Thread(ThreadScrollFunc);
            tScroll.Start();

            frmUpdater = new frmUpdate();
        }

       
        void txtUrl_LostFocus(object sender, EventArgs e)
        {
            GotLostFocus(txtUrl, false);
        }
        void txtUrl_GotFocus(object sender, EventArgs e)
        {
            GotLostFocus(txtUrl, true);
        }
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            tScroll.Abort();
            this.Dispose();
            //clsDownload.Dispose();
            Environment.Exit(-1);
            //clsDownload.Dispose();
        }
        private void txtUrl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
            }
        }
        #endregion        
        #region AboutFromEvent
        private void lblBlog_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://garyngzhongbo.blogspot.com");
        }
        private void lblGoogle_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://code.google.com/p/youtubesniffer/");
        }
        private void linkLabel1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://youtubesniffer.tk/");
        }
        private void btnCheckUpdate_Click(object sender, EventArgs e)
        {
            if (frmUpdater.Visible == false)
            {
                frmUpdater = new frmUpdate();
                frmUpdater.Show();
            }
        }
        #endregion
    }

}
