using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Net;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Diagnostics;
using System.IO;
using System.Threading;
namespace YouTube_Grabber
{
    class clsUpdater
    {
        #region Variable
        int[] intNewVersion = new int[4];
        int[] intCurrentVersion = new int[4];
        string strDownloadLink = "";
        bool bIsCritical = false;
        bool bGotUpdate = false;
        string strArticleLink = "";
        string strHistory;
        WebClient wbUpdate = new WebClient();
        PictureBox picBar;
        public delegate void ZeroParamEventHandler();

        public event System.ComponentModel.AsyncCompletedEventHandler DownloadUpdateCompleted;

        public event ZeroParamEventHandler InfoLoaded;
        Thread tLoadInfo;
        #endregion
        public clsUpdater(PictureBox picProgressBar)
        {
            tLoadInfo = new Thread(LoadInfo);
            tLoadInfo.Start();
            picBar = picProgressBar;
        }

        public void Dispose()
        {
            wbUpdate.CancelAsync();
            wbUpdate.Dispose();
            picBar.Dispose();
            tLoadInfo.Abort();
        }

        void LoadInfo()
        {
            ReadOnlineInfo();
            ReadCurrentInfo();
            InfoLoaded();
        }

        #region PrivateFunction

        void ReadOnlineInfo()
        {
            XmlDocument xOnline = new XmlDocument();
            xOnline.Load(@"https://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/C%23/YouTube%20Video%20Grabber/Settings.xml");
            intNewVersion[0] = Convert.ToInt32(xOnline.SelectSingleNode("Settings/Version/Major").InnerText);
            intNewVersion[1] = Convert.ToInt32(xOnline.SelectSingleNode("Settings/Version/Minor").InnerText);
            intNewVersion[2] = Convert.ToInt32(xOnline.SelectSingleNode("Settings/Version/Build").InnerText);
            intNewVersion[3] = Convert.ToInt32(xOnline.SelectSingleNode("Settings/Version/Revision").InnerText);
            strHistory = xOnline.SelectSingleNode("Settings/History").InnerText;
            bIsCritical = Convert.ToBoolean(xOnline.SelectSingleNode("Settings/CriticalUpdate/IsCriticleUpdate").InnerText);
            if (bIsCritical == true)
            {
                strArticleLink = xOnline.SelectSingleNode("Settings/CriticalUpdate/ArticleLink").InnerText;
            }
            else
            {
                strDownloadLink = xOnline.SelectSingleNode("Settings/DownloadLink").InnerText;
            }
        }
        void ReadCurrentInfo()
        {
            XmlDocument xCurrent = new XmlDocument();
            xCurrent.Load(@"Settings.xml");
            intCurrentVersion[0] = Convert.ToInt32(xCurrent.SelectSingleNode("Settings/Version/Major").InnerText);
            intCurrentVersion[1] = Convert.ToInt32(xCurrent.SelectSingleNode("Settings/Version/Minor").InnerText);
            intCurrentVersion[2] = Convert.ToInt32(xCurrent.SelectSingleNode("Settings/Version/Build").InnerText);
            intCurrentVersion[3] = Convert.ToInt32(xCurrent.SelectSingleNode("Settings/Version/Revision").InnerText);
        }
        Boolean CompareVersion()
        {
            if (intCurrentVersion[0] < intNewVersion[0])
            {
                return true;
            }
            else if (intCurrentVersion[0] == intNewVersion[0])
            {
                if (intCurrentVersion[1] < intCurrentVersion[1])
                {
                    return true;
                }
                else if (intCurrentVersion[1] == intNewVersion[1])
                {
                    if (intCurrentVersion[2] < intNewVersion[2])
                    {
                        return true;
                    }
                    else if (intCurrentVersion[2] == intNewVersion[2])
                    {
                        if (intCurrentVersion[3] < intNewVersion[3])
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        void wbSettings_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            Process.Start(Application.ExecutablePath);
            Environment.Exit(-1);
        }
        void wbUpdate_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            DownloadUpdateCompleted(sender,e);
        }
        void wbUpdate_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            DrawProgressBar(picBar, e.ProgressPercentage);
        }
        #endregion
        #region PublicFuncion
        /// <summary>
        /// Next, Call IsGotOtherWindows
        /// </summary>
        public void StartDownloadUpdate()
        {
            if (bIsCritical) return;
            wbUpdate.DownloadProgressChanged += new DownloadProgressChangedEventHandler(wbUpdate_DownloadProgressChanged);
            wbUpdate.DownloadFileCompleted += new System.ComponentModel.AsyncCompletedEventHandler(wbUpdate_DownloadFileCompleted);
            wbUpdate.DownloadFileAsync(new Uri(strDownloadLink), NewVersion + ".tmp");
        }
        public void ApplyUpdate()
        {
            if (File.Exists(@"Updates\") == false)
            {
                Directory.CreateDirectory("Updates");
            }
            if (File.Exists(@"Updates\YouTube Grabber - " + CurrentVersion + ".exe") == false)
            {
                File.Move("YouTube Grabber.exe", @"Updates\YouTube Grabber - " + CurrentVersion + ".exe");
                File.Move(NewVersion + ".tmp", "YouTube Grabber.exe");
            }
            WebClient wbSettings = new WebClient();
            wbSettings.DownloadFileAsync(new Uri(@"https://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/C%23/YouTube%20Video%20Grabber/Settings.xml"), "Settings.xml");
            wbSettings.DownloadFileCompleted += new System.ComponentModel.AsyncCompletedEventHandler(wbSettings_DownloadFileCompleted);
        }
        public void DrawProgressBar(PictureBox picBar, int Value)
        {
            Bitmap bBar = new Bitmap(picBar.Width, picBar.Height);
            Graphics gBar = Graphics.FromImage(bBar);
            int intWidth = (int)((float)(picBar.Width / 100.00f) * Value);
            int intHeight = picBar.Height;
            if (Value != 0)
            {



                Color clrDark = Color.FromArgb(0, 132, 255);
                Color clrLight = Color.FromArgb(0, 255, 255);

                Rectangle rectDown = new Rectangle(new Point(0, intHeight / 2), new Size(intWidth, intHeight / 2));
                LinearGradientBrush lgbDown = new LinearGradientBrush(rectDown, clrDark, clrLight, LinearGradientMode.Vertical);
                gBar.FillRectangle(lgbDown, rectDown);

                Rectangle rectUp = new Rectangle(Point.Empty, new Size(intWidth, intHeight / 2 + 1));
                LinearGradientBrush lgbUp = new LinearGradientBrush(rectUp, clrLight, clrDark, LinearGradientMode.Vertical);
                gBar.FillRectangle(lgbUp, rectUp);


                int intMaskWidth = (int)(picBar.Width - (float)(picBar.Width / 100.00f) * Value);
                Color clrMaskDark = Color.FromArgb(160, 160, 160);
                Color clrMaskLight = Color.FromArgb(226, 226, 226);
                if (intMaskWidth > 0)
                {
                    Rectangle rectMaskDown = new Rectangle(new Point(intWidth, intHeight / 2), new Size(intMaskWidth, intHeight / 2));
                    LinearGradientBrush lgbMaskDown = new LinearGradientBrush(rectMaskDown, clrMaskLight, clrMaskDark, LinearGradientMode.Vertical);
                    gBar.FillRectangle(lgbMaskDown, rectMaskDown);

                    Rectangle rectMaskUp = new Rectangle(new Point(intWidth, 0), new Size(intMaskWidth, intHeight / 2 + 1));
                    LinearGradientBrush lgbMaskUp = new LinearGradientBrush(rectMaskUp, clrMaskDark, clrMaskLight, LinearGradientMode.Vertical);
                    gBar.FillRectangle(lgbMaskUp, rectMaskUp);

                }

            }
            else if (Value == 0)
            {
                int intMaskWidth = (int)(picBar.Width - (float)(picBar.Width / 100.00f) * Value);
                Color clrMaskDark = Color.FromArgb(160, 160, 160);
                Color clrMaskLight = Color.FromArgb(226, 226, 226);
                if (intMaskWidth > 0)
                {
                    Rectangle rectMaskDown = new Rectangle(new Point(intWidth, intHeight / 2), new Size(intMaskWidth, intHeight / 2));
                    LinearGradientBrush lgbMaskDown = new LinearGradientBrush(rectMaskDown, clrMaskLight, clrMaskDark, LinearGradientMode.Vertical);
                    gBar.FillRectangle(lgbMaskDown, rectMaskDown);

                    Rectangle rectMaskUp = new Rectangle(new Point(intWidth, 0), new Size(intMaskWidth, intHeight / 2 + 1));
                    LinearGradientBrush lgbMaskUp = new LinearGradientBrush(rectMaskUp, clrMaskDark, clrMaskLight, LinearGradientMode.Vertical);
                    gBar.FillRectangle(lgbMaskUp, rectMaskUp);

                }
            }
            Font fntStringFont = new Font("Consolas", 24);
            string strValue = Value.ToString() + "%";
            SizeF sfStringSize = gBar.MeasureString(strValue, fntStringFont);
            PointF pfCenter = new PointF(bBar.Width / 2 - sfStringSize.Width / 2, bBar.Height / 2 - sfStringSize.Height / 2);
            gBar.DrawString(strValue, fntStringFont, Brushes.Black, pfCenter);
            picBar.Image = bBar;
        }
        #endregion
        #region Properties
        public string NewVersion
        {

            get
            {
                if (intNewVersion[0].ToString() == "")
                {
                    return "";
                }
                else
                {
                    return intNewVersion[0].ToString() + "." + intNewVersion[1].ToString() + "." + intNewVersion[2].ToString() + "." + intNewVersion[3].ToString();
                }

            }
        }
        public string CurrentVersion
        {

            get
            {
                if (intCurrentVersion[0].ToString() == "")
                {
                    return "";
                }
                else
                {
                    return intCurrentVersion[0].ToString() + "." + intCurrentVersion[1].ToString() + "." + intCurrentVersion[2].ToString() + "." + intCurrentVersion[3].ToString();
                }

            }
        }
        public string DownloadLink
        {
            get
            {
                return strDownloadLink;
            }
        }
        /// <summary>
        /// Next, call ArticleLink
        /// or StartDownloadUpdates
        /// </summary>
        public bool IsCritical
        {
            get
            {
                return bIsCritical;
            }
        }
        public string NewDownloadLink
        {
            get
            {
                if (bIsCritical)
                {
                    return "";
                }
                return strDownloadLink;

            }
        }
        public string ArticleLink
        {
            get
            {
                if (bIsCritical)
                {
                    return strArticleLink;
                }
                return "";
            }
        }
        /// <summary>
        /// Next, call IsCritical
        /// </summary>
        public bool CheckForNewVersion
        {
            get
            {
                bGotUpdate = CompareVersion();
                return bGotUpdate;
            }

        }

        public string UpdateHistory
        {
            get
            {
                return strHistory;
            }
        }
        #endregion
    }
}
