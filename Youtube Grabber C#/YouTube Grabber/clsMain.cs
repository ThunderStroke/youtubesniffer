using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web;
using System.Text.RegularExpressions;
using System.Threading;

namespace YouTube_Grabber
{
    class clsMain : IDisposable
    {
        #region PublicVariable
        /// <summary>
        /// Event's progress structure
        /// </summary>
        public struct ProgressEventArgs
        {
            public string Status;
            public List<string> strDownloadLinks;
            public List<string> strQuality;
            public List<string> strType;
            public List<string> strSize;
        }
        /// <summary>
        /// Video info structure
        /// </summary>
        public struct VideoInfoEventArgs
        {
            public string strTitle;
            public string strAuthor;
            public string strPic;
            public string strLikes;
            public string strDislikes;
            public string strViews;
            public string strLength;

        }
        public delegate void ProgressEventHandler(ProgressEventArgs Status);
        public delegate void VideoInfoEventHandler(VideoInfoEventArgs VideoInfo);
        public delegate void ErrorHandler(string err);

        /// <summary>
        /// Event to report progress
        /// </summary>
        public event ProgressEventHandler Progress;
        /// <summary>
        /// Event occurs when download links processed
        /// </summary>
        public event ProgressEventHandler DownloadLinksDone;
        /// <summary>
        /// Event occurs when video's info received
        /// </summary>
        public event VideoInfoEventHandler VideoInfoReceived;
        /// <summary>
        /// Handles the error
        /// </summary>
        public event ErrorHandler Error;
        #endregion
        #region PrivateVariable
        ProgressEventArgs peaStatus = new ProgressEventArgs();
        VideoInfoEventArgs veaInfo = new VideoInfoEventArgs();
        WebClient wbDown = new WebClient();
        string strVideoUrl = "";
        string strHtml = "";
        string strDecode = "";
        Thread tDownloadLinks;
        Thread tGetVideoInfo;
        #endregion
        #region PublicFunction

        /// <summary>
        /// YouTube Video Grabber Class
        /// </summary>
        /// <param name="strYouTubeVideoUrl">The Video Link Of A YouTube Video</param>
        public clsMain(string strYouTubeVideoUrl)
        {
            strVideoUrl = strYouTubeVideoUrl;
            peaStatus.strDownloadLinks = new List<string>();
            peaStatus.strQuality = new List<string>();
            peaStatus.strType = new List<string>();
            peaStatus.strSize = new List<string>();
        }

        /// <summary>
        /// Download and Decode the Webpage
        /// </summary>
        public void Start()
        {
            try
            {
                ReportProgress("Started");
                wbDown.Encoding = System.Text.Encoding.UTF8;
                //wbDown.Headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (iPad; U; CPU OS 3_2 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Version/4.0.4 Mobile/7B334b Safari/531.21.102011-10-16 20:23:50";
                if (strVideoUrl.Substring(0, 4) != "http") 
                {
                strVideoUrl = "http://" + strVideoUrl;
                }


                wbDown.DownloadStringAsync(new Uri(strVideoUrl));
                wbDown.DownloadStringCompleted += new DownloadStringCompletedEventHandler(wbDown_DownloadStringCompleted);
            }
            catch (Exception ex)
            {
                Error(ex.Message);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        #endregion
        #region PrivateFunction
        void wbDown_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            try
            {
                ReportProgress("Webpage source code recieved");
                strHtml = e.Result.ToString();
                strHtml = Regex.Replace(strHtml, @"\r\n?|\n", "");
                strDecode = HttpUtility.UrlDecode(HttpUtility.UrlDecode(HttpUtility.UrlDecode(strHtml.Substring(strHtml.IndexOf("var swf = "), strHtml.IndexOf(".innerHTML = swf") - strHtml.IndexOf("var swf = ")))));
                System.Diagnostics.Debug.Print(strDecode);
                tDownloadLinks = new Thread(ProcessDownloadLinks);
                tGetVideoInfo = new Thread(GetVideoInfo);
                tDownloadLinks.Start();
                tGetVideoInfo.Start();
            }
            catch (Exception ex)
            {
                Error(ex.Message);
            }

            //ProcessDownloadLinks();
        }
        void ProcessDownloadLinks()
        {
                string[] strDownloadLinks = Regex.Split(strDecode, "&url=http://");

                for (int i = 1; i < strDownloadLinks.Count(); i++)
                {
                    {
                        //if (!strDownloadLinks[i].Contains("o-o---preferred---")) continue;
 
                        if (strDownloadLinks[i].Contains(@"\u0026amp;"))
                        {
                            strDownloadLinks[i] = strDownloadLinks[i].Substring(0, strDownloadLinks[i].IndexOf(@"\u0026amp;"));
                        }
                        strDownloadLinks[i] = strDownloadLinks[i].Replace("&sig=", "&signature=");
                        strDownloadLinks[i] = "http://" + strDownloadLinks[i];

                        GetVideoType(strDownloadLinks[i]);
                        peaStatus.strDownloadLinks.Add(strDownloadLinks[i]);
                        string strQuality = RegexMatch(strDownloadLinks[i], "&quality=([0-9a-z]+)");
                        if (strQuality.Contains("large"))
                        {
                            strQuality = strQuality.Replace("large", "480p");
                        }
                        else if (strQuality.Contains("medium"))
                        {
                            strQuality = strQuality.Replace("medium", "320p");
                        }
                        else if (strQuality.Contains("small"))
                        {
                            strQuality = strQuality.Replace("small", "240p");
                        }
                        peaStatus.strQuality.Add(strQuality);
                           
                    }
                    ReportProgress("Preparing video download links (This may take a while)", false);
                }

                ReportProgress("Donwload links done", true);
                tDownloadLinks.Abort();
            //GetVideoInfo();
        }

        void GetVideoType(string url)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "HEAD";
                //req.Timeout = 100000;
                HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
                string strType = "*." + resp.ContentType.Replace("video/", "");
                if (strType.Contains("x-flv"))
                {
                    strType =strType.Replace("x-flv", "flv");
                }
                else if (strType.Contains("3gpp"))
                {
                    strType = strType.Replace("3gpp", "3gp");
                }

                peaStatus.strType.Add(strType);
                

                peaStatus.strSize.Add(formatBytes(resp.ContentLength));
                req.Abort();
            }
            catch (System.Net.WebException ex)
            {
                Error("Ewww.. Error... Exiting.");
                throw ex;
            }
        }

        string RegexMatch(string strMatch, string pattern)
        {
            Match m = Regex.Match(strMatch, pattern);
            if (m.Success)
            {
                return m.Groups[1].Value;
            }
            else
            {
                return "";
            }
        }

        void ReportProgress(string Status, bool bDone = false)
        {
            Monitor.Enter(this);
            peaStatus.Status = Status;
            Progress(peaStatus);
            if (bDone == true)
            {
                DownloadLinksDone(peaStatus);

            }
            Monitor.Exit(this);
        }

        string formatBytes(float bytes)
        {
            string[] Suffix = { "B", "KB", "MB", "GB", "TB" };
            int i;
            double dblSByte = 0;
            for (i = 0; (int)(bytes / 1024) > 0; i++, bytes /= 1024)
                dblSByte = bytes / 1024.0;
            return String.Format("{0:0.00} {1}", dblSByte, Suffix[i]);
        }

        #region VideoInfoFunc
        void GetVideoInfo()
        {

            veaInfo.strTitle = GetVideoTitle();
            veaInfo.strAuthor = GetVideoAuthor();
            veaInfo.strPic = GetVideoPic();
            veaInfo.strLikes = GetVideoLikes();
            veaInfo.strDislikes = GetVideoDislikes();
            veaInfo.strViews = GetVideoViews();
            veaInfo.strLength = GetVideoLength();
            ReportProgress("Video info recieved", false);
            VideoInfoReceived(veaInfo);
            //ReportProgress("one", false);
            tGetVideoInfo.Abort();

        }
        string GetVideoTitle()
        {
            string strVideoTitle = RegexMatch(strHtml, @"<title>\s*(.+?)\s*</title>");
            strVideoTitle = HttpUtility.HtmlDecode(strVideoTitle);
            strVideoTitle = Regex.Replace(strVideoTitle, "[\\/?:*\"><|]+", "", RegexOptions.Compiled);
            return strVideoTitle;
        }
        string GetVideoAuthor()
        {
            string strVideoAuthor = strHtml.Substring(strHtml.IndexOf("<span itemprop=\"author\" itemscope itemtype=\"http://schema.org/Person\"> "), strHtml.IndexOf("</span>", strHtml.IndexOf("<span itemprop=\"author\" itemscope itemtype=\"http://schema.org/Person\"> ")) - strHtml.IndexOf("<span itemprop=\"author\" itemscope itemtype=\"http://schema.org/Person\"> "));
            strVideoAuthor = strVideoAuthor.Replace("<span itemprop=\"author\" itemscope itemtype=\"http://schema.org/Person\"> ", "").Replace("<link itemprop=\"url\"", "").Replace("href=\"", "").Replace("\">", "").Replace("</span>", "").Trim();
            return strVideoAuthor;
        }
        string GetVideoPic()
        {
            int i = strHtml.IndexOf("<link itemprop=\"thumbnailUrl\"");
            int f = strHtml.IndexOf("\">", i) - i;
            string strVideoPic = strHtml.Substring(i, f).Replace("<link itemprop=\"thumbnailUrl\"", "").Replace("href=\"", "").Trim();
            return strVideoPic;
        }
        string GetVideoLikes()
        {
            try
            {
                int i = strHtml.IndexOf("<span class=\"likes\">");
                int h = strHtml.IndexOf("</span>", i) - i;
                string strVideoLikes = strHtml.Substring(i, h).Replace("<span class=\"likes\">", "").Replace("</span>", "").Trim();
                return strVideoLikes;
            }
            catch
            {
                return "Unable to receive video's likes number";
            }
        }
        string GetVideoDislikes()
        {
            try
            {
                int i = strHtml.IndexOf("<span class=\"dislikes\">");
                int h = strHtml.IndexOf("</span>", i) - i;
                string strVideoDislikes = strHtml.Substring(i, h).Replace("<span class=\"dislikes\">", "").Replace("</span>", "").Trim();
                return strVideoDislikes;
            }
            catch
            {
                return "Unable to receive video's dislikes number";
            }
        }
        string GetVideoViews()
        {
            int i = strHtml.IndexOf("<span class=\"watch-view-count\">");
            int h = strHtml.IndexOf("</strong>", i) - i;
            string strVideoViews = strHtml.Substring(i, h).Replace("<span class=\"watch-view-count\">", "").Replace("<strong>", "").Replace("</strong>", "").Trim();
            return strVideoViews;
        }
        string GetVideoLength()
        {
            string strVideoLength = RegexMatch(strDecode, @"length_seconds=([0-9]+)\\");
            int intVideoLength = Convert.ToInt32(strVideoLength);
            strVideoLength = String.Format("{0:00}:{1:00}:{2:00}", intVideoLength / 3600, (intVideoLength / 60) % 60, intVideoLength % 60);
            return strVideoLength;
        }

        protected virtual void Dispose(bool bDispose)
        {
            if (bDispose)
            {
                tDownloadLinks.Abort();
                tGetVideoInfo.Abort();
            }
        }


        #endregion
        #endregion


    }
}
