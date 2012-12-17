using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Diagnostics;

namespace YouTube_Grabber
{
    class clsDownloader : WebClient
    {
        int intCurrentClassIndex;
        public clsDownloader(int Index)
        {
            intCurrentClassIndex = Index;
        }

        public delegate void OnDownloadFileCompletedIndexEventHandler(System.ComponentModel.AsyncCompletedEventArgs e, int Index);
        public event OnDownloadFileCompletedIndexEventHandler DownloadFileCompletedIndex;

        public delegate void OnDownloadProgressChangedIndexEventHandler(DownloadProgressChangedEventArgs e, int Index, string time, string speed);
        public event OnDownloadProgressChangedIndexEventHandler DownloadProgressChangedIndex;

        DateTime dt;
        double time;

        Stopwatch sw = new Stopwatch();

        protected override void OnDownloadProgressChanged(DownloadProgressChangedEventArgs e)
        {
            if (dt.Ticks == 0)
            {
                dt = DateTime.Now;
            }
            time = (DateTime.Now.Ticks - dt.Ticks) / (TimeSpan.TicksPerMillisecond * 1000);



            if (sw.ElapsedMilliseconds == 0)
            {
                sw.Start();
            }

            float dbSec = sw.ElapsedMilliseconds / 1000;
            string strSpeed;
            try
            {
                strSpeed = clsMain.formatBytes(e.BytesReceived / dbSec);
            }
            catch
            {
                strSpeed = "";
            }

            //System.Diagnostics.Debug.Print(((DateTime.Now.Ticks - dt.Ticks) / (TimeSpan.TicksPerMillisecond*1000) ).ToString());
            //System.Diagnostics.Debug.Print(((time / 60) % 60).ToString());
            //System.Diagnostics.Debug.Print((e.BytesReceived.ToString() + "\t" + (DateTime.Now.Ticks - dt.Ticks).ToString()));
            DownloadProgressChangedIndex(e, intCurrentClassIndex, String.Format("{0:00}:{1:00}:{2:00}", (int)time / 3600, (int)((time / 60) % 60), (int)(time % 60)), strSpeed);

        }
        protected override void OnDownloadFileCompleted(System.ComponentModel.AsyncCompletedEventArgs e)
        {
            DownloadFileCompletedIndex(e, intCurrentClassIndex);


            sw.Reset();
        }


    }
}
