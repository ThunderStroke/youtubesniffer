using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace GrabberGetInfo
{
    public partial class frmMain : Form
    {
        public static string[] strInfoPath;
        public frmMain()
        {
            
            InitializeComponent();
        }
        string[] strVersion;
        string strBuildDate;
        private void frmMain_Load(object sender, EventArgs e)
        {
            strBuildDate = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString() ;
            txtDate.Text = strBuildDate;
            txtDL.Text = @"https://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/C%23/YouTube%20Video%20Grabber/YouTube%20Grabber.exe";
            StreamReader srXML = new StreamReader(strInfoPath[0]);
            string[] strTmp = Regex.Split(srXML.ReadToEnd(), Environment.NewLine);
            foreach (string strTemp in strTmp)
            {
                if (strTemp == "") continue;
                if (strTemp.Substring(0, 1) == "[" && strTemp.Contains("AssemblyVersion"))
                {
                    string strVer = strTemp.Replace("[assembly: AssemblyVersion(\"", "").Replace("\")]", "");
                    txtVer.Text = strVer;
                    strVersion = strVer.Split('.');
                    break;
                }
            }
            srXML.Close();
            txtHistory.Text += txtVer.Text + Environment.NewLine + "- ";
            TextBoxFocus();

        }
        XmlTextWriter xtwXML = new XmlTextWriter(Application.StartupPath + @"\Settings.xml", Encoding.UTF8);

        private void btnDone_Click(object sender, EventArgs e)
        {
            xtwXML.Formatting = Formatting.Indented;
            if (txtHistory.Text != "")
            {
                xtwXML.WriteStartElement("Settings");
                WriteVersion();
                WriteDownloadLink();
                WriteHistory();
                WriteCriticalUpdate();

                xtwXML.WriteEndElement();
                xtwXML.Close();
                Environment.Exit(-1);

            }

        }

        void WriteCriticalUpdate()
        {
            xtwXML.WriteStartElement("CriticalUpdate");
            xtwXML.WriteStartElement("IsCriticleUpdate");
            xtwXML.WriteString(cbCritical.Checked.ToString());
            xtwXML.WriteEndElement();
            if (cbCritical.Checked == true)
            {
                xtwXML.WriteStartElement("ArticleLink");
                xtwXML.WriteString(txtArticleURL.Text);
                xtwXML.WriteEndElement();
            }

            xtwXML.WriteEndElement();
        }


        void WriteHistory()
        {
            xtwXML.WriteStartElement("History");
            xtwXML.WriteString(txtHistory.Text);
            xtwXML.WriteEndElement();

            if (ckDontUpdateLog.Checked == true) return;

            string strToWrite = txtHistory.Text;
            bool bNeedWriteDate = true;
            if (File.Exists("Update Log.txt"))
            {

                StreamReader srHistory = new StreamReader("Update Log.txt");
                string strUpdate = srHistory.ReadToEnd();
                string[] strTemp = Regex.Split(strUpdate, Environment.NewLine);
                
                foreach (string strCurrent in strTemp)
                {
                    if (strCurrent.Trim() == strBuildDate)
                    {
                        //txtHistory.Text = txtHistory.Text.Replace(strBuildDate + Environment.NewLine, "");
                        bNeedWriteDate =false ;
                        break;
                    }
                }
                srHistory.Close();
            }
            if (bNeedWriteDate == true)
            {
                strToWrite = Environment.NewLine +  strBuildDate + Environment.NewLine + strToWrite;
            }

            StreamWriter swHistory = new StreamWriter("Update Log.txt", true);
            swHistory.WriteLine(strToWrite);
            swHistory.Close();
        }

        void WriteDownloadLink()
        {
            xtwXML.WriteStartElement("DownloadLink");
            xtwXML.WriteString(txtDL.Text);
            xtwXML.WriteEndElement();
        }

        void WriteVersion()
        {
            xtwXML.WriteStartElement("Version");

            xtwXML.WriteStartElement("Major");
            xtwXML.WriteString(strVersion[0]);
            xtwXML.WriteEndElement();

            xtwXML.WriteStartElement("Minor");
            xtwXML.WriteString(strVersion[1]);
            xtwXML.WriteEndElement();

            xtwXML.WriteStartElement("Build");
            xtwXML.WriteString(strVersion[2]);
            xtwXML.WriteEndElement();

            xtwXML.WriteStartElement("Revision");
            xtwXML.WriteString(strVersion[3]);
            xtwXML.WriteEndElement();

            xtwXML.WriteEndElement();

        }

        void TextBoxFocus()
        {
            txtHistory.Focus();
            txtHistory.SelectionStart = txtHistory.Text.Length;
        }

        private void btnCore_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[Core Updated]\t";
            TextBoxFocus();
        }

        private void btnNewUI_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[New UI]\t";
            TextBoxFocus();
        }

        private void btnNewFunc_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[New Function]\t";
            TextBoxFocus();
        }

        private void btnUIUpdated_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[UI Updated]\t";
            TextBoxFocus();
        }

        private void btnBugFixed_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[Bug Fixed]\t";
            TextBoxFocus();
        }

        private void btnFuncRemove_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[Function Removed]\t";
            TextBoxFocus();
        }

        private void btnUIRemoved_Click(object sender, EventArgs e)
        {
            txtHistory.Text += "[UI Removed]\t";
            TextBoxFocus();
        }

        private void txtDate_TextChanged(object sender, EventArgs e)
        {
            strBuildDate = txtDate.Text;
        }

        private void cbCritical_CheckedChanged(object sender, EventArgs e)
        {
            if (cbCritical.Checked)
            {
                txtArticleURL.Enabled = true;
            }
        }

    }
}
