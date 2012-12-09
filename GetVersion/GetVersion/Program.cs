using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace GetVersion
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            StreamReader srXML = new StreamReader(args[0]);
            string[] strTmp = System.Text.RegularExpressions.Regex.Split(srXML.ReadToEnd(), Environment.NewLine);
            foreach (string strTemp in strTmp)
            {
                if (strTemp == "") continue;
                if (strTemp.Substring(0, 1) == "[" && strTemp.Contains("AssemblyVersion"))
                {
                    XmlTextWriter xtwXML = new XmlTextWriter(Application.StartupPath + @"\Version.xml", Encoding.UTF8);
                    xtwXML.Formatting = Formatting.Indented;

                    string strVersion = strTemp.Replace("[assembly: AssemblyVersion(\"", "").Replace("\")]", "");
                    string[] strVer = strVersion.Split('.');
                    xtwXML.WriteStartElement("Version");

                    xtwXML.WriteStartElement("Major");
                    xtwXML.WriteString(strVer[0]);
                    xtwXML.WriteEndElement();

                    xtwXML.WriteStartElement("Minor");
                    xtwXML.WriteString(strVer[1]);
                    xtwXML.WriteEndElement();

                    xtwXML.WriteStartElement("Build");
                    xtwXML.WriteString(strVer[2]);
                    xtwXML.WriteEndElement();

                    xtwXML.WriteStartElement("Revision");
                    xtwXML.WriteString(strVer[3]);
                    xtwXML.WriteEndElement();

                    xtwXML.WriteEndElement();
                    xtwXML.Close();
                    break;
                }
            }



        }
    }
}
