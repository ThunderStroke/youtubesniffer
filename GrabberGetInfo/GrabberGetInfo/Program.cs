using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace GrabberGetInfo
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            
            if (args.Count() == 0)
            {
                Environment.Exit(-1);
            }

            frmMain.strInfoPath = args;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmMain());
        }
    }
}
