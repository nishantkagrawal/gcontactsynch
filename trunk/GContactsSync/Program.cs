using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using GContactsSync.Properties;
using GContactsSyncLib;
using Google.Contacts;

namespace GContactsSync
{
    static class Program
    {
        private delegate void SimpleDelegate();
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainForm frm = new MainForm();
            Config.CheckShowConfig();
            List<Contact> c = new GoogleAdapter(Config.Username, Config.Password).Contacts;
            frm.Invoke(new SimpleDelegate(frm.StartSync), null);
            Application.Run();
        }
    }
}
