using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using GContactsSync.Properties;

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
            if (Settings.Default.Direction == 0)
            {
                frm.rbGoogleToOutlook.Checked = true;
            }
            else
            {
                frm.rbOutlookToGoogle.Checked = true;
            }
            frm.numInterval.Value = Settings.Default.Interval;
            frm.txtUser.Text = Settings.Default.User;
            if (Settings.Default.Password != "")
            {
                frm.txtPassword.Text = EncryptDecrypt.Decrypt(Settings.Default.Password);
            }
            IntPtr h = frm.Handle;
            if (frm.txtUser.Text == "")
                frm.Show();
            else
            {
                frm.Invoke(new SimpleDelegate(frm.StartSync), null);
            }
            Application.Run();
        }
    }
}
