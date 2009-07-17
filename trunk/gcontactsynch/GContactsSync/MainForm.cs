using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.Contacts;
using Microsoft.Office.Interop.Outlook;
using System.Threading;
using GContactsSync.Properties;

namespace GContactsSync
{
    public partial class MainForm : Form
    {
        bool ShouldClose = false;
        Thread synchThread;
        int indAnim;

        public MainForm()
        {
            InitializeComponent();
            synchThread = new Thread(new ThreadStart(StartThread));
        }

        private void OContactFetched (object sender, ContactItem contact)
        {
            
        }

        private void ContactFeched(object sender, Contact contact)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }


        private delegate void ShowErrorDelegate(string msg);
        private void ShowError (string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new ShowErrorDelegate(ShowError), msg);
            }
            else
            {
                this.Show();
                MessageBox.Show(msg, "Error");
            }
        }

        public void SyncError (object sender, System.Exception ex)
        {
            ShowError(ex.Message); 
        }

        private delegate void StartEndSyncingDelegate(object sender);

        public void StartSyncing(object sender)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new StartEndSyncingDelegate(StartSyncing), sender);
            }
            else
            {
                txtLog.AppendText("---Starting to sync...\r\n");
                ntfIcon.Text = "Starting to Synchronize";
                //tmrSyncing.Interval = 100;
                tmrSyncing.Enabled = true;
            }
        }
        private void EndSyncing(object sender)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new StartEndSyncingDelegate(EndSyncing), sender);
            }
            else
            {
                tmrSyncing.Enabled = false;
                indAnim = 0;
                ntfIcon.Icon = this.Icon;
                ntfIcon.Text = "GContactsSync";
            }
        }

        private void StartThread()
        {
            ContactSync cs = new ContactSync(txtUser.Text, txtPassword.Text,Convert.ToInt32(numInterval.Value*60000));
            cs.GoogleSynched += GoogleSynched;
            cs.OutlookSynched += OutlookSynched;
            cs.StartSyncing += StartSyncing;
            cs.EndSyncing += EndSyncing;
            cs.Error += SyncError;
            if (rbGoogleToOutlook.Checked)
            {
                cs.SyncToOutlook();
            }
            else
            {
                cs.SyncToGoogle();
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveSettings();
            
            if (synchThread.ThreadState == ThreadState.Running)
            {
                synchThread.Abort();
                txtLog.Clear();
            }
            Hide();
            StartSync();
            
        }

        private void SaveSettings()
        {
            if (rbGoogleToOutlook.Checked)
            {
                Settings.Default.Direction = 0;
            }
            else
            {
                Settings.Default.Direction = 1;
            }
            Settings.Default.Interval = Convert.ToInt32(numInterval.Value);
            Settings.Default.User = txtUser.Text;
            Settings.Default.Password = EncryptDecrypt.Encrypt(txtPassword.Text);
            Settings.Default.Save();
        }
        
        public void StartSync()
        {
            txtLog.Clear();
            synchThread = new Thread(new ThreadStart(StartThread));
            synchThread.Start();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!ShouldClose)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShouldClose = true;
            Close();
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private delegate void AddLogTextDelegate(String Text);

        private void AddLogText(String Text)
        {
            if (this.txtLog.InvokeRequired)
            {
                // This is a worker thread so delegate the task.
                this.txtLog.Invoke(new AddLogTextDelegate(this.AddLogText), Text);
            }
            else
            {
                // This is the UI thread so perform the task.
                this.txtLog.AppendText(Text+"\r\n");
                this.txtLog.Select(txtLog.Text.Length, 0);
            }
        }

        public void OutlookSynched(object sender, ContactItem contact, int current, int total)
        {
            ntfIcon.Text = "Google->Outlook " + current.ToString() + " of " + total.ToString();
            AddLogText("Google->Outlook " + current.ToString() + " of " + total.ToString());
        }
        public void GoogleSynched(object sender, Contact contact, int current, int total)
        {
            ntfIcon.Text = "Outlook->Google " + current.ToString() + " of " + total.ToString();
            AddLogText("Outlook->Google " + current.ToString() + " of " + total.ToString());
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ntfIcon.Icon = Icon.FromHandle(((Bitmap)imgSyncing.Images[indAnim]).GetHicon());
            indAnim = (indAnim + 1) % imgSyncing.Images.Count;
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            txtUser.Focus();
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void tpLog_Enter(object sender, EventArgs e)
        {
            txtLog.Focus();
            txtLog.Select(txtLog.Text.Length, 0);
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex=1;
            Show();
        }

        private void tabControl1_TabIndexChanged_1(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                txtLog.Focus();
                txtLog.Select(txtLog.Text.Length, 0);
            }
        }

    }
}
