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
using GContactsSyncLib;
using NLog;

namespace GContactsSync
{
    public partial class MainForm : Form
    {
        static Logger logger = ContactSync.logger;
        bool ShouldClose = false;
        Thread synchThread;
        Thread fullSyncThread;
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

        public void SynchError (object sender, System.Exception ex)
        {
            ShowError(ex.Message); 
        }

        private delegate void StartEndSynchingDelegate(object sender);

        public void StartSynching(object sender)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new StartEndSynchingDelegate(StartSynching), sender);
            }
            else
            {
                ntfIcon.Text = "Starting to Synchronize";
                //tmrSyncing.Interval = 100;
                tmrSynching.Enabled = true;
            }
        }
        private void EndSynching(object sender)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new StartEndSynchingDelegate(EndSynching), sender);
            }
            else
            {
                tmrSynching.Enabled = false;
                indAnim = 0;
                ntfIcon.Icon = this.Icon;
                ntfIcon.Text = "GContactsSync";
            }
        }

        private void StartThread()
        {
            ContactSync cs = new ContactSync(Config.Username, Config.Password,Convert.ToInt32(Config.Interval*60000));
            cs.GoogleSynched += GoogleSynched;
            cs.OutlookSynched += OutlookSynched;
            cs.StartSynching += StartSynching;
            cs.EndSynching += EndSynching;
            cs.Error += SynchError;
            cs.SyncFromGoogle();
        }
        
        public void StartSync()
        {
            synchThread = new Thread(new ThreadStart(StartThread));
            synchThread.Start();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Config.ShowConfig();
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
            tsAbort.Available = (synchThread.ThreadState == ThreadState.Running);
            tsSep.Visible = tsAbort.Available;
            tsStart.Visible = !tsAbort.Available;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShouldClose = true;
            Close();
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        

        public void OutlookSynched(object sender, ContactItem contact, int current, int total)
        {
            ntfIcon.Text = "Google->Outlook " + current.ToString() + " of " + total.ToString();
            //logger ("Google->Outlook " + current.ToString() + " of " + total.ToString());
        }
        public void GoogleSynched(object sender, Contact contact, int current, int total)
        {
            ntfIcon.Text = "Outlook->Google " + current.ToString() + " of " + total.ToString();
            //AddLogText("Outlook->Google " + current.ToString() + " of " + total.ToString());
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ntfIcon.Icon = Icon.FromHandle(((Bitmap)imgSyncing.Images[indAnim]).GetHicon());
            indAnim = (indAnim + 1) % imgSyncing.Images.Count;
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void tpLog_Enter(object sender, EventArgs e)
        {
            
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Show();
        }

        private void tabControl1_TabIndexChanged_1(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
        }

        private void tsAbort_Click(object sender, EventArgs e)
        {
            if (synchThread.ThreadState == ThreadState.Running)
            {
                synchThread.Abort();
            }
        }

        private void tsStart_Click(object sender, EventArgs e)
        {
            
        }

        private void ctxNotificationIcon_DoubleClick(object sender, EventArgs e)
        {

        }

        private void outlookToGmailToolStripMenuItem_Click(object sender, EventArgs e)
        {

            fullSyncThread = new Thread(new ParameterizedThreadStart(StartFullSync));
            fullSyncThread.Start(Config.Direction.dirToGoogle);
            
        }

        private void StartFullSync(object direction)
        {
            Config.Direction dir = (Config.Direction)direction;
        }

    }
}
