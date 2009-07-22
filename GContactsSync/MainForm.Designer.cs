namespace GContactsSync
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.ntfIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.ctxNotificationIcon = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsAbort = new System.Windows.Forms.ToolStripMenuItem();
            this.tsStart = new System.Windows.Forms.ToolStripMenuItem();
            this.outlookToGmailToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gMailToOutlookToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsSep = new System.Windows.Forms.ToolStripSeparator();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tmrSynching = new System.Windows.Forms.Timer(this.components);
            this.imgSyncing = new System.Windows.Forms.ImageList(this.components);
            this.ctxNotificationIcon.SuspendLayout();
            this.SuspendLayout();
            // 
            // ntfIcon
            // 
            this.ntfIcon.ContextMenuStrip = this.ctxNotificationIcon;
            this.ntfIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("ntfIcon.Icon")));
            this.ntfIcon.Text = "GContactsSync";
            this.ntfIcon.Visible = true;
            this.ntfIcon.DoubleClick += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // ctxNotificationIcon
            // 
            this.ctxNotificationIcon.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsAbort,
            this.tsStart,
            this.tsSep,
            this.openToolStripMenuItem,
            this.toolStripMenuItem1,
            this.closeToolStripMenuItem});
            this.ctxNotificationIcon.Name = "contextMenuStrip1";
            this.ctxNotificationIcon.Size = new System.Drawing.Size(181, 142);
            this.ctxNotificationIcon.DoubleClick += new System.EventHandler(this.ctxNotificationIcon_DoubleClick);
            this.ctxNotificationIcon.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStrip1_Opening);
            // 
            // tsAbort
            // 
            this.tsAbort.Name = "tsAbort";
            this.tsAbort.Size = new System.Drawing.Size(180, 22);
            this.tsAbort.Text = "Abort Sync";
            this.tsAbort.Click += new System.EventHandler(this.tsAbort_Click);
            // 
            // tsStart
            // 
            this.tsStart.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.outlookToGmailToolStripMenuItem,
            this.gMailToOutlookToolStripMenuItem});
            this.tsStart.Name = "tsStart";
            this.tsStart.Size = new System.Drawing.Size(180, 22);
            this.tsStart.Text = "Full Sync";
            this.tsStart.Click += new System.EventHandler(this.tsStart_Click);
            // 
            // outlookToGmailToolStripMenuItem
            // 
            this.outlookToGmailToolStripMenuItem.Name = "outlookToGmailToolStripMenuItem";
            this.outlookToGmailToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.outlookToGmailToolStripMenuItem.Text = "Outlook to Gmail";
            this.outlookToGmailToolStripMenuItem.Click += new System.EventHandler(this.outlookToGmailToolStripMenuItem_Click);
            // 
            // gMailToOutlookToolStripMenuItem
            // 
            this.gMailToOutlookToolStripMenuItem.Name = "gMailToOutlookToolStripMenuItem";
            this.gMailToOutlookToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.gMailToOutlookToolStripMenuItem.Text = "GMail to Outlook";
            this.gMailToOutlookToolStripMenuItem.Click += new System.EventHandler(this.gMailToOutlookToolStripMenuItem_Click);
            // 
            // tsSep
            // 
            this.tsSep.Name = "tsSep";
            this.tsSep.Size = new System.Drawing.Size(177, 6);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.openToolStripMenuItem.Text = "Open Configuration";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.toolStripMenuItem1.Text = "Show Log";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // tmrSynching
            // 
            this.tmrSynching.Interval = 500;
            this.tmrSynching.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // imgSyncing
            // 
            this.imgSyncing.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSyncing.ImageStream")));
            this.imgSyncing.TransparentColor = System.Drawing.Color.Transparent;
            this.imgSyncing.Images.SetKeyName(0, "GCS.ico");
            this.imgSyncing.Images.SetKeyName(1, "GCSAnim1.ico");
            this.imgSyncing.Images.SetKeyName(2, "GCSAnim2.ico");
            this.imgSyncing.Images.SetKeyName(3, "GCSAnim3.ico");
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(330, 312);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Google Outlook Contacts Sync";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.ctxNotificationIcon.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon ntfIcon;
        private System.Windows.Forms.ContextMenuStrip ctxNotificationIcon;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.Timer tmrSynching;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ImageList imgSyncing;
        private System.Windows.Forms.ToolStripMenuItem tsAbort;
        private System.Windows.Forms.ToolStripSeparator tsSep;
        private System.Windows.Forms.ToolStripMenuItem tsStart;
        private System.Windows.Forms.ToolStripMenuItem outlookToGmailToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem gMailToOutlookToolStripMenuItem;
    }
}

