using System.Windows.Forms;
namespace GContactsSyncLib
{
    partial class ConfigForm 
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfigForm));
            this.button1 = new System.Windows.Forms.Button();
            this.ntfIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.ctxNotificationIcon = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsAbort = new System.Windows.Forms.ToolStripMenuItem();
            this.tsStart = new System.Windows.Forms.ToolStripMenuItem();
            this.tsSep = new System.Windows.Forms.ToolStripSeparator();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tmrSynching = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.numInterval = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.rbOutlookToGoogle = new System.Windows.Forms.RadioButton();
            this.rbGoogleToOutlook = new System.Windows.Forms.RadioButton();
            this.label6 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tpConfig = new System.Windows.Forms.TabPage();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.imgSyncing = new System.Windows.Forms.ImageList(this.components);
            this.ctxNotificationIcon.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numInterval)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tpConfig.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(218, 232);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ntfIcon
            // 
            this.ntfIcon.ContextMenuStrip = this.ctxNotificationIcon;
            this.ntfIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("ntfIcon.Icon")));
            this.ntfIcon.Text = "GContactsSync";
            this.ntfIcon.Visible = true;
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
            this.ctxNotificationIcon.Size = new System.Drawing.Size(181, 120);
            // 
            // tsAbort
            // 
            this.tsAbort.Name = "tsAbort";
            this.tsAbort.Size = new System.Drawing.Size(180, 22);
            this.tsAbort.Text = "Abort Sync";
            // 
            // tsStart
            // 
            this.tsStart.Name = "tsStart";
            this.tsStart.Size = new System.Drawing.Size(180, 22);
            this.tsStart.Text = "Sync Now";
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
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.toolStripMenuItem1.Text = "Show Log";
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.closeToolStripMenuItem.Text = "Close";
            // 
            // tmrSynching
            // 
            this.tmrSynching.Interval = 500;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(23, 135);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Synchronization Interval:";
            // 
            // numInterval
            // 
            this.numInterval.Location = new System.Drawing.Point(172, 135);
            this.numInterval.Name = "numInterval";
            this.numInterval.Size = new System.Drawing.Size(49, 20);
            this.numInterval.TabIndex = 2;
            this.numInterval.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(223, 135);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "minutes";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(17, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(264, 126);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Google Account";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel2);
            this.panel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(232)))), ((int)(((byte)(238)))), ((int)(((byte)(250)))));
            this.panel1.Location = new System.Drawing.Point(6, 19);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(252, 100);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(232)))), ((int)(((byte)(238)))), ((int)(((byte)(250)))));
            this.panel2.Controls.Add(this.txtPassword);
            this.panel2.Controls.Add(this.txtUser);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(244, 93);
            this.panel2.TabIndex = 0;
            // 
            // txtPassword
            // 
            this.txtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPassword.Location = new System.Drawing.Point(78, 49);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(132, 20);
            this.txtPassword.TabIndex = 1;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // txtUser
            // 
            this.txtUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUser.Location = new System.Drawing.Point(78, 23);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(132, 20);
            this.txtUser.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label4.Location = new System.Drawing.Point(17, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Password:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label3.Location = new System.Drawing.Point(17, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Username:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(23, 160);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(112, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Update* Direction:";
            // 
            // rbOutlookToGoogle
            // 
            this.rbOutlookToGoogle.AutoSize = true;
            this.rbOutlookToGoogle.Checked = true;
            this.rbOutlookToGoogle.Location = new System.Drawing.Point(46, 176);
            this.rbOutlookToGoogle.Name = "rbOutlookToGoogle";
            this.rbOutlookToGoogle.Size = new System.Drawing.Size(150, 17);
            this.rbOutlookToGoogle.TabIndex = 3;
            this.rbOutlookToGoogle.TabStop = true;
            this.rbOutlookToGoogle.Text = "Outlook overwrites Google";
            this.rbOutlookToGoogle.UseVisualStyleBackColor = true;
            // 
            // rbGoogleToOutlook
            // 
            this.rbGoogleToOutlook.AutoSize = true;
            this.rbGoogleToOutlook.Location = new System.Drawing.Point(46, 199);
            this.rbGoogleToOutlook.Name = "rbGoogleToOutlook";
            this.rbGoogleToOutlook.Size = new System.Drawing.Size(150, 17);
            this.rbGoogleToOutlook.TabIndex = 4;
            this.rbGoogleToOutlook.Text = "Google overwrites Outlook";
            this.rbGoogleToOutlook.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 258);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(230, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "* contact creation will happen on both directions";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tpConfig);
            this.tabControl1.Location = new System.Drawing.Point(3, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(324, 307);
            this.tabControl1.TabIndex = 10;
            // 
            // tpConfig
            // 
            this.tpConfig.Controls.Add(this.groupBox1);
            this.tpConfig.Controls.Add(this.button1);
            this.tpConfig.Controls.Add(this.label6);
            this.tpConfig.Controls.Add(this.label1);
            this.tpConfig.Controls.Add(this.rbGoogleToOutlook);
            this.tpConfig.Controls.Add(this.numInterval);
            this.tpConfig.Controls.Add(this.rbOutlookToGoogle);
            this.tpConfig.Controls.Add(this.label2);
            this.tpConfig.Controls.Add(this.label5);
            this.tpConfig.Location = new System.Drawing.Point(4, 22);
            this.tpConfig.Name = "tpConfig";
            this.tpConfig.Padding = new System.Windows.Forms.Padding(3);
            this.tpConfig.Size = new System.Drawing.Size(316, 281);
            this.tpConfig.TabIndex = 0;
            this.tpConfig.Text = "Configuration";
            this.tpConfig.UseVisualStyleBackColor = true;
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
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
            // ConfigForm
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(330, 312);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConfigForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Google Outlook Contacts Sync";
            this.Load += new System.EventHandler(this.ConfigForm_Load);
            this.ctxNotificationIcon.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numInterval)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tpConfig.ResumeLayout(false);
            this.tpConfig.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.NotifyIcon ntfIcon;
        private System.Windows.Forms.ContextMenuStrip ctxNotificationIcon;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.Timer tmrSynching;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tpConfig;
        public System.Windows.Forms.NumericUpDown numInterval;
        public System.Windows.Forms.TextBox txtUser;
        public System.Windows.Forms.TextBox txtPassword;
        public System.Windows.Forms.RadioButton rbOutlookToGoogle;
        public System.Windows.Forms.RadioButton rbGoogleToOutlook;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ImageList imgSyncing;
        private System.Windows.Forms.ToolStripMenuItem tsAbort;
        private System.Windows.Forms.ToolStripSeparator tsSep;
        private System.Windows.Forms.ToolStripMenuItem tsStart;
    }
}

