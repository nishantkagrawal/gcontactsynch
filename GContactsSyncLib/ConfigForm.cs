using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace GContactsSyncLib
{
    public partial class ConfigForm : Form
    {
        
        public ConfigForm()
        {
            InitializeComponent();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            SaveSettings();    
        }

        private void SaveSettings()
        {
            if (rbGoogleToOutlook.Checked)
            {
                Config.SyncDirection = Config.Direction.dirToOutlook;
            }
            else
            {
                Config.SyncDirection = Config.Direction.dirToGoogle;
            }
            Config.Interval = Convert.ToInt32(numInterval.Value);
            Config.Username = txtUser.Text;
            Config.Password = txtPassword.Text;
            Config.Save();
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {
            if (Config.SyncDirection == Config.Direction.dirToOutlook)
            {
                rbGoogleToOutlook.Checked = true;
            }
            else
            {
                rbOutlookToGoogle.Checked = true;
            }
            numInterval.Value = Config.Interval;
            txtUser.Text = Config.Username;
            txtPassword.Text = Config.Password;            
        }
        
        
    }
}
