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
            Config.Interval = Convert.ToInt32(numInterval.Value);
            Config.Username = txtUser.Text;
            Config.Password = txtPassword.Text;
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {
            numInterval.Value = Config.Interval;
            txtUser.Text = Config.Username;
            txtPassword.Text = Config.Password;
            txtUser.Focus();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            SaveSettings();
        }
        
        
    }
}
