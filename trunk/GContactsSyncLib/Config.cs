using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GContactsSyncLib.Properties;

namespace GContactsSyncLib
{
    public class Config
    {
        public enum Direction {dirToGoogle,dirToOutlook};

        public static void ShowConfig()
        {
            ConfigForm frm = new ConfigForm();
            frm.ShowDialog();
        }

        public static void CheckShowConfig()
        {
            if (Config.Username == "")
                ShowConfig();
        }

        public static string Username
        {
            get { return Settings.Default.User; }
            set { Settings.Default.User = value; }
        }

        public static string Password
        {
            get { return Settings.Default.Password !="" ? EncryptDecrypt.Decrypt(Settings.Default.Password) : ""; }
            set { Settings.Default.Password= value == "" ? "" : EncryptDecrypt.Encrypt(value); }
        }

        public static Direction SyncDirection
        {
            get { return (Direction) Enum.ToObject(typeof(Direction), Settings.Default.Direction); }
            set { Settings.Default.Direction = (int)value; }
        }

        public static int Interval
        {
            get { return Settings.Default.Interval; }
            set { Settings.Default.Interval = value; }
        }

        public static void Save()
        {
            Settings.Default.Save();
        }
    }
}
