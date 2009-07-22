using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NET.Commons;
using NLog;
using NLog.Targets;

namespace GContactsSyncLib
{
    public class Config
    {
        public enum Direction {dirToGoogle,dirToOutlook};

        private static PropertyFile Settings;
        static Config()
        {
            String confPath = LocalUserPath()+"\\Config.dat";
            Settings = new PropertyFile(confPath);
            if (!Settings.Exists())
            {
                Settings.Put("Interval", "3");
                Settings.Put("Username", "");
                Settings.Put("Password", "");
            }
        }

        private static Logger GetLoggerWithFileName(string FileName)
        {
            FileTarget target = new FileTarget();
            target.Layout = "${longdate} {$processname} ${level} - ${message}";
            target.FileName = LocalUserPath() + "/logs/"+FileName+".log";
            target.ArchiveFileName = LocalUserPath() + "/logs/archives/"+FileName+".{#####}.log";
            target.ArchiveNumbering = FileTarget.ArchiveNumberingMode.Sequence;
            target.ArchiveAboveSize = 10240000;
            target.ConcurrentWrites = true;
            target.KeepFileOpen = false;
            target.Encoding = "iso-8859-2";
            NLog.Config.SimpleConfigurator.ConfigureForTargetLogging(target, LogLevel.Debug);
            return LogManager.GetCurrentClassLogger();
        }

        public static Logger GetAppLogger()
        {
            return GetLoggerWithFileName("GContactsSyncApp");
        }

        public static Logger GetAddinLogger()
        {
            return GetLoggerWithFileName("GContactsSyncAddin");
        }

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

        public static string LocalUserPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\GContactsSync";
        }

        public static string Username
        {
            get { return Settings.Get("Username"); }
            set { Settings.Put("Username", value) ; }
        }

        public static string Password
        {
            get 
            {
                string pwd = Settings.Get("Password");
                return pwd !="" ? EncryptDecrypt.Decrypt(pwd) : ""; 
            }
            set 
            { 
                if (value != "")
                    Settings.Put("Password",EncryptDecrypt.Encrypt(value)); 
            }
        }

        public static int Interval
        {
            get { return int.Parse(Settings.Get("Interval")); }
            set { Settings.Put("Interval", value.ToString()); }
        }

    }
}
