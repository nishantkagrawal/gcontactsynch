using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Google.Contacts;
using Microsoft.Office.Interop.Outlook;

namespace GContactsSyncLib
{
    public class MutexManager
    {
        static string OutlookMutexFile = Config.LocalUserPath() + "\\BlockedOutlookContacts.dat";
        static string GoogleMutexFile = Config.LocalUserPath() + "\\BlockedGoogleContacts.dat";

        public static bool IsBlocked(ContactItem item)
        {
            if (!File.Exists(OutlookMutexFile))
                return false;
            StreamReader sr = new StreamReader(OutlookMutexFile);
            try
            {
                while (!sr.EndOfStream)
                {
                    List<string> data = new List<string>(sr.ReadLine().Split(','));
                    if (data[0] == item.FullName)
                    {
                        return true;
                    }
                    data.RemoveAt(0);
                    if ((data.Contains(item.Email1Address)) ||
                        (data.Contains(item.Email2Address)) ||
                        (data.Contains(item.Email3Address)))
                        return true;
                }
            }
            finally
            {
                sr.Close();
            }
            return false;

        }

        public static bool IsBlocked(Contact item)
        {
            if (!File.Exists(GoogleMutexFile))
                return false;
            StreamReader sr = new StreamReader(GoogleMutexFile);
            try
            {
                while (!sr.EndOfStream)
                {
                    List<string> data = new List<string>(sr.ReadLine().ToLower().Split(','));
                    if (data[0] == item.Title.ToLower())
                    {
                        return true;
                    }
                    data.RemoveAt(0);
                    if (data.Intersect(item.Emails.Select(e=>e.Address.ToLower()).ToList()).Count()>0)
                        return true;
                }
            }
            finally
            {
                sr.Close();
            }
            return false;

        }

        public static void AddToBlockedContacts(Contact c)
        {
            bool added = false;
            int retries = 0;
            do
            {
                try
                {
                    StreamWriter sw = new StreamWriter(OutlookMutexFile);
                    try
                    {
                        sw.WriteLine(c.Title + "," + string.Join(",", c.Emails.Select(e => e.Address).ToArray()));
                    }
                    finally
                    {
                        sw.Close();
                    }
                    added = true;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(500);
                    retries++;
                }
            } while (!added && retries <= 3);
        }

        public static void AddToBlockedContacts(ContactItem c)
        {
            bool added = false;
            int retries = 0;
            do
            {
                try
                {
                    StreamWriter sw = new StreamWriter(GoogleMutexFile,true);
                    try
                    {
                        sw.WriteLine(c.FullName+ "," + c.Email1Address+","+c.Email2Address+","+c.Email3Address);
                    }
                    finally
                    {
                        sw.Close();
                    }
                    added = true;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(500);
                    retries++;
                }
            } while (!added && retries <= 3);
        }



        public static void ClearBlockedContact(ContactItem item)
        {
            int retries = 0;
            bool saved = false;
            bool found = false;
            List<string> lines = new List<string>();
            List<string> linesToSave;
            string FilePath = OutlookMutexFile;

            if (!File.Exists(FilePath))
                return;
            StreamReader sr = new StreamReader(FilePath);
            while (!sr.EndOfStream)
            {
                lines.Add(sr.ReadLine());
            }
            sr.Close();
            linesToSave = new List<string>(lines);
            foreach (string line in lines)
            {
                List<string> data = new List<string>(line.Split(','));
                if (data[0] == item.FullName)
                {
                    found = true;
                }
                else
                {
                    data.RemoveAt(0);
                    if ((data.Contains(item.Email1Address)) ||
                        (data.Contains(item.Email2Address)) ||
                        (data.Contains(item.Email3Address)))
                        found = true;
                }
                if (found)
                {
                    linesToSave.Remove(line);
                    //don't break... may have multiple instances
                    //break;
                }
            }
            if (found)
            {
                do
                {
                    try
                    {
                        StreamWriter sw = new StreamWriter(FilePath);
                        try
                        {
                            foreach (string line in linesToSave)
                            {
                                sw.WriteLine(line);
                            }
                        }
                        finally
                        {
                            sw.Close();
                        }
                        saved = true;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(500);
                        retries++;
                    }
                }
                while (!saved && retries <= 3);
            }
        }
            
            
    

    public static void ClearBlockedContact(Contact item)
        {
            int retries = 0;
            bool saved = false;
            bool found = false;
            List<string> lines = new List<string>();
            string FilePath = GoogleMutexFile;

            if (!File.Exists(FilePath))
                return;
            StreamReader sr = new StreamReader(FilePath);
            while (!sr.EndOfStream)
            {
                lines.Add(sr.ReadLine());
            }
            sr.Close();
            foreach (string line in lines)
            {
                List<string> data = new List<string>(line.ToLower().Split(','));
                if (data[0] == item.Title.ToLower())
                {
                    found = true;
                }
                else
                {
                    data.RemoveAt(0);
                    if (data.Intersect(item.Emails.Select(e=>e.Address.ToLower()).ToList()).Count()>0)
                        found = true;
                }
                if (found)
                {
                    lines.Remove(line);
                    break;
                }
            }
            if (found)
            {
                do
                {
                    try
                    {
                        StreamWriter sw = new StreamWriter(FilePath);
                        try
                        {
                            foreach (string line in lines)
                            {
                                sw.WriteLine(line);
                            }
                        }
                        finally
                        {
                            sw.Close();
                        }
                        saved = true;
                    }
                    catch (IOException)
                    {
                        System.Threading.Thread.Sleep(500);
                        retries++;
                    }
                }
                while (!saved && retries <= 3);
            }
        }
    }

}
