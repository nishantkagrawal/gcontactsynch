using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Contacts;
using Google.GData.Extensions;
using System.Threading;
using GContactsSyncLib;
using System.IO;
using Google.GData.Contacts;
using System.Windows.Forms;
using NLog;

namespace GContactsSync
{
    class ContactSync 
    {
        public static object SynchRoot = new object();
        public static Logger logger = Config.GetAppLogger();
        public delegate void OutlookSynchedHandler(object sender, ContactItem contact,  int current, int total);
        public delegate void GoogleSynchedHandler(object sender, Contact contact, int current, int total);
        public delegate void ErrorHandler(object sender, System.Exception ex);
        public delegate void StartSynchingHandler(object sender);
        public event OutlookSynchedHandler OutlookSynched;
        public event GoogleSynchedHandler GoogleSynched;
        public event ErrorHandler Error;
        public event StartSynchingHandler StartSynching;
        public event StartSynchingHandler EndSynching;

        
        GoogleAdapter googleAdapter;
        OutlookAdapter outlookAdapter;
        int Interval;

        public ContactSync(string Username, string Password, int Interval)
        {
            googleAdapter = new GoogleAdapter(Username, Password,null);
            outlookAdapter = new OutlookAdapter();
            this.Interval = Interval;
        }

        
        private void SetLastGoogleFetch(DateTime d)
        {
            StreamWriter sw = new StreamWriter(Config.LocalUserPath()+"\\GoogleFetch.dat");
            sw.Write(d.ToString("yyyy-MM-dd HH:mm:ss"));
            sw.Close();
        }

        private DateTime GetLastGoogleFetch ()
        {
            DateTime result;
            try
            {
                StreamReader sr = new StreamReader(Config.LocalUserPath() + "\\GoogleFetch.dat");
                try
                {
                    result = DateTime.ParseExact(sr.ReadLine(), "yyyy-MM-dd HH:mm:ss", null);
                }
                finally
                {
                    sr.Close();
                }
            } catch
            {
                result = DateTime.Now;
            }
            return result;
        }

        public void SyncFromGoogle ()
        {
            do
            {

                try
                {
                    Monitor.Enter(ContactSync.SynchRoot);
                    try
                    {
                        logger.Info("Starting sync from Google...");
                        DateTime lastFetch = GetLastGoogleFetch();
                        logger.Debug("Last fetch date: " + lastFetch.ToShortDateString());

                        if (StartSynching != null) StartSynching(this);
                        try
                        {
                            logger.Info("Obtaining updated Google contacts");
                            List<Contact> contacts = googleAdapter.ContactsChangedSince(lastFetch);
                            foreach (Contact c in contacts)
                            {
                                if (MutexManager.IsBlocked(c))
                                {
                                    logger.Debug("Removing contact " + GoogleAdapter.ContactDisplay(c) + " from mutex file");
                                    MutexManager.ClearBlockedContact(c);
                                    continue;
                                }
                                logger.Info("Syncronizing " + GoogleAdapter.ContactDisplay(c));
                                logger.Debug("Adding " + GoogleAdapter.ContactDisplay(c) + " to mutex list");
                                MutexManager.AddToBlockedContacts(c);
                                outlookAdapter.UpdateContactFromGoogle(c, null);
                                logger.Info(GoogleAdapter.ContactDisplay(c) + " syncrhonized");
                                if (Thread.CurrentThread.ThreadState == ThreadState.AbortRequested)
                                {
                                    return;
                                }
                            }
                        }
                        finally
                        {
                            if (EndSynching != null) EndSynching(this);
                        }
                        logger.Debug("Setting last fetch date");
                        SetLastGoogleFetch(DateTime.Now);
                        logger.Info("Waiting for " + Interval + " milliseconds.");
                    }
                    finally
                    {
                        Monitor.Exit(ContactSync.SynchRoot);
                    }
                    Thread.Sleep(Interval);

                }
                catch (System.Exception ex)
                {
                    logger.Info(ex.Message);
                    logger.Debug(ex.StackTrace.ToString());
                    if (ex.Message.ToLower().Contains("captcha"))
                    {
                        MessageBox.Show("You need to enter a captcha to validate your account.\r\n" +
                            "GContactsSync will redirect you to the validation page.\r\n" +
                            "Press OK on the next message box when you have finished validating.");
                        System.Diagnostics.Process.Start("https://www.google.com/accounts/DisplayUnlockCaptcha");
                        MessageBox.Show("Press OK after validating your Google Account");
                    }
                    else
                    {
                        MessageBox.Show("Error " + ex.Message);
                    }
                }
                
            } while (Thread.CurrentThread.ThreadState != ThreadState.AbortRequested);

        }

        
                public void FullSyncToOutlook()
        {
            FullSync(Config.Direction.dirToOutlook);
        }
        
        public void FullSyncToGoogle()
        {
            FullSync(Config.Direction.dirToGoogle);
        }

        public void FullSync(Config.Direction dir)
        {
            MutexManager.StartingFullSync();
            try
            {
                try
                {
                    if (dir == Config.Direction.dirToOutlook)
                    {
                        SetLastGoogleFetch(DateTime.Now);
                    }
                    if (StartSynching != null) StartSynching(this);
                    List<ContactItem> OutlookContacts;
                    List<Contact> GoogleContacts;

                    logger.Info("Obtaining Outlook contacts...");
                    OutlookContacts = outlookAdapter.Contacts;
                    logger.Info("Obtaining Google Contacts...");
                    GoogleContacts = googleAdapter.Contacts;

                    EmailComparer comparer = new EmailComparer();
                    int i = 0;
                    int total = OutlookContacts.Count() + GoogleContacts.Count();
                    foreach (ContactItem item in OutlookContacts)
                    {
                        try
                        {
                            try
                            {
                                if (Thread.CurrentThread.ThreadState == ThreadState.AbortRequested)
                                    break;
                                i++;
                                var qryFindGoogleContact = GoogleContacts.Where(c => c.Title == item.FullName ||
                                    c.Emails.Contains(new EMail(item.Email1Address), comparer) ||
                                    c.Emails.Contains(new EMail(item.Email2Address), comparer) ||
                                    c.Emails.Contains(new EMail(item.Email3Address), comparer));
                                if (qryFindGoogleContact.Count() > 0)
                                {
                                    if (dir == Config.Direction.dirToOutlook)
                                    {
                                        logger.Info(String.Format("{0}/{1} Updating Outlook contact: " + OutlookAdapter.ContactItemDisplay(item), i, total));
                                        Contact gContact = qryFindGoogleContact.First();
                                        logger.Info("Updating outlook contact");
                                        SyncContact(gContact, item, dir);
                                    }
                                    else
                                        logger.Info(String.Format("{0}/{1} Processing...", i, total));
                                }
                                else
                                {
                                    logger.Info(String.Format("{0}/{1} Creating Google contact for " + OutlookAdapter.ContactItemDisplay(item), i, total));
                                  //  MutexManager.AddToBlockedContacts(item);
                                    CreateGoogleContact(item);
                                }
                            }
                            finally
                            {
                                if (OutlookSynched != null) OutlookSynched(this, item, i, OutlookContacts.Count());
                            }
                        }
                        catch (System.Exception ex)
                        {
                            if (ex is ThreadAbortException)
                                break;
                        }
                    }

                    foreach (Contact item in GoogleContacts)
                    {
                        try
                        {

                            try
                            {
                                if (Thread.CurrentThread.ThreadState == ThreadState.AbortRequested)
                                    break;
                                i++;
                                var qryFindOutlookContact = OutlookContacts.Where(c => c.FullName == item.Title ||
                                    item.Emails.Contains(new EMail(c.Email1Address), comparer) ||
                                    item.Emails.Contains(new EMail(c.Email2Address), comparer) ||
                                    item.Emails.Contains(new EMail(c.Email3Address), comparer));
                                if (qryFindOutlookContact.Count() > 0)
                                {
                                    if (dir == Config.Direction.dirToGoogle)
                                    {
                                        logger.Info(String.Format("{0}/{1} Updating Google contact: " + GoogleAdapter.ContactDisplay(item), i, total));
                                        ContactItem oContact = qryFindOutlookContact.First();
                                        SyncContact(item, oContact, dir);
                                    }
                                    else
                                        logger.Info(String.Format("{0}/{1} Processing...", i, total));
                                }
                                else
                                {
                                    logger.Info(String.Format("{0}/{1} Creating Outlook contact for " + GoogleAdapter.ContactDisplay(item), i, total));
                                    CreateOutlookContact(item);
                                }
                            }
                            finally
                            {
                                if (GoogleSynched != null) GoogleSynched(this, item, i, GoogleContacts.Count());
                            }
                        }
                        catch (System.Exception ex)
                        {
                            if (ex is ThreadAbortException)
                                break;
                        }

                    }
                    if (EndSynching != null) EndSynching(this);

                }
                catch (System.Exception ex)
                {
                    if ((EndSynching != null))
                        EndSynching(this);
                    else if (!(ex is ThreadAbortException) & (Error != null))
                        Error(this, ex);

                }
            }
            finally
            {
                if (dir == Config.Direction.dirToGoogle)
                {
                    SetLastGoogleFetch(DateTime.Now);
                }
                MutexManager.EndingFullSync();
            }
        }

        private void CreateOutlookContact(Contact item)
        {
            outlookAdapter.CreateContactFromGoogle(item);
        }

        private void CreateGoogleContact(ContactItem item)
        {
            googleAdapter.CreateContactFromOutlook(item);
        }

        private void SyncContact(Contact gContact, ContactItem item, Config.Direction dir)
        {
            if (dir == Config.Direction.dirToGoogle)
            {
                googleAdapter.UpdateContactFromOutlook(item, gContact);
            }
            else
            {
                outlookAdapter.UpdateContactFromGoogle(gContact, item);
            }
        }
    }
}
