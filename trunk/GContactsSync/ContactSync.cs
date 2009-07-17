using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Contacts;
using Google.GData.Extensions;
using System.Threading;

namespace GContactsSync
{
    class ContactSync 
    {
        public delegate void OutlookSynchedHandler(object sender, ContactItem contact,  int current, int total);
        public delegate void GoogleSynchedHandler(object sender, Contact contact, int current, int total);
        public delegate void ErrorHandler(object sender, System.Exception ex);
        public delegate void StartSyncingHandler(object sender);
        public event OutlookSynchedHandler OutlookSynched;
        public event GoogleSynchedHandler GoogleSynched;
        public event ErrorHandler Error;
        public event StartSyncingHandler StartSyncing;
        public event StartSyncingHandler EndSyncing;

        public enum Direction {dirToGoogle,dirToOutlook};
        GoogleEnumerator gEnum;
        OutlookEnumerator oEnum;
        int Interval;

        public ContactSync(string Username, string Password, int Interval)
        {
            gEnum = new GoogleEnumerator(Username, Password);
            oEnum = new OutlookEnumerator();
            this.Interval = Interval;
        }

        public void SyncToOutlook()
        {
            Sync(Direction.dirToOutlook);
        }
        
        public void SyncToGoogle()
        {
            Sync(Direction.dirToGoogle);
        }

        public void Sync(Direction dir)
        {
            try
            {
                do
                {
                    if (StartSyncing != null) StartSyncing(this);
                    List<ContactItem> OutlookContacts;
                    List<Contact> GoogleContacts;

                    OutlookContacts = oEnum.Contacts;
                    GoogleContacts = gEnum.Contacts;

                    //if (dir == Direction.dirToGoogle)
                    //{
                    //    GoogleContacts = GoogleContacts.Where(c =>
                    //        !OutlookContacts.Select(o => o.FullName).Contains(c.Title) &
                    //        OutlookContacts.Select(o => o.Email1Address).Intersect(c.Emails.Select(e => e.Address)) == null &
                    //        OutlookContacts.Select(o => o.Email2Address).Intersect(c.Emails.Select(e => e.Address)) == null &
                    //        OutlookContacts.Select(o => o.Email3Address).Intersect(c.Emails.Select(e => e.Address)) == null).ToList();
                    //}
                    //else
                    //{
                    //    OutlookContacts = OutlookContacts.Where(o =>
                    //        !GoogleContacts.Select(c => c.Title).Contains(o.FullName) &
                    //        GoogleContacts.Where(c => c.Emails.Select(e => e.Address).Contains(o.Email1Address)).Count() == 0 &
                    //        GoogleContacts.Where(c => c.Emails.Select(e => e.Address).Contains(o.Email2Address)).Count() == 0 &
                    //        GoogleContacts.Where(c => c.Emails.Select(e => e.Address).Contains(o.Email3Address)).Count() == 0).ToList();
                    //}

                    EmailComparer comparer = new EmailComparer();
                    int i = 0;
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
                                    if (dir == Direction.dirToOutlook)
                                    {
                                        Contact gContact = qryFindGoogleContact.First();
                                        SyncContact(gContact, item, dir);
                                    }
                                }
                                else
                                {
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
                        }
                    }
                    i = 0;
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
                                    if (dir == Direction.dirToGoogle)
                                    {
                                        ContactItem oContact = qryFindOutlookContact.First();
                                        SyncContact(item, oContact, dir);
                                    }
                                }
                                else
                                {
                                    CreateOutlookContact(item);
                                }
                            }
                            finally
                            {
                                if (GoogleSynched != null) GoogleSynched(this, item, i, GoogleContacts.Count());
                            }
                        }
                        catch (System.Exception)
                        {
                        }

                    }
                    if (EndSyncing != null) EndSyncing(this);
                    Thread.Sleep(Interval);
                } while (Thread.CurrentThread.ThreadState != ThreadState.AbortRequested);
            }
            catch (System.Exception ex)
            {
                if (Error != null) Error(this, ex);
            }

        }

        private void CreateOutlookContact(Contact item)
        {
            oEnum.CreateContactFromGoogle(item);
        }

        private void CreateGoogleContact(ContactItem item)
        {
            gEnum.CreateContactFromOutlook(item);
        }

        private void SyncContact(Contact gContact, ContactItem item, Direction dir)
        {
            if (dir == Direction.dirToGoogle)
            {
                gEnum.UpdateContactFromOutlook(item, gContact);
            }
            else
            {
                oEnum.UpdateContactFromGoogle(gContact, item);
            }
        }
    }
}
