using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.Contacts;
using Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using System.Windows.Forms;
using NLog;
using System.Runtime.InteropServices;

namespace GContactsSync
{
    public class GoogleAdapter
    {
        private Logger logger = null;
        public delegate void GContactFetchedHandler(object sender, Contact contact);
        public event GContactFetchedHandler GContactFeched;
        string AppName = typeof(GoogleAdapter).Namespace;
        RequestSettings rs;
        private List<Contact> _Contacts = new List<Contact>();

        public List<Contact> Contacts
        {
            get
            {
                if (_Contacts.Count == 0)
                    EnumerateGoogleContacts();
                return _Contacts;
            }
        }

        private  GoogleAdapter()
        {

        }
        public GoogleAdapter(string UserName,string Password, Logger logger)
        {
            rs = new RequestSettings(AppName,UserName,Password);
            this.logger = logger;
        }

        public string ContactDisplay(Contact c)
        {
            return c.ToString();
        }

        public List<Contact> ContactsChangedSince(DateTime d)
        {
            List<Contact> result = new List<Contact>();
            ContactsRequest cr = new ContactsRequest(rs);
            ContactsQuery q = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            q.StartDate = d;
            Feed<Contact> feed = cr.Get<Contact>(q);
            foreach (Contact c in feed.Entries)
            {
                result.Add(c);
            }
            return result;
        }


        private void LogError(string msg)
        {
            if (logger != null)
            {
                logger.Error(msg);
            }
        }

        private void LogDebug(string msg)
        {
            if (logger != null)
            {
                logger.Debug(msg);
            }
        }

        private void LogInfo(string msg)
        {
            if (logger != null)
            {
                logger.Info(msg);
            }
        }

        public delegate void CreateContactFromOutlookDelegate(ContactItem oContact);

        public void CreateContactFromOutlookAsync(ContactItem oContact)
        {
            CreateContactFromOutlookDelegate cc = new CreateContactFromOutlookDelegate(CreateContactFromOutlook);
            cc.BeginInvoke(oContact, null, null);
        }

        public void CreateContactFromOutlook(ContactItem oContact)
        {
            try
            {
                LogInfo("Creating contact from Outlook");
                Contact gContact = new Contact();
                ContactsRequest cr = new ContactsRequest(rs);
                Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
                UpdateContactDataFromOutlook(oContact, gContact);
                cr.Insert(feedUri, gContact);
                LogInfo("Contact created");
            }
            catch (System.Exception ex)
            {
                LogError(ex.Message);
                LogDebug(ex.StackTrace.ToString());
            }
        }

        public delegate void UpdateContactDelegate (ContactItem oContact, Contact gContact);

        public void UpdateContactFromOutlookAsync (ContactItem oContact, Contact gContact)
        {
            UpdateContactDelegate ud = new UpdateContactDelegate(UpdateContactFromOutlook);
            ud.BeginInvoke(oContact, gContact, null, null);
        }

        [DllImport("User32.dll")]
        public static extern int MessageBox(int h, string m, string c, int type);

        public void UpdateContactFromOutlook(ContactItem oContact, Contact gContact)
        {
            try
            {   
                //MessageBox.Show("In Update");
                LogInfo("Updating contact from Outlook");
                if (gContact == null)
                {
                    LogInfo("Finding contact by FullName");
                    var gContactsQuery = Contacts.Where(c => c.Title == oContact.FullName);
                    if (gContactsQuery.Count() > 0)
                    {
                        gContact = gContactsQuery.First();
                        LogInfo("Contact found: " + gContact.Title);
                    }
                    else
                        throw new System.Exception("Contact not found by FullName: " + oContact.FullName);
                }
                ContactsRequest cr = new ContactsRequest(rs);
                UpdateContactDataFromOutlook(oContact, gContact);
                cr.Update(gContact);
                LogInfo("Contact updated");
            }
            catch (System.Exception ex)
            {
                LogError(ex.Message);
                LogDebug(ex.StackTrace.ToString());
            }
        }

        private void UpdateContactDataFromOutlook(ContactItem oContact, Contact gContact)
        {
            var BusinessAddressQuery = gContact.PostalAddresses.Where(a => a.Work);
            if (BusinessAddressQuery.Count() > 0)
            {
                gContact.PostalAddresses.Remove(BusinessAddressQuery.First());
            }
            if (oContact.BusinessAddress != null)
            {
                PostalAddress BusinessAddress = new PostalAddress();
                BusinessAddress.Rel = ContactsRelationships.IsWork;
                BusinessAddress.Value = oContact.BusinessAddress;
                gContact.PostalAddresses.Add(BusinessAddress);
            }
            
            var HomeAddressQuery = gContact.PostalAddresses.Where(a => a.Home);
            if (HomeAddressQuery.Count() > 0)
            {
                gContact.PostalAddresses.Remove(HomeAddressQuery.First());
            }
            if (oContact.HomeAddress != null)
            {
                PostalAddress HomeAddress = new PostalAddress();
                HomeAddress.Rel = ContactsRelationships.IsHome;
                HomeAddress.Value = oContact.HomeAddress;
                gContact.PostalAddresses.Add(HomeAddress);
            }
            gContact.Emails.Clear();
            if (oContact.Email1Address != null)
            {
                EMail email = new EMail(oContact.Email1Address);
                email.Rel = ContactsRelationships.IsOther;
                gContact.Emails.Add(email);
            }
            if (oContact.Email2Address  != null)
            {
                EMail email = new EMail(oContact.Email2Address);
                email.Rel = ContactsRelationships.IsOther;
                gContact.Emails.Add(email);
            } 
            if (oContact.Email3Address != null)
            {
                EMail email = new EMail(oContact.Email3Address);
                email.Rel = ContactsRelationships.IsOther;
                gContact.Emails.Add(email);
            } 
            gContact.Title = oContact.FullName;


            var BusinessPhoneQuery = gContact.Phonenumbers.Where(a => a.Rel == ContactsRelationships.IsWork);
            if (BusinessPhoneQuery.Count() > 0)
            {
                gContact.Phonenumbers.Remove(BusinessPhoneQuery.First());
            }
            if (oContact.BusinessTelephoneNumber != null)
            {
                PhoneNumber BusinessPhone = new PhoneNumber();
                BusinessPhone.Rel = ContactsRelationships.IsWork;
                BusinessPhone.Value = oContact.BusinessTelephoneNumber;
                gContact.Phonenumbers.Add(BusinessPhone);
            }
            var HomePhoneQuery = gContact.Phonenumbers.Where(a => a.Rel == ContactsRelationships.IsHome);
            if (HomePhoneQuery.Count() > 0)
            {
                gContact.Phonenumbers.Remove(HomePhoneQuery.First());
            }
            if (oContact.HomeTelephoneNumber != null)
            {
                PhoneNumber HomePhone = new PhoneNumber();
                HomePhone.Rel = ContactsRelationships.IsHome;
                HomePhone.Value = oContact.HomeTelephoneNumber;
                gContact.Phonenumbers.Add(HomePhone);
            }
            var MobilePhoneQuery = gContact.Phonenumbers.Where(a => a.Rel == ContactsRelationships.IsMobile);
            if (MobilePhoneQuery.Count() > 0)
            {
                gContact.Phonenumbers.Remove(MobilePhoneQuery.First());
            }
            if (oContact.MobileTelephoneNumber != null)
            {
                PhoneNumber MobilePhone = new PhoneNumber();
                MobilePhone.Rel = ContactsRelationships.IsMobile;
                MobilePhone.Value = oContact.MobileTelephoneNumber;
                gContact.Phonenumbers.Add(MobilePhone);
            }
            var BusinessFaxQuery = gContact.Phonenumbers.Where(a => a.Rel == ContactsRelationships.IsWorkFax);
            if (BusinessFaxQuery.Count() > 0)
            {
                gContact.Phonenumbers.Remove(BusinessFaxQuery.First());
            }
            if (oContact.BusinessFaxNumber != null)
            {
                PhoneNumber BusinessFaxPhone = new PhoneNumber();
                BusinessFaxPhone.Rel = ContactsRelationships.IsWorkFax;
                BusinessFaxPhone.Value = oContact.BusinessFaxNumber;
                gContact.Phonenumbers.Add(BusinessFaxPhone);
            }
            var HomeFaxQuery = gContact.Phonenumbers.Where(a => a.Rel == ContactsRelationships.IsHomeFax);
            if (HomeFaxQuery.Count() > 0)
            {
                gContact.Phonenumbers.Remove(HomeFaxQuery.First());
            }
            if (oContact.HomeFaxNumber != null)
            {
                PhoneNumber HomeFaxPhone = new PhoneNumber();
                HomeFaxPhone.Rel = ContactsRelationships.IsHomeFax;
                HomeFaxPhone.Value = oContact.HomeFaxNumber;
                gContact.Phonenumbers.Add(HomeFaxPhone);
            }
            gContact.Organizations.Clear();
            if (oContact.CompanyName != null)
            {
                Organization org = new Organization();
                org.Name = oContact.CompanyName;
                org.Rel = ContactsRelationships.IsWork;
                if (oContact.JobTitle != null)
                    org.Title = oContact.JobTitle;
                gContact.Organizations.Add(org);
            }
            else
                gContact.Organizations.Clear();
            gContact.Title = oContact.FullName;

        }

        
        public int EnumerateGoogleContacts()
        {
            int SynchedContacts = 0;
            rs.AutoPaging = true;
            ContactsRequest cr = new ContactsRequest(rs);
            Feed<Contact> fc = cr.GetContacts();
            foreach (Contact gContact in fc.Entries)
            {
                _Contacts.Add(gContact);
                if (GContactFeched != null)
                {
                    GContactFeched (this,gContact);
                }
            }
            return SynchedContacts;
        }


    }

}

