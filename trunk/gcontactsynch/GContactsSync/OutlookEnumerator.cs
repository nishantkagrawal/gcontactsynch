using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Contacts;
using Google.GData.Extensions;


namespace GContactsSync
{
    class OutlookEnumerator
    {
        Application outlook;
        public delegate void OContactFetchedHandler(object sender, ContactItem contact);
        public event OContactFetchedHandler OContactFeched;
        private List<ContactItem> _Contacts = new List<ContactItem>();
        NameSpace nspace;
        MAPIFolder contactsFolder;
            

        public void CreateContactFromGoogle (Contact gContact)
        {
            ContactItem oContact = (ContactItem)outlook.CreateItem(OlItemType.olContactItem);
            UpdateContactDataFromGoogle(gContact, oContact);
        }

        public void UpdateContactFromGoogle (Contact gContact, ContactItem oContact)
        {
            UpdateContactDataFromGoogle(gContact, oContact);
        }

        private void UpdateContactDataFromGoogle(Contact gContact, ContactItem oContact)
        {
            if (gContact.Organizations.Count > 0)
            {
                oContact.CompanyName = gContact.Organizations[0].Name;
                oContact.JobTitle = gContact.Organizations[0].Title;
            }
           
            //oContact.WebPage = gContact.weex
            var qryBusinessAddress = gContact.PostalAddresses.Where(p => !p.Home);
            if (qryBusinessAddress.Count() > 0)
            {
                oContact.BusinessAddress = qryBusinessAddress.First().Value;
            }
            else
                oContact.BusinessAddress = "";
            var qryHomeAddress = gContact.PostalAddresses.Where(p => p.Home);
            if (qryHomeAddress.Count() > 0)
            {
                oContact.HomeAddress = qryHomeAddress.First().Value;
            }
            else
                oContact.HomeAddress = "";
            try
            {
                oContact.Email1Address = gContact.Emails[0].Address;
                oContact.Email1DisplayName = gContact.Title;
            }
            catch
            {
                oContact.Email1Address = "";
                oContact.Email1DisplayName = "";
            }
            try
            {
                oContact.Email2Address = gContact.Emails[1].Address;
                oContact.Email2DisplayName = gContact.Title;
            }
            catch
            {
                oContact.Email2Address = "";
                oContact.Email2DisplayName = "";
            }
            try
            {
                oContact.Email3Address = gContact.Emails[2].Address;
                oContact.Email3DisplayName = gContact.Title;
            }
            catch
            {
                oContact.Email3Address = "";
                oContact.Email3DisplayName = "";
            }
            if (gContact.PrimaryEmail != null)
            {
                if (gContact.Title != gContact.PrimaryEmail.Address)
                {
                    oContact.FullName = gContact.Title;
                }
            }
            else
            {
                oContact.FullName = gContact.Title;
            }
            var qryBusinessPhone = gContact.Phonenumbers.Where(p => p.Work);
            if (qryBusinessPhone.Count() > 0)
            {
                oContact.BusinessTelephoneNumber = qryBusinessPhone.First().Value;
            }
            else
                oContact.BusinessTelephoneNumber = "";
            var qryHomePhone = gContact.Phonenumbers.Where(p => p.Home);
            if (qryHomePhone.Count() > 0)
            {
                oContact.HomeTelephoneNumber = qryHomePhone.First().Value;
            }
            else
                oContact.HomeTelephoneNumber = "";
            var qryWorkFax = gContact.Phonenumbers.Where(p => p.Rel == ContactsRelationships.IsWorkFax);
            if (qryWorkFax.Count() > 0)
            {
                oContact.BusinessFaxNumber = qryWorkFax.First().Value;
            }
            else
                oContact.BusinessFaxNumber = "";
            var qryHomeFax = gContact.Phonenumbers.Where(p => p.Rel == ContactsRelationships.IsHomeFax);
            if (qryHomeFax.Count() > 0)
            {
                oContact.HomeFaxNumber = qryHomeFax.First().Value;
            }
            else
                oContact.HomeFaxNumber = "";
            var qryMobile = gContact.Phonenumbers.Where(p => p.Rel == ContactsRelationships.IsMobile);
            if (qryMobile.Count() > 0)
            {
                oContact.MobileTelephoneNumber = qryMobile.First().Value;
            }
            else
                oContact.MobileTelephoneNumber = "";
            oContact.Save();
        }

        public List<ContactItem> Contacts
        {
            get
            {
                if (_Contacts.Count == 0)
                    EnumerateOutlookContacts();
                return _Contacts;
            }
        }
        public OutlookEnumerator()
        {
            outlook = new Application();
            nspace = outlook.GetNamespace("MAPI");
            contactsFolder = nspace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
        }

        public int EnumerateOutlookContacts()
        {
            foreach (var item in contactsFolder.Items)
            {
                ContactItem contact = (ContactItem)item;
                _Contacts.Add(contact);
                if (OContactFeched != null)
                    OContactFeched(this, contact);
            }
            return 0;
        }
    }


}
