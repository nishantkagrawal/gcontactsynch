// ##################################################################################
//
// Sample for InboxFolder ItemAdd Event 
// by Helmut Obertanner ( flash [at] x4u dot de
//
// ##################################################################################
namespace GContactsSyncAddin
{
    using System;
    using System.Diagnostics;
    using Microsoft.Office.Core;
    using System.Collections.Generic;
    using System.Windows.Forms;
    using Extensibility;
    using System.Runtime.InteropServices;
    
    // Note: When AddIn was created with Visual Studio Wizard, remove reference to Office.dll and
    // add reference to COM Office11 (Office 2003)
    using Ol = Microsoft.Office.Interop.Outlook;
    using MSO = Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using GContactsSync;
    using GContactsSyncLib;


    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the MyAddin21Setup project 
    // by right clicking the project in the Solution Explorer, then choosing install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("75c7f02f-d056-4d47-ae2e-968fbd804f07"), ProgId("GContactsSync.Connect")]
    public class AddIn : Object, Extensibility.IDTExtensibility2
    {

        /// <summary>
        /// The Outlook Application Object
        /// </summary>
        private Ol.Application myApplicationObject;

        private List<String> ContactsList = new List<string>();
        /// <summary>
        /// My COM AddIn Instance
        /// </summary>
        private object myAddInInstance;

        /// <summary>
        /// My Inbox MAPI Folder Object
        /// </summary>
        private Ol.MAPIFolder ContactsFolder=null;

        private Ol.Items items = null;

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public AddIn()
        {
        }

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            try
            {
                // Get the initial Application object
                myApplicationObject = (Ol.Application)application;
                myAddInInstance = addInInst;
               
                if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
                {
                //    MessageBox.Show("On Connection - Startup");
                    OnStartupComplete(ref custom);
                }

            } catch (System.Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
            }

        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            try
            {
                if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
                {
                    OnBeginShutdown(ref custom);
                }
                if (ContactsFolder != null)
                {
                    // Unregister Events
                   // ContactsFolder.Items.ItemAdd -= new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
                    // Release Reference to COM Object
                    ContactsFolder= null;

                }
                myAddInInstance = null;
                myApplicationObject = null;            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref System.Array custom)
        {
            try
            {
                // Get the ContactsFolder Object
                Config.CheckShowConfig();
                   
                NameSpace nspace = myApplicationObject.GetNamespace("MAPI");
                ContactsFolder = nspace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                
                // Register for It``emAdd Event
                // Note: if more then 16 Items are added, the event doesn't come.
                items = ContactsFolder.Items;
                items.ItemAdd += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
                items.ItemChange += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);
                items.ItemRemove += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemRemoveEventHandler(Items_ItemRemove);
                RefreshContacts();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void RefreshContacts()
        {
            ContactsList.Clear();
            foreach (var item in ContactsFolder.Items)
            {
                ContactsList.Add(((ContactItem)item).EntryID);
            }
        }

        private List<string> RemovedItems()
        {
            List<String> CurrentList = new List<string>();
            List<String> Result = new List<string>();
            foreach (var item in ContactsFolder.Items)
            {
                CurrentList.Add(((ContactItem)item).EntryID);
            }
            foreach (var item in ContactsList)
	        {
                if (CurrentList.IndexOf(item) == -1)
                {
                    Result.Add(item);
                }
            }
            return Result;
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        private void Items_ItemChange(object Item)
        {
            try
            {
                if (Item is ContactItem)
                {
                    ContactItem item = (ContactItem)Item;
                    GoogleAdapter ga = new GoogleAdapter(Config.Username, Config.Password);
                    ga.UpdateContactFromOutlookAsync(item, null);
                }
                
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Items_ItemRemove()
        {
        }

        /// <summary>
        /// EventHandler for ItemAdd Event of InboxFolder
        /// </summary>
        /// <param name="Item">The Item wich was added to InboxFolder</param>
        private void Items_ItemAdd(object Item)
        {
            try
            {
                // Check the ItemType, could be a MeetingAccept or something else
                if (Item is Ol.ContactItem)
                {
                    // Cast to ContactItem Object
                    Ol.ContactItem contact = (Ol.ContactItem)Item;

                    
                    // Do something with Item
                    GoogleAdapter ga = new GoogleAdapter(Config.Username, Config.Password);
                    ga.CreateContactFromOutlookAsync(contact);
                    // Release COM Object
                    contact = null;
                    
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}