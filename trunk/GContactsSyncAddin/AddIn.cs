// ##################################################################################
// Addin Skeleton from:
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
//    using System.Windows.Forms;
    using Extensibility;
    using System.Runtime.InteropServices;
    using Ol = Microsoft.Office.Interop.Outlook;
    using MSO = Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using GContactsSync;
    using GContactsSyncLib;
    using NLog;
    using NLog.Config;
    using NLog.Targets;
    using System.Windows.Forms;
    using System.IO;


    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("75c7f02f-d056-4d47-ae2e-968fbd804f07"), ProgId("GContactsSync.Connect")]
    public class AddIn : Object, Extensibility.IDTExtensibility2
    {
        private static Logger logger = Config.GetAddinLogger(); 
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
                logger.Info("GContactsSync Connecting.");
                myApplicationObject = (Ol.Application)application;
                myAddInInstance = addInInst;
               
                if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
                {
                //    MessageBox.Show("On Connection - Startup");
                    OnStartupComplete(ref custom);
                }
                logger.Info("GContactsSync connected succesfully.");
            } catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace.ToString());
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
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace.ToString());
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
                logger.Debug("Setting ItemAdd Event");
                items.ItemAdd += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
                logger.Debug("Setting ItemChange Event");
                items.ItemChange += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);
                logger.Debug("Setting ItemRemove Event");
                items.ItemRemove += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemRemoveEventHandler(Items_ItemRemove);
                RefreshContacts();
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace.ToString());
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

        private string ContactItemDisplay(ContactItem item)
        {
            return item.FullName.Length == 0 ? item.Email1Address : item.FullName;
        }

        
        private void Items_ItemChange(object Item)
        {
            try
            {      
                if (Item is ContactItem)
                {
                    ContactItem item = (ContactItem)Item;
                    if (MutexManager.IsBlocked(item))
                    {
                        logger.Debug("Removing contact "+ContactItemDisplay(item)+ " from mutex file");
                        MutexManager.ClearBlockedContact(item);
                        return;
                    }
                    logger.Info("Adding " + ContactItemDisplay(item) + " to mutex file");
                    MutexManager.AddToBlockedContacts(item);
                    logger.Info("Updating " + ContactItemDisplay(item));
                    GoogleAdapter ga = new GoogleAdapter(Config.Username, Config.Password,logger);
                    ga.UpdateContactFromOutlookAsync(item, null);
                }
                
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace.ToString());
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

                    logger.Info("Adding " + ContactItemDisplay(contact));        
                    // Do something with Item
                    GoogleAdapter ga = new GoogleAdapter(Config.Username, Config.Password,logger);
                    ga.CreateContactFromOutlookAsync(contact);
                    // Release COM Object
                    contact = null;
                    
                }
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace.ToString());
                MessageBox.Show(ex.Message);
            }
        }
    }
}