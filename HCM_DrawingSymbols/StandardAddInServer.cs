using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using Microsoft.Win32;

namespace HCMToolsInventorAddIn
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// This will add the buttons and implement for of High Country Millworks Custom Add-ins
    /// </summary>

    [GuidAttribute("5f967b11-035a-4d58-999f-b4460c676acf")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        Inventor.ApplicationEvents appEvents = null;
        //private Inventor.UserInputEvents m_userEvents;
        UserInputEventHandler mUserInputEventHandler;
        ApplicationEventsHandler appEventsHandler = new ApplicationEventsHandler();
        TransactionEventsHandler transEventsHandler;


        //buttons

        //combo-boxes

        //user interface event
        private UserInterfaceEvents userInterfaceEvents;

        //ribbon panels

        //event handler delegates

        
        public StandardAddInServer()
        {
            

        }

        #region ApplicationAddInServer Members
        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            AddinGlobal.InventorApp = addInSiteObject.Application;
            try
            {

                mUserInputEventHandler = new UserInputEventHandler(AddinGlobal.InventorApp.CommandManager.UserInputEvents);
                mUserInputEventHandler.Register();
                appEventsHandler.register();
                transEventsHandler = new TransactionEventsHandler(AddinGlobal.InventorApp.TransactionManager.TransactionEvents);
                transEventsHandler.register();
            }
            catch (Exception e)
            {
                MessageBox.Show("Addin Load Error:" + e.ToString());
            }
        }

       
        public void Deactivate()
        {
            // Release objects.
            //***** button List release****

        }



        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
               
                return null;
            }
        }
        #endregion

    }


}

