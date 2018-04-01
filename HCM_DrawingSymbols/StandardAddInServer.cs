using System;
using System.Drawing;
using System.Linq;
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
        ApplicationEventsHandler appEventsHandler;
        TransactionEventsHandler transEventsHandler;


        //buttons
        private SyncSymbolsButton syncSymbolsButton;

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
                appEventsHandler = new ApplicationEventsHandler(AddinGlobal.InventorApp.ApplicationEvents);
                appEventsHandler.Register();
                transEventsHandler = new TransactionEventsHandler(AddinGlobal.InventorApp.TransactionManager.TransactionEvents);
                transEventsHandler.Register();
                
                GuidAttribute addInCLSID;
                addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(StandardAddInServer), typeof(GuidAttribute));
                string addInCLSIDString;
                addInCLSIDString = "{" + addInCLSID.Value + "}";
                
                //create buttons
                syncSymbolsButton = new SyncSymbolsButton("Sync Sketch Symbols", "Autodesk:HCMDrawingSymbols:SyncSymbolsButton",
                    CommandTypesEnum.kShapeEditCmdType, addInCLSIDString, "Syncs Custom Sketch Symbols With Document Data",
                    "Sync Drawing Symbols", ButtonDisplayEnum.kDisplayTextInLearningMode);

                if (firstTime == true)
                {
                    UserInterfaceManager uiMan = AddinGlobal.InventorApp.UserInterfaceManager;
                    if (uiMan.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
                    {

                        //Ribbon pRibbon = uiMan.Ribbons["Part"];
                        //Ribbon aRibbon = uiMan.Ribbons["Assembly"];
                        
                        Ribbon dRibbon = uiMan.Ribbons["Drawing"];
                        //RibbonTab dHCMTab = dRibbon.RibbonTabs.Add("HCM Tools", "HCMCustomTools", addInCLSIDString, "", false, false);
                        
                        RibbonTab drawingAnnotateTab = dRibbon.RibbonTabs[2];
                        AddinGlobal.RibbonPanelId = "{d04e0c45-dec6-4881-bd3f-a7a81b99f307}";
                        //CommandControls cmdCtrlsDataPart = AddinGlobal.RibbonPanel.CommandControls;
                        AddinGlobal.RibbonPanelId = "{d04e0c45-dec7-4881-bd3f-a7a81b99f307}";
                        //CommandControls cmdCtrlsAssembly = AddinGlobal.RibbonPanel.CommandControls;
                        AddinGlobal.RibbonPanel = drawingAnnotateTab.RibbonPanels[1];//.Add("DrawingTools", "InventorAddinServer.RibbonPanel_" + Guid.NewGuid().ToString(),
                            //AddinGlobal.RibbonPanelId, String.Empty, false);
                        CommandControls cmdCtlsDrawing = AddinGlobal.RibbonPanel.CommandControls;

                        cmdCtlsDrawing.AddButton(syncSymbolsButton.ButtonDefinition, true, true, "", false);
                        
                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Addin Load Error:" + e.ToString() + e.StackTrace);
            }
        }

       
        public void Deactivate()
        {
            if (syncSymbolsButton.ButtonDefinition != null) Marshal.ReleaseComObject(syncSymbolsButton.ButtonDefinition);
            mUserInputEventHandler.Unregister();
            appEventsHandler.Unregister();
            transEventsHandler.Unregister();
          
            //Marshal.ReleaseComObject(m_inventorApplication);
            //m_inventorApplication = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();

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

