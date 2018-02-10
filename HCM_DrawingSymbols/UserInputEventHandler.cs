using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;
using System.Windows.Forms;

namespace HCMToolsInventorAddIn
{
    public class UserInputEventHandler
    {
        private UserInputEvents mEventsObj;
        public UserInputEventHandler(UserInputEvents obj)
        {
            mEventsObj = obj;
        }

        public void Register()
        {
            mEventsObj.OnActivateCommand += UserInputEvents_OnActivateCommand_Handler;
        }

        private void UserInputEvents_OnActivateCommand_Handler(string CommandName, NameValueMap Context)
        {
            Inventor.Application app = AddinGlobal.InventorApp;
            Document invDoc = app.ActiveDocument;
            if (CommandName == "AppFilePrintCmd" && invDoc is DrawingDocument)
            {
                updatePlotStamp(invDoc);
            }
            if (CommandName == "AppFilePublishCmd" && invDoc is DrawingDocument)
            {
                updatePlotStamp(invDoc);
            }
            if (CommandName == "VaultCheckinTop" && invDoc is DrawingDocument)
            {
                updatePlotStamp(invDoc);
            }
            if (CommandName == "VaultCheckin" && invDoc is DrawingDocument)
            {
                updatePlotStamp(invDoc);
            }
            
        }

        public void Unregister()
        {
            mEventsObj.OnActivateCommand -= UserInputEvents_OnActivateCommand_Handler;
        }
        private void updateDrawingStatus(Document doc)
        {

        }

        private void updatePlotStamp(Document doc)
        {
            DateTime saveDate = System.IO.File.GetLastWriteTime(doc.FullFileName);
            string modDate = saveDate.ToString();
            string userInitials = (System.Environment.UserName.Substring(4, 3)).ToUpper();
            string plotStamp = ("Date: " + DateTime.Now.ToString() + " By: " + userInitials + " Last Modified: " +
                modDate);
            doc.updateIProperty("Plot Stamp", plotStamp);
            DrawingDocument drawDoc = (DrawingDocument)doc;
            if ((int)(doc.PropertySets["Design Tracking Properties"]["Design Status"].Value) == 3)
            {
                doc.updateIProperty("Drawing Status", "APPROVED FOR FABRICATION");
            }
            else
            {
                doc.updateIProperty("Drawing Status", "PRELIMINARY **DO NOT USE FOR FABRICATION**");
            }
            foreach (Sheet oSheet in drawDoc.Sheets)
            {
                oSheet.Update();
            }
        }
    }
}
