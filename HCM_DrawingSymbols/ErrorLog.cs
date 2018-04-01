using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Inventor;
using System.Windows.Forms;
using System.Reflection;

namespace HCMToolsInventorAddIn
{

    class ErrorLog
    {
        public static void LogError(string errorMessage, Document activeDoc, string transName)
        {
            try
            {
                string userInitials = (System.Environment.UserName.Substring(4, 3)).ToUpper();
                string dateTimeString = DateTime.Now.ToString();
                dateTimeString = dateTimeString.Replace("/", "_");
                dateTimeString = dateTimeString.Replace(":", ".");
                string logFileName = userInitials + "_" + dateTimeString + "_Error Log.txt";
                string logPath = @"C:\Users\Public\Documents\HCM Addin Logs\"; //@"\\HCMS01\Public\AUTOCAD DRAFTING\Inventor\041714\Error Logs\";
                var fullPath = System.IO.Path.Combine(logPath, logFileName);
                StreamWriter theLog;
                if (!System.IO.File.Exists(fullPath))
                {
                    (new FileInfo(fullPath)).Directory.Create();
                    theLog = new StreamWriter(fullPath, true);
                    theLog.WriteLine(DateTime.Now + "\n" + userInitials + " Error Log\n");
                    theLog.WriteLine(activeDoc.FullDocumentName + " Transaction Name: " + transName + "\n");
                    theLog.WriteLine(errorMessage);
                    theLog.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Writing to Log File\n" + ex.Message + ex.StackTrace);
            }
        }

        public static void LogError(Exception error, Document activeDoc, string transName)
        {
            var message = error.Message + "\n" + error.StackTrace;
            LogError(message, activeDoc, transName);
        }
    }
}
