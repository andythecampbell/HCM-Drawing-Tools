using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HCMToolsInventorAddIn
{
    public class ActionResult
    {
        public ActionResult(bool success)
        {
            Success = success;
        }
        public ActionResult(bool success, string message)
        {
            Success = success;
            Message = message;
        }

        public ActionResult(Exception exception)
        {
            Success = false;
            Exception = exception;
            Message = exception.Message;
        }

        public bool Success { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
    }
}
