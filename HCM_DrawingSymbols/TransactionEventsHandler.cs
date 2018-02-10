using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;
using System.Windows.Forms;

namespace HCMToolsInventorAddIn
{
    class TransactionEventsHandler
    {
        private TransactionEvents mEventsObj;
        public TransactionEventsHandler(TransactionEvents obj)
        {
            mEventsObj = obj;
        }
        public void register()
        {
            
            //mEventsObj.OnUndo += mEventsObj_OnUndo;
            //mEventsObj.OnRedo += mEventsObj_OnRedo;
            //mEventsObj.OnCommit += mEventsObj_OnCommit;
            //mEventsObj.OnDelete += mEventsObj_OnDelete;
            //mEventsObj.OnAbort += mEventsObj_OnAbort;
            
        }

        void mEventsObj_OnRedo(Transaction TransactionObject, NameValueMap Context, EventTimingEnum BeforeOrAfter,
            out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            throw new NotImplementedException();
        }

        void mEventsObj_OnAbort(Transaction TransactionObject, NameValueMap Context, EventTimingEnum BeforeOrAfter)
        {
               throw new NotImplementedException();
        }

        void mEventsObj_OnDelete(Transaction TransactionObject, NameValueMap Context, EventTimingEnum BeforeOrAfter)
        {
               //throw new NotImplementedException();
        }

        void mEventsObj_OnUndo(Transaction TransactionObject, NameValueMap Context, EventTimingEnum BeforeOrAfter,
            out HandlingCodeEnum HandlingCode)
        {
            throw new NotImplementedException();
        }

        void mEventsObj_OnCommit(Transaction TransactionObject, NameValueMap Context, EventTimingEnum BeforeOrAfter,
            out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (BeforeOrAfter == EventTimingEnum.kBefore)
            {
                if (TransactionObject.DisplayName == "Create Callout Symbol")
                {
                    //MessageBox.Show("On Commit " + TransactionObject.Id);
                    Document doc = TransactionObject.Document;
                    DrawingDocument drawingDoc = (DrawingDocument)TransactionObject.Document;
                    SketchedSymbol currentSymbol = drawingDoc.ActiveSheet.SketchedSymbols[drawingDoc.ActiveSheet.SketchedSymbols.Count];
                    
                    if (currentSymbol.Name == "ANN-FINISH")
                    {
                        TransactionManager transMan = TransactionObject.Parent;

                        try
                        {
                          
                        }
                        catch(Exception ex)
                        {
                            //ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, "Trans Event Handler " + TransactionObject.DisplayName);
                            MessageBox.Show("transaction error: " + ex.Message);
                        }
                        try
                        {
                            
                            //MessageBox.Show("after trans Identify");
                            //Inventor.TextBox finishTextBox = currentSymbol.Definition.Sketch.TextBoxes[1];
                            //MessageBox.Show("finishTextBox " + finishTextBox.Text);
                            ////changeCallout(finishTextBox, currentSymbol);
                            //string value = "value";
                            //currentSymbol.SetPromptResultText(finishTextBox, value);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("text box assignment error: " + ex.Message + "\n" + ex.StackTrace);
                            //ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, "Trans Event Handler " + TransactionObject.DisplayName);
                        }
                    }
                }
            }
            
        }

        public void changeCallout(Inventor.TextBox box, SketchedSymbol symbol)
        {
            MessageBox.Show("symbol " + symbol.Name);
            MessageBox.Show(symbol.GetResultText(box));
            //symbol.SetPromptResultText(box, "New Text");
        }


    }
}
