using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Inventor;

namespace HCMToolsInventorAddIn
{
    public class ApplicationEventsHandler
    {
        public ApplicationEventsHandler(ApplicationEvents appEvents)
        {
            _appEvents = appEvents;
        }

        List<HcmView> _hcmViews = new List<HcmView>();
        private ApplicationEvents _appEvents;

        public void Register()
        {            
            _appEvents.OnDocumentChange += appEvents_OnDocumentChange; 
            _appEvents.OnSaveDocument += appEvents_OnSaveDocument;
            _appEvents.OnOpenDocument += appEvents_OnOpenDocument;
            _appEvents.OnNewDocument += appEvents_OnNewDocument;
            _appEvents.OnActivateDocument += appEvents_OnActivateDocument;
            //appEvents.OnNewEditObject += appEvents_OnNewEditObject;
        }

        public void Unregister()
        {
            Inventor.ApplicationEvents appEvents = AddinGlobal.InventorApp.ApplicationEvents;
            appEvents.OnDocumentChange -= appEvents_OnDocumentChange;
            appEvents.OnSaveDocument -= appEvents_OnSaveDocument;
            appEvents.OnOpenDocument -= appEvents_OnOpenDocument;
            appEvents.OnNewDocument -= appEvents_OnNewDocument;
            appEvents.OnActivateDocument -= appEvents_OnActivateDocument;
        }
        void appEvents_OnNewEditObject(object EditObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        //registers document events as new documents are created
        void appEvents_OnNewDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                DocumentEvents docEvents = DocumentObject.DocumentEvents;
                docEvents.OnChange += docEvents_OnChange;
                docEvents.OnDelete += docEvents_OnDelete;
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        //registers document events as new documents are activated
        void appEvents_OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        void appEvents_OnDocumentChange(_Document DocumentObject, EventTimingEnum BeforeOrAfter, CommandTypesEnum ReasonsForChange, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            string displayNameString = (string)(Context.get_Value("DisplayName"));
            string[] internalNamesString = new string[10];
            if (Context.Count > 3 && Context.get_Name(3) == "InternalNamesList")
            {
                try
                {
                    internalNamesString = (string[])(Context.get_Value("InternalNamesList"));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("context error" + ex.Message);
                }
            }

            #region After Events
            
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                
                if (ReasonsForChange == CommandTypesEnum.kQueryOnlyCmdType)
                {
                   
                }
                TransactionManager transMan = AddinGlobal.InventorApp.TransactionManager;
                Transaction trans = transMan.StartTransaction(AddinGlobal.InventorApp.ActiveDocument, "zzz");
                   

                        //adds view label when view is placed
                if (displayNameString == "Create Drawing View" || displayNameString == "Rollout Drawing View")
                {
                    try
                    {
                        DrawingDocument drawDoc = (DrawingDocument)DocumentObject;
                        Sheet oSheet = drawDoc.ActiveSheet;
                        int addedViewIndex = oSheet.DrawingViews.Count;
                        DrawingView addedDrawingView = oSheet.DrawingViews[addedViewIndex];
                        addedDrawingView.addCustomViewLabel();
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, trans.DisplayName);
                        MessageBox.Show("Error on View HCM Label Creations " + ex.Message + " " + ex.StackTrace);
                    }

                }

                    //adds view labels and callouts when view is placed
                else if (displayNameString == "Create Profile And Section View")
                {
                    try
                    {
                        DrawingDocument drawDoc = (DrawingDocument)DocumentObject;
                        Sheet oSheet = drawDoc.ActiveSheet;
                        int addedViewIndex = oSheet.DrawingViews.Count;
                        DrawingView addedDrawingView = oSheet.DrawingViews[addedViewIndex];
                        addedDrawingView.addCustomViewLabel();
                        if (addedDrawingView is SectionDrawingView)
                        {
                            SectionDrawingView addedSection = (SectionDrawingView)addedDrawingView;
                            addedSection.addSectionViewSymbol();
                        }
                        else if (addedDrawingView is DetailDrawingView)
                        {
                            DetailDrawingView detailView = (DetailDrawingView)addedDrawingView;
                            detailView.addDetailViewSymbol();
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Problems Creating HCM Section Callout");
                        ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, "Creating Section Callout");
                    }
                }

                    //updates view labels and section callouts as view properties change
                else if (displayNameString == "Edit View Properties")
                {
                    Document doc = AddinGlobal.InventorApp.ActiveDocument;
                    if (doc.ActivatedObject is Sheet)
                    {
                        Sheet ivSheet = (Sheet)doc.ActivatedObject;
                        foreach (DrawingView dView in ivSheet.DrawingViews)  //TODO: Convert loops to linq statements with .Cast
                        {
                            foreach (DrawingView view in dView.Parent.DrawingViews)
                            {
                                var updated = view.updateViewLabel();
                                if (!updated.Success)
                                {
                                    ErrorLog.LogError(updated.Message + " View Label " + dView.Name + doc.DisplayName, doc, trans.DisplayName);
                                }
                            }
                            if (dView is SectionDrawingView)
                            {
                                SectionDrawingView sectionView = (SectionDrawingView)dView;
                                var update = sectionView.updateCalloutSymbol();
                                if (!update.Success)
                                {
                                    ErrorLog.LogError(update.Message + " Callout " + sectionView.Name + doc.DisplayName, doc, trans.DisplayName);
                                }
                            }
                            if (dView is DetailDrawingView)
                            {
                                DetailDrawingView detailView = (DetailDrawingView)dView;
                                var update = detailView.updateCalloutSymbol();
                                if (!update.Success)
                                {
                                    ErrorLog.LogError(update.Message + " Callout " + dView.Name + doc.DisplayName, doc, trans.DisplayName);
                                }
                            }
                        }
                    }
                }

                    //changes or updates view labels and callouts as view is pasted to new sheet
                else if (displayNameString == "Paste View")
                {
                    try
                    {
                        DrawingDocument drawDoc = (DrawingDocument)DocumentObject;
                        Sheet oSheet = drawDoc.ActiveSheet;
                        DrawingView view = oSheet.DrawingViews[oSheet.DrawingViews.Count];
                        view.addCustomViewLabel();
                        if (view is SectionDrawingView)
                        {
                            SectionDrawingView addedSection = (SectionDrawingView)view;
                            addedSection.addSectionViewSymbol();
                        }
                        else if (view is DetailDrawingView)
                        {
                            DetailDrawingView detailView = (DetailDrawingView)view;
                            detailView.addDetailViewSymbol();
                        }
                        foreach (Sheet drawingSheet in drawDoc.Sheets)
                        {
                            foreach (DrawingView drawingView in drawingSheet.DrawingViews)
                            {
                                if (drawingView is SectionDrawingView && drawingView.ParentView == view)
                                {
                                    SectionDrawingView sectionToUpdate = (SectionDrawingView)drawingView;
                                    sectionToUpdate.addSectionViewSymbol();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, trans.DisplayName);
                        MessageBox.Show(ex.Message + " " + ex.StackTrace);
                    }
                }

                    //Updates custom callout as detail fence moves
                else if (displayNameString == "Change Detail View Callout")
                {
                    try
                    {
                        DrawingDocument drawDoc = (DrawingDocument)AddinGlobal.InventorApp.ActiveDocument;
                        Sheet oSheet = drawDoc.ActiveSheet;
                        List<DetailDrawingView> allDetailViews = new List<DetailDrawingView>();
                        foreach (Sheet docSheet in drawDoc.Sheets)
                        {
                            foreach (DrawingView docSheetDrawingView in docSheet.DrawingViews)
                            {
                                if (docSheetDrawingView is DetailDrawingView)
                                {
                                    allDetailViews.Add((DetailDrawingView)docSheetDrawingView);
                                }
                            }
                        }



                        foreach (var detailDrawingView in allDetailViews)
                        {
                            if (detailDrawingView.AttributeSets.NameIsUsed["ViewData"])
                            {
                                if (detailDrawingView.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
                                {
                                    Object matchType = new Object();
                                    byte[] calloutRefKey =
                                        (byte[])
                                        detailDrawingView.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                                    object calloutObject = drawDoc.ReferenceKeyManager.BindKeyToObject(
                                        calloutRefKey,
                                        0, out matchType);
                                    SketchedSymbol calloutLabel = (SketchedSymbol) calloutObject;
                                    var leaderEndPoint =
                                        calloutLabel.Leader.AllNodes[calloutLabel.Leader.AllNodes.Count];
                                    Vector2d moveBy = leaderEndPoint.Position.VectorTo(detailDrawingView.FenceCornerOne);
                                    foreach (LeaderNode node in calloutLabel.Leader.AllNodes)
                                    {
                                        Point2d newPosition = node.Position;
                                        newPosition.TranslateBy(moveBy);
                                        node.Position = newPosition;
                                    }
                                }
                            }
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, trans.DisplayName);
                        MessageBox.Show("After Event Error " + ex.Message + ex.StackTrace);
                    }
                }
                //Updates tags when the sheet order changes
                else if (displayNameString == "Move Sheet Instance")
                {
                    try
                    {
                        DrawingDocument drawDoc = (DrawingDocument)AddinGlobal.InventorApp.ActiveDocument;
                        foreach (Sheet oSheet in drawDoc.Sheets)
                        {
                            foreach (DrawingView oView in oSheet.DrawingViews)
                            {
                                if (oView is SectionDrawingView)
                                {
                                    SectionDrawingView oSectionView = (SectionDrawingView)oView;
                                    oSectionView.updateCalloutSymbol();
                                }
                                else if (oView is DetailDrawingView)
                                {
                                    DetailDrawingView oDetailView = (DetailDrawingView)oView;
                                    oDetailView.updateCalloutSymbol();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, trans.DisplayName);
                        MessageBox.Show("Error on Sheet Move " + ex.Message);
                    }
                }
                    
                //Moves View Tags and Callouts with as drwing view moves
                else if (displayNameString == "Drawing View Center")
                {
                    try
                    {
                        _hcmViews.ForEach(v => v.SyncLabelLocation());
                        _hcmViews.Clear();
                    }
                    catch (Exception ex)
                    {
                        _hcmViews.Clear();
                        MessageBox.Show("Error Moving HCM View Label " + ex.Message + ex.StackTrace);
                    }
                } //end drawing view center
                    
                //recreates section tag if drawing veiw changes
                if (displayNameString == "Reverse direction")
                {
                    DrawingDocument drawDoc = (DrawingDocument)AddinGlobal.InventorApp.ActiveDocument;
                    Sheet oSheet = drawDoc.ActiveSheet;
                    List<DrawingView> viewsOnSheet = new List<DrawingView>();
                    foreach (Sheet drawingSheet in drawDoc.Sheets)
                    {
                        foreach (DrawingView oDrawingView in drawingSheet.DrawingViews)
                        {
                            if (oDrawingView is SectionDrawingView && oDrawingView.ParentView.Parent == oSheet)
                            {
                                SectionDrawingView oSectionView = (SectionDrawingView)oDrawingView;
                                oSectionView.DeleteSectionViewSymbol();
                                oSectionView.addSectionViewSymbol();
                            }
                        }
                    }
                }
                try
                {
                    trans.MergeWithPrevious = true;
                }
                catch (Exception ex)
                {
                    //not sure why this throws
                }

                trans.End();
                    
                }
                if (displayNameString == "Edit Property" &&
                        AddinGlobal.InventorApp.ActiveDocument is DrawingDocument)
                {
                    //implmentation to come
                }
            #endregion
            //End After Evetns

            // Before Events
            else
            {
                if (displayNameString == "Drawing View Center")
                {
                    Inventor.Application app = AddinGlobal.InventorApp;
                    DrawingDocument drawDoc = (DrawingDocument)app.ActiveDocument;
                    Sheet oSheet = drawDoc.ActiveSheet;
                    foreach(object ob in drawDoc.SelectSet)
                    {
                        if (ob is DrawingView)
                        {
                            DrawingView drawView = (DrawingView)ob;
                            _hcmViews.Add(new HcmView(drawView));
                            foreach (DrawingView alignedView in oSheet.DrawingViews)
                            {
                                if (alignedView.ParentView == drawView)
                                {
                                    _hcmViews.Add(new HcmView(alignedView));
                                }
                            }
                        }
                    }
                }
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
        
        //adds documents events on document open
        void appEvents_OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            if (DocumentObject != null)
            {
                DocumentEvents docEvents = DocumentObject.DocumentEvents;
                docEvents.OnChange += docEvents_OnChange;
                docEvents.OnDelete += docEvents_OnDelete;
            }
        }

        //removes custom labels and callouts when view is deleted
        void docEvents_OnDelete(object Entity, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kBefore && Entity is DrawingView)
            {
                try
                {
                    DrawingView oDrawView = (DrawingView)Entity;
                    Sheet oSheet = oDrawView.Parent;
                    DrawingDocument oDrawDoc = (DrawingDocument)oSheet.Parent;
                    Object matchType = new Object();
                    if (oDrawView.AttributeSets.NameIsUsed["ViewData"]
                        && oDrawView.AttributeSets["ViewData"].NameIsUsed["LabelRefKey"])
                    {
                        byte[] labelRefKey = (byte[])oDrawView.AttributeSets["ViewData"]["LabelRefKey"].Value;
                        object viewLabelObject = oDrawDoc.ReferenceKeyManager.BindKeyToObject(labelRefKey,
                            0, out matchType);
                        SketchedSymbol viewLabel = (SketchedSymbol)viewLabelObject;
                        viewLabel.Delete();
                    }

                    if (oDrawView is SectionDrawingView)
                    {
                        SectionDrawingView sectionView = (SectionDrawingView)oDrawView;
                        sectionView.DeleteSectionViewSymbol();
                    }
                    if (oDrawView is DetailDrawingView)
                    {
                        DetailDrawingView detailView = (DetailDrawingView)oDrawView;
                        if (detailView.AttributeSets.NameIsUsed["ViewData"]
                        && detailView.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
                        {
                            object outObject, outContext;
                            byte[] calloutRefKey = (byte[])oDrawView.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                            if(oDrawDoc.ReferenceKeyManager.CanBindKeyToObject(calloutRefKey, 0, out outObject, out outContext))
                            {
                            object viewLabelObject = oDrawDoc.ReferenceKeyManager.BindKeyToObject(calloutRefKey,
                                0, out matchType);
                            
                            SketchedSymbol calloutLabel = (SketchedSymbol)viewLabelObject;
                            calloutLabel.Delete();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error removing callout " + ex.Message + " " + ex.StackTrace);
                }
            }
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        void docEvents_OnChange(CommandTypesEnum ReasonsForChange, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            string displayNameString = (string)(Context.get_Value("DisplayName"));
            string[] internalNamesString = new string[10];
            if (Context.Count > 3 && Context.get_Name(3) == "InternalNamesList")
            {
                try
                {
                    internalNamesString = (string[])(Context.get_Value("InternalNamesList"));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("context error" + ex.Message);
                }
            }
            if (BeforeOrAfter == EventTimingEnum.kBefore)
            {
                if (displayNameString == "Create Drawing View" ||
                    displayNameString == "Rollout Drawing View")
                {

                }
            }
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        void appEvents_OnSaveDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kBefore)
            {
                if (DocumentObject is DrawingDocument)
                {
                    Document drawingDoc = (Document)DocumentObject;
                }
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        void ApplicationEvents_OnActivateView(Inventor.View ViewObject, EventTimingEnum BeforeOrAfter, NameValueMap Context,
            out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }
        
    }
}
