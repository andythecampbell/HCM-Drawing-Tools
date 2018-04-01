using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Inventor;

namespace HCMToolsInventorAddIn
{
    class SyncSymbolsButton : Button
    {
        public SyncSymbolsButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType) : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {
        }

        //public SyncSymbolsButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType) : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        //{
        //}

        public const string _detailCallout = "Detail Call Out";
        public const string _detailLabelTag = "Detail View Label";
        public const string _viewLabelTag = "Prompted View Tag";
        public const string _sectionTagPrefix = "Section View ";
        

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            var app = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            var document = (DrawingDocument)app.ActiveEditDocument;
            var result = UpdateAllSymbols(document);
        }

        private ActionResult UpdateAllSymbols(DrawingDocument doc)
        {
            var calloutResult = UpdateAllCallouts(doc);
            var viewLabelResult = UpdateAllViewLabels(doc);
            if (!calloutResult.Success)
            {
                ErrorLog.LogError(calloutResult.Message, (Document) doc, "Manaul Callout Update");
            }
            if (!viewLabelResult.Success)
            {
                ErrorLog.LogError(calloutResult.Message, (Document) doc, "Manaul View Label Update");
            }
            return new ActionResult(true);
        }

        private ActionResult UpdateAllViewLabels(DrawingDocument doc)
        {
            try
            {
                var views = doc.Sheets.Cast<Sheet>().SelectMany(s => s.DrawingViews.Cast<DrawingView>());
                var orphanedViews = views.Where(v => !v.updateViewLabel().Success);
                foreach (var view in orphanedViews)
                {
                    view.addCustomViewLabel();
                }
                var viewLabelRefKeys = views.Select(GetViewLabelRefkey);
                var viewLabelSymbols = doc.Sheets.Cast<Sheet>().SelectMany(s => s.SketchedSymbols.Cast<SketchedSymbol>())
                    .Where(s => s.Definition.Name == _viewLabelTag);
                foreach (var viewLabel in viewLabelSymbols)
                {
                    var refKey = new byte[10];
                    viewLabel.GetReferenceKey(ref refKey);
                    if (!viewLabelRefKeys.Any(r => r.SequenceEqual(refKey)))
                    {
                        viewLabel.Delete();
                    }
                }
                return new ActionResult(true);
            }
            catch (Exception e)
            {
                return new ActionResult(e);
            }
        }

        private ActionResult UpdateAllCallouts(DrawingDocument doc)
        {
            try
            {
                var views = doc.Sheets.Cast<Sheet>().SelectMany(s => s.DrawingViews.Cast<DrawingView>());
                var sectionViews = views.OfType<SectionDrawingView>();
                var detailViews = views.OfType<DetailDrawingView>();
                var orphanedSectionViews = sectionViews.Where(v => !v.updateCalloutSymbol().Success);
                foreach (var view in orphanedSectionViews)
                {
                    view.addSectionViewSymbol();
                }
                var orphanedDetailViews = detailViews.Where(v => !v.updateCalloutSymbol().Success);
                foreach (var view in orphanedDetailViews)
                {
                    view.addDetailViewSymbol();
                }

                var sketchSymbols = doc.Sheets.Cast<Sheet>().SelectMany(s => s.SketchedSymbols.Cast<SketchedSymbol>());
                var sectionSketchSymbols = sketchSymbols.Where(s => s.Definition.Name.StartsWith(_sectionTagPrefix));
                var detailSketchSymbols = sketchSymbols.Where(s => s.Definition.Name == _detailCallout);
                var syncedKeys = sectionViews.Select(v => GetSectionCalloutRefKey((DrawingView)v)).ToList();
                syncedKeys.AddRange(detailViews.Select(v => GetSectionCalloutRefKey((DrawingView)v)));
                //var syncedSectionSymbols = new List<SketchedSymbol>();// = sectionSketchSymbols.Where(s => !sectionViews.Any(v => GetSectionCalloutRefKey(v) == s.GetReferenceKey()))
                
                foreach (var symbol in sectionSketchSymbols.Concat(detailSketchSymbols))
                {
                    byte[] refKey = { }; ;
                    symbol.GetReferenceKey(ref refKey);
                    if (!syncedKeys.Contains(refKey))
                    {
                        symbol.Delete();
                    }
                }
                return new ActionResult(true);
            }
            catch (Exception ex)
            {
                return new ActionResult(ex);
            }

        }

        private byte[] GetSectionCalloutRefKey(DrawingView view)
        {
            if (view.AttributeSets.NameIsUsed["ViewData"]
                && view.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
            {
                return (byte[]) view.AttributeSets["ViewData"]["CalloutRefKey"].Value;
            }
            return new byte[] { };
        }
        private byte[] GetViewLabelRefkey(DrawingView view)
        {
            if (view.AttributeSets.NameIsUsed["ViewData"]
                && view.AttributeSets["ViewData"].NameIsUsed["LabelRefKey"])
            {
                return (byte[]) view.AttributeSets["ViewData"]["LabelRefKey"].Value;
            }
            return new byte[] { };
        }

    }
}
