using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;

namespace HCMToolsInventorAddIn
{
    public class HcmView
    {

        public HcmView(DrawingView view)
        {
            View = view;
            ViewLabel = FindViewLabel(view);
            _viewToLabel = view.Position.VectorTo(ViewLabel.Position);
            AttachedCallouts = BuildAttachedCallouts(FindAttachedCallouts(view), view);

        }


        public DrawingView View { get; set; }
        public SketchedSymbol ViewLabel { get; }
        public Dictionary<SketchedSymbol, Vector2d> AttachedCallouts { get; }

        private Vector2d _viewToLabel;


        public void SyncLabelLocation()
        {
            if (ViewLabel != null)
            {
                Point2d position = View.Position;
                position.TranslateBy(_viewToLabel);
                ViewLabel.Position = position;
            }
            if (AttachedCallouts.Any())
            {
                
            }
        }

        private Dictionary<SketchedSymbol, Vector2d> BuildAttachedCallouts(List<SketchedSymbol> calloutSymbols,
            DrawingView view)
        {
            Dictionary<SketchedSymbol, Vector2d> dictionary = new Dictionary<SketchedSymbol, Vector2d>();
            foreach (SketchedSymbol sketchedSymbol in calloutSymbols)
            {
                var viewToSymbol = view.Position.VectorTo(sketchedSymbol.Position);
                dictionary.Add(sketchedSymbol, viewToSymbol);
            }
            return dictionary;
        }

        private List<SketchedSymbol> FindAttachedCallouts(DrawingView view)
        {
            List<SketchedSymbol> attachedCalloutSymbols = new List<SketchedSymbol>();
            DrawingDocument drawDoc = (DrawingDocument) view.Parent.Parent;
            foreach (Sheet sheet in drawDoc.Sheets)
            {
                foreach (DrawingView drawingView in sheet.DrawingViews)
                {
                    if ((drawingView is DetailDrawingView || drawingView is SectionDrawingView) &&
                        drawingView.ParentView == view)
                    {
                        attachedCalloutSymbols.Add(FindCallloutSymbol(drawingView));
                    }
                }
            }
            return attachedCalloutSymbols;
        }

        private SketchedSymbol FindViewLabel(DrawingView view)
        {
            DrawingDocument drawDoc = (DrawingDocument) view.Parent.Parent;
            if (view.AttributeSets.NameIsUsed["ViewData"])
            {
                if (view.AttributeSets["ViewData"].NameIsUsed["LabelRefKey"])
                {
                    Object matchType = new Object();
                    byte[] labelRefKey = (byte[]) view.AttributeSets["ViewData"]["LabelRefKey"].Value;
                    object labelObject = drawDoc.ReferenceKeyManager.BindKeyToObject(
                        labelRefKey,
                        0, out matchType);
                    return (SketchedSymbol) labelObject;
                    //SketchedSymbol calloutLabel = (SketchedSymbol)calloutObject;
                    //Vector2d moveBy =
                    //    persistantPlaces[i].VectorTo(persistantDetailViews[i].FenceCornerOne);
                    //Point2d newPosition;
                    //foreach (LeaderNode node in calloutLabel.Leader.AllNodes)
                    //{
                    //    newPosition = node.Position;
                    //    newPosition.TranslateBy(moveBy);
                    //    node.Position = newPosition;
                    //}
                }
            }
            return null;
        }

        private SketchedSymbol FindCallloutSymbol(DrawingView view)
        {
            DrawingDocument drawDoc = (DrawingDocument) view.Parent.Parent;
            if (view.AttributeSets.NameIsUsed["ViewData"])
            {
                if (view.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
                {
                    Object matchType = new Object();
                    byte[] labelRefKey = (byte[]) view.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                    object labelObject = drawDoc.ReferenceKeyManager.BindKeyToObject(
                        labelRefKey,
                        0, out matchType);
                    return (SketchedSymbol) labelObject;
                }
            }
            return null;

        }
    }
}
