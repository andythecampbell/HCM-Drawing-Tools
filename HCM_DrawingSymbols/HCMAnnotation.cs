using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace HCMToolsInventorAddIn
{
    public class HCMAnnotation
    {

        //Creates sketch with typical HCM Drawer annotation
        public void DoorAnnotate()
        {
            Inventor.Application oApp  = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            Document currentDoc = oApp.ActiveDocument;
            Transaction doorAnnotationTransaction = null;
            if (currentDoc is DrawingDocument)
            {
                try
                {
                    DrawingDocument oDrawDoc = (DrawingDocument)currentDoc;
                    TransactionManager oTransMan = oApp.TransactionManager;
                    doorAnnotationTransaction = oTransMan.StartTransaction((_Document)oDrawDoc, "Add Door Annotation");
                    Sheet oSheet = oDrawDoc.ActiveSheet;
                    Object hingeSideOb = oApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select Hinge Side");
                    Object swingSideOb = oApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select Swing Side");
                    DrawingCurveSegment hingeEdgeCurveSeg = (DrawingCurveSegment)hingeSideOb;
                    DrawingCurveSegment swingEdgeCurveSeg = (DrawingCurveSegment)swingSideOb;
                    DrawingCurve hingeEdgeCurve = hingeEdgeCurveSeg.Parent;
                    DrawingCurve swingEdgeCurve = swingEdgeCurveSeg.Parent;
                    DrawingView oView = hingeEdgeCurve.Parent;
                    DrawingSketch oSketch = oView.Sketches.Add();
                    SketchLine oSketchHingeSide = (SketchLine)oSketch.AddByProjectingEntity(hingeEdgeCurve);
                    SketchLine oSketchSwingSide = (SketchLine)oSketch.AddByProjectingEntity(swingEdgeCurve);
                    Point2d startPoint = oSketchSwingSide.StartSketchPoint.Geometry;
                    Point2d endPoint = oSketchSwingSide.EndSketchPoint.Geometry;
                    Point2d midPointTransient = oSketchHingeSide.Geometry.MidPoint;
                    oSketch.Edit();
                    SketchLine firstHingeLine = oSketch.SketchLines.AddByTwoPoints(startPoint, oSketchHingeSide.Geometry.MidPoint);
                    SketchLine secondHingeLine = oSketch.SketchLines.AddByTwoPoints(oSketchHingeSide.Geometry.MidPoint, endPoint);
                    oSketch.LineType = LineTypeEnum.kDashedLineType;
                    oSketch.LineWeight = .01;
                    oSketch.ExitEdit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Creating Door Annotation " + ex.Message + ex.StackTrace);
                }
                doorAnnotationTransaction.End();
            }
            else
            {
                MessageBox.Show("Current Document Must Be A Drawing");
            }
        }

        //Creates sketch with typical HCM Drawer annotation
        public void DrawerAnnotate()
        {
            Inventor.Application oApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            Document currentDoc = oApp.ActiveDocument;
            Transaction doorAnnotationTransaction = null;
            if (currentDoc is DrawingDocument)
            {
                try
                {
                    DrawingDocument oDrawDoc = (DrawingDocument)currentDoc;
                    TransactionManager oTransMan = oApp.TransactionManager;
                    doorAnnotationTransaction = oTransMan.StartTransaction((_Document)oDrawDoc, "Add Door Annotation");
                    Sheet oSheet = oDrawDoc.ActiveSheet;
                    Object drawerTop = oApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select Top Of Drawer Front");
                    Object drawerBot = oApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select Bottom Of Drawer Front");
                    DrawingCurveSegment drawerTopCurveSeg = (DrawingCurveSegment)drawerTop;
                    DrawingCurveSegment drawerBotCurveSeg = (DrawingCurveSegment)drawerBot;
                    DrawingCurve drawerTopCurve = drawerTopCurveSeg.Parent;
                    DrawingCurve drawerBotCurve = drawerBotCurveSeg.Parent;
                    DrawingView oView = drawerTopCurve.Parent;
                    DrawingSketch oSketch = oView.Sketches.Add();
                    SketchLine oSketchDrawerTop = (SketchLine)oSketch.AddByProjectingEntity(drawerTopCurve);
                    SketchLine oSketchDrawerBot = (SketchLine)oSketch.AddByProjectingEntity(drawerBotCurve);
                    Point2d startPoint;
                    Point2d endPoint;
                    if (oSketchDrawerTop.StartSketchPoint.Geometry.X > oSketchDrawerTop.EndSketchPoint.Geometry.X)
                    {
                        startPoint = oSketchDrawerTop.StartSketchPoint.Geometry;
                    }
                    else
                    {
                        startPoint = oSketchDrawerTop.EndSketchPoint.Geometry;
                    }
                    if (oSketchDrawerBot.StartSketchPoint.Geometry.X < oSketchDrawerBot.EndSketchPoint.Geometry.X)
                    {
                        endPoint = oSketchDrawerBot.StartSketchPoint.Geometry;
                    }
                    else
                    {
                        endPoint = oSketchDrawerBot.EndSketchPoint.Geometry;
                    }


                    oSketch.Edit();
                    SketchLine firstHingeLine = oSketch.SketchLines.AddByTwoPoints(startPoint, endPoint);
                    oSketch.LineType = LineTypeEnum.kDashedLineType;
                    oSketch.LineWeight = .01;
                    oSketch.ExitEdit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Creating Door Annotation " + ex.Message + ex.StackTrace);
                }
                doorAnnotationTransaction.End();
            }
            else
            {
                MessageBox.Show("Current Document Must Be A Drawing");
            }
        }

    }
}
