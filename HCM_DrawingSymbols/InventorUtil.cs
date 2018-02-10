using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Inventor;
using System.Windows.Forms;
using System.Reflection;

namespace HCMToolsInventorAddIn
{
    public static class InventorUtil
    {
        /// <summary>
        /// Determines the two parallel planes with the greatest distance between them
        /// </summary>
        /// <param name="faceOnPlane">Indicates the plane of the sufaces to look at.</param>
        /// <returns>List of 2 paralle faces that are the furthest apart on the suface body</returns>

        public const string _detailCallout = "Detail Call Out";
        public const string _detailLabelTag = "Detail View Label";
        public const string _viewLabelTag = "Prompted View Tag";
        public const string _sectionTagPrefix = "Section View ";

        public static List<Face> OutsideFaces(Face faceOnPlane)
        {
            Double thickness = 0;
            Plane facePlane = (Plane)faceOnPlane.Geometry;
            Face face1 = null;
            Face face2 = null;
            List<Face> outsideFaces = new List<Face>();
            foreach (Face oFace in faceOnPlane.Parent.Faces)
            {
                if (oFace.Geometry is Plane)
                {
                    Plane oPlane = (Plane)oFace.Geometry;
                    if (oPlane.get_IsParallelTo(facePlane, .01))
                    {
                        outsideFaces.Add(oFace);
                    }
                }
            }
            for (int i = 0; i < outsideFaces.Count; i++)
            {
                Plane loopPlane = (Plane)outsideFaces[i].Geometry;
                if (facePlane.DistanceTo(loopPlane.RootPoint) > thickness)
                {
                    thickness = facePlane.DistanceTo(loopPlane.RootPoint);
                    face1 = outsideFaces[i];
                }
            }
            for (int i = 0; i < outsideFaces.Count; i++)
            {
                Plane loopPlane = (Plane)outsideFaces[i].Geometry;
                if (facePlane.DistanceTo(loopPlane.RootPoint) == thickness)
                {
                    face2 = outsideFaces[i];
                    break;
                }
            }
            outsideFaces.Clear();
            outsideFaces.Add(face1);
            outsideFaces.Add(face2);
            return outsideFaces;
        }

        public static double length(this Inventor.Edge edge)
        {
            return (edge.StartVertex.Point.DistanceTo(edge.StopVertex.Point));
        }

        public static bool NameExists(this RenderStyles styles, string name)
        {
            try
            {
                RenderStyle style = styles[name];
                return true;
            }
            catch
            {
                return false;
            }
        }

        //public static void appearanceDataSet(Document oDoc)
        //{
        //    if (oDoc is PartDocument)
        //    {
        //        PartDocument partDoc = (PartDocument)oDoc;
        //        if (partDoc.ComponentDefinition.Parameters.ParameterTables.Count > 0)
        //        {
        //            ParameterTable paramTable = partDoc.ComponentDefinition.Parameters.ParameterTables[1];
        //            Worksheet oWorksheet = (Worksheet)paramTable.WorkSheet;
        //            Range range = oWorksheet.Range["E2", "E500"];
        //            string cellValue = null;
        //            List<HCM_Finish> finishes = new List<HCM_Finish>();
        //            List<string> finishNames = new List<string>();
        //            foreach (Microsoft.Office.Interop.Excel.Range cell in range)
        //            {
        //                try
        //                {
        //                    cellValue = (string)cell.Text;
        //                }
        //                catch
        //                {
        //                    continue;
        //                }
        //                if (cellValue.Length > 4 && !finishNames.Contains(cellValue))
        //                {
        //                    int row = cell.Row;
        //                    finishes.Add(new HCM_Finish(cellValue, oWorksheet.Range["F" + row].Text,
        //                        oWorksheet.Range["G" + row].Text, oWorksheet.Range["H" + row].Text));
        //                    finishNames.Add(cellValue);
        //                }
        //            }
        //            int matchCount = 0;
        //            RenderStyles styles = partDoc.RenderStyles;
        //            foreach (RenderStyle oStyle in partDoc.RenderStyles)
        //            {
        //                foreach (HCM_Finish finish in finishes)
        //                {
        //                    if (finish.FullName == oStyle.Name)
        //                    {
        //                        RenderStyle matchedStyle = oStyle;
        //                        RenderStyle copiedStyle = null;
        //                        if (matchedStyle.StyleLocation == StyleLocationEnum.kLibraryStyleLocation)
        //                        {
        //                            matchedStyle = matchedStyle.ConvertToLocal();
        //                        }
        //                        try
        //                        {
        //                            copiedStyle = matchedStyle.Copy(matchedStyle.Name + "cpy1");
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            MessageBox.Show("Copy Error " +matchedStyle.Name + ex.Message);
        //                            continue;
        //                        }
        //                        copiedStyle.updateStyleAttribute("FinishData", "Description1", ValueTypeEnum.kStringType,
        //                            finish.Description1);
        //                        copiedStyle.updateStyleAttribute("FinishData", "Description2", ValueTypeEnum.kStringType,
        //                           finish.Description2);
        //                        copiedStyle.updateStyleAttribute("FinishData", "Description3", ValueTypeEnum.kStringType,
        //                           finish.Description3);
        //                        MessageBox.Show(copiedStyle.Name + ", " + copiedStyle.StyleLocation);
        //                        try
        //                        {
        //                            //copiedStyle.SaveToGlobal();
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            MessageBox.Show("Save to Global Error: " + ex.Message);
        //                        }
        //                    }
        //                }
        //            }
        //            foreach (RenderStyle style in partDoc.RenderStyles)
        //            {
        //                if (!style.UpToDate)
        //                {
        //                    style.SaveToGlobal();
        //                }
        //            }
        //            MessageBox.Show("Match Count: " + matchCount);
        //        }
        //    }
        //}

        
        public static void updateStyleAttribute(this RenderStyle style, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = style.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateBodyAttribute(this SurfaceBody body, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {

            AttributeSets attSets = body.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateEdgeAttribute(this Edge edge, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = edge.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateDrawingAttribute(this DrawingDocument drawing, String attributeSetName, string attributeName,
           ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = drawing.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static ParameterTable getMaterialTable(this DrawingDocument drawDoc)
        {
            ParameterTable matTable = null;
            byte[] tableRefKey = {};
                foreach (Document doc in drawDoc.ReferencedDocuments)
                {
                    if (doc is AssemblyDocument)
                    {
                        AssemblyDocument asemDoc = (AssemblyDocument)doc;
                        foreach(ParameterTable table in asemDoc.ComponentDefinition.Parameters.ParameterTables)
                        {
                            if (table.FileName.Contains("aterial"))
                            {
                                matTable = table;
                                table.GetReferenceKey(ref tableRefKey, 0);
                                break;
                            }
                        }
                        foreach(PartDocument partDoc in asemDoc.ReferencedNestedPartDocuements())
                        {
                            foreach (SurfaceBody body in partDoc.ComponentDefinition.SurfaceBodies)
                            {
                                foreach(SurfaceBody derBody in body.DerivedBaseBody())
                                {
                                    if (derBody.ComponentDefinition.Document is PartDocument)
                                    {
                                        PartDocument derPart = (PartDocument)derBody.ComponentDefinition.Document;
                                        foreach (ParameterTable table in derPart.ComponentDefinition.Parameters.ParameterTables)
                                        {
                                            if (table.FileName.Contains("aterial"))
                                            {
                                                matTable = table;
                                                table.GetReferenceKey(ref tableRefKey, 0);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                    else if(doc is PartDocument)
                    {
                        PartDocument partDoc = (PartDocument)doc;
                        foreach (ParameterTable table in partDoc.ComponentDefinition.Parameters.ParameterTables)
                        {
                            if (table.FileName.Contains("aterial"))
                            {
                                matTable = table;
                                break;
                            }
                        }
                    }
                }
            return matTable;
        }

        //public static void updateFinishSchedule(this DrawingDocument drawDoc)
        //{
        //    Inventor.Application app = AddinGlobal.InventorApp;
        //    Sheet oSheet = drawDoc.Sheets[1];
        //    List<SketchedSymbol> calloutSymbols = new List<SketchedSymbol>();
        //    SortedSet<string> finishList = new SortedSet<string>();
        //    foreach (Sheet anySheet in drawDoc.Sheets)
        //    {
        //        foreach (SketchedSymbol symbol in anySheet.SketchedSymbols)
        //        {
        //            if (symbol.Name == "SCH-DESC-FINISH")
        //            {
        //                calloutSymbols.Add(symbol);
        //            }
        //            if (symbol.Name == "ANN-FINISH")
        //            {
        //                Inventor.TextBox box = symbol.Definition.Sketch.TextBoxes[1];
        //                finishList.Add(symbol.GetResultText(box));
        //            }
        //        }
        //    }
        //    if (finishList.Count < 1)
        //    {
        //        MessageBox.Show("No Finish Callouts");
        //        return;
        //    }
        //    ParameterTable paramTable = drawDoc.getMaterialTable();
        //    Worksheet workSheet = null;
        //    try
        //    {
        //        workSheet = (Worksheet)paramTable.WorkSheet;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Problem Accessing Spreadsheet\n" + ex.Message + ex.StackTrace);
        //        return;
        //    }
        //    if (calloutSymbols.Count > finishList.Count)
        //    {
        //        for (int i = calloutSymbols.Count; i > finishList.Count; i--)
        //        {
        //            calloutSymbols[i - 1].Delete();
        //        }
        //    }
        //    else if (calloutSymbols.Count < finishList.Count)
        //    {
        //        SketchedSymbolDefinition symbolDef = drawDoc.SketchedSymbolDefinitions["SCH-DESC-FINISH"];
        //        Point2d position;
        //        string[] promptString = { " ", " ", " ", " " };
        //        for (int i = calloutSymbols.Count; i < finishList.Count; i++)
        //        {
        //            if (calloutSymbols.Count < 1)
        //            {
        //                position = app.TransientGeometry.CreatePoint2d(40.005, 26.3525);
        //            }
        //            else
        //            {
        //                double yCoord = calloutSymbols.Last().Position.Y - 1.27;
        //                position = app.TransientGeometry.CreatePoint2d(40.005, yCoord);
        //            }
        //            SketchedSymbol symbol = oSheet.SketchedSymbols.Add(symbolDef, position, 0, 1, promptString);
        //            calloutSymbols.Add(symbol);
        //        }
        //    }
        //    Inventor.TextBox nameTextBox = calloutSymbols[0].Definition.Sketch.TextBoxes[5];
        //    Inventor.TextBox d1TextBox = calloutSymbols[0].Definition.Sketch.TextBoxes[2];
        //    Inventor.TextBox d2TextBox = calloutSymbols[0].Definition.Sketch.TextBoxes[3];
        //    Inventor.TextBox d3TextBox = calloutSymbols[0].Definition.Sketch.TextBoxes[4];
        //    int finishCounter = 0;
        //    foreach (string finishName in finishList)
        //    {
        //        if (calloutSymbols[finishCounter].GetResultText(nameTextBox) != finishName)
        //        {
        //            int row = 1;
        //            Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("E2", "E500");
        //            foreach (Microsoft.Office.Interop.Excel.Range cell in range)
        //            {
        //                string cellValue = (string)cell.Text;
        //                if ((string)cell.Value == null)
        //                {
        //                    continue;
        //                }
        //                if (cellValue.Length > 4 && cellValue.Substring(4) == finishName)
        //                {
        //                    row = cell.Row;
        //                    break;
        //                }
        //            }
        //            calloutSymbols[finishCounter].SetPromptResultText(nameTextBox, finishName);
        //            if (row != 1)
        //            {
        //                Microsoft.Office.Interop.Excel.Range d1Cell = workSheet.get_Range("F" + row, "F" + row);
        //                Microsoft.Office.Interop.Excel.Range d2Cell = workSheet.get_Range("G" + row, "G" + row);
        //                Microsoft.Office.Interop.Excel.Range d3Cell = workSheet.get_Range("H" + row, "H" + row);
        //                calloutSymbols[finishCounter].SetPromptResultText(d1TextBox, (string)d1Cell.Text);
        //                calloutSymbols[finishCounter].SetPromptResultText(d2TextBox, (string)d2Cell.Text);
        //                calloutSymbols[finishCounter].SetPromptResultText(d3TextBox, (string)d3Cell.Text);
        //            }
        //        }
        //        finishCounter++;
        //    }
        //    try
        //    {
                
        //        Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)workSheet.Parent;
        //        workbook.Close(false, Type.Missing, false);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Problem Closing Excel File " + ex.Message + ex.StackTrace);
        //    }
        //}

        public static void AlignVert(this Leader baseLeader, List<Leader> leadersToAlign)
        {
            Inventor.Application app = AddinGlobal.InventorApp;
            List<double> xNodeValues = new List<double>();
            double[] coords = new double[2];
            foreach (LeaderNode node in baseLeader.AllNodes)
            {
                xNodeValues.Add(node.Position.X);
            }

            foreach (Leader leader in leadersToAlign)
            {
                for (int i = 1; i < xNodeValues.Count; i++)
                {
                    Point2d position = app.TransientGeometry.CreatePoint2d(xNodeValues[i - 1], 
                        leader.AllNodes[i].Position.Y);
                    leader.AllNodes[i].Position = position;
                }
            }
        }

        public static List<RenderStyle> allRenderstyles(this DrawingDocument drawDoc)
        {
            List<RenderStyle> renderstyles = new List<RenderStyle>();
            foreach (Sheet oSheet in drawDoc.Sheets)
            {
                foreach (SketchedSymbol oSymbol in oSheet.SketchedSymbols)
                {
                    if (oSymbol.Definition.Name == "ANN-FINISH")
                    {
                        if (oSymbol.AttributeSets.NameIsUsed["SymbolData"] &&
                            oSymbol.AttributeSets["SymbolData"].NameIsUsed["FinishName"])
                        {
                            Inventor.TextBox finishTextBox = oSymbol.Definition.Sketch.TextBoxes[1];
                            string finishName = (string)oSymbol.AttributeSets["SymbolData"]["FinishName"].Value;
                            RenderStyle renderStyle = drawDoc.RenderStyles[finishName];
                            renderstyles.Add(renderStyle);
                        }
                    }
                }
            }
            return renderstyles.Distinct().ToList();
        }

        public static void updateViewAttribute(this DrawingView view, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = view.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateSymbolAttribute(this SketchedSymbol symbol, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = symbol.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateFaceAttribute(this Face face, String attributeSetName, string attributeName,
            ValueTypeEnum attributeType, object value)
        {
            AttributeSets attSets = face.AttributeSets;
            AttributeSet attSet;
            Inventor.Attribute attribute;
            if (attSets.NameIsUsed[attributeSetName])
            {
                attSet = attSets[attributeSetName];
                if (attSet.NameIsUsed[attributeName])
                {
                    attribute = attSet[attributeName];
                    attribute.Value = value;
                }
                else
                {
                    attribute = attSet.Add(attributeName, attributeType, value);
                }
            }
            else
            {
                attSet = attSets.Add(attributeSetName, false);
                attribute = attSet.Add(attributeName, attributeType, value);
            }
        }

        public static void updateIProperty(this Document doc, String propName, object value)
        {
            Property prop = null;
            PropertySet customProps = doc.PropertySets["Inventor User Defined Properties"];
            bool propExists = true;
            try
            {
                prop = customProps[propName];
            }
            catch (Exception ex)
            {
                propExists = false;
            }
            if (!propExists)
            {
                prop = customProps.Add(value, propName, null);
                
            }
            else
            {
                prop.Value = value;
            }
        }
        
        //public static string getTableValue(this ParameterTable table, string valueToMatch, string columnReturned)
        //{
        //    string cellValue = "";
        //    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)table.WorkSheet;
        //    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "A500");
        //    int row = 0;
        //    string finishName;
        //    foreach (Microsoft.Office.Interop.Excel.Range cell in range)
        //    {
        //        if ((string)cell.Value2 == valueToMatch)
        //        {
        //            row = cell.Row;
        //            break;
        //        }
        //    }
        //    if (row != 0)
        //    {
        //        range = workSheet.get_Range(columnReturned + row, columnReturned + row);
        //        finishName = (string)range.Text;
        //        cellValue = (string)range.Text;
        //    }
        //    return cellValue;
        //}

        public static List<PartDocument> ReferencedNestedPartDocuements(this AssemblyDocument assemDoc)
        {
            List<PartDocument> partDocs = new List<PartDocument>();
            foreach (Document refDoc in assemDoc.ReferencedDocuments)
            {
                if (refDoc is PartDocument)
                {
                    partDocs.Add((PartDocument)refDoc);
                }
                else if (refDoc is AssemblyDocument)
                {
                    foreach(PartDocument partDoc in ReferencedNestedPartDocuements((AssemblyDocument)refDoc))
                    {
                        partDocs.Add(partDoc);
                    }
                }
            }
            return partDocs.Distinct().ToList();
        }

        public static List<SurfaceBody> DerivedBaseBody(this SurfaceBody partBody)
        {
            List<SurfaceBody> derBodies = new List<SurfaceBody>();
            bool toHere = false;
            if(partBody.CreatedByFeature is ReferenceFeature)
            {
                Document doc = (Document)partBody.ComponentDefinition.Document;
                if (doc is PartDocument)
                {
                    PartDocument partDocument = (PartDocument)doc;
                    ReferenceComponents refComps = partDocument.ComponentDefinition.ReferenceComponents;
                    if (refComps.DerivedPartComponents.Count > 0)
                    {
                        DerivedPartComponents derComponents = refComps.DerivedPartComponents;
                        DerivedPartEntities derSolids = null;
                        SurfaceBody derBody;
                        foreach (DerivedPartComponent derComponent in derComponents)
                        {
                            try
                            {
                                if (derComponent.HealthStatus == HealthStatusEnum.kUpToDateHealth && derComponent.LinkedToFile
                                    && !derComponent.SuppressLinkToFile)
                                {
                                    toHere = true;
                                    derSolids = derComponent.Definition.Solids;
                                    foreach (DerivedPartEntity derEntity in derSolids)
                                    {
                                        if (derEntity.ReferencedEntity is SurfaceBody)
                                        {
                                            derBody = (SurfaceBody)derEntity.ReferencedEntity;
                                            if (derEntity.IncludeEntity)
                                            {
                                                derBodies.AddRange(derBody.DerivedBaseBody());
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Still FUBAR " + toHere + partDocument.DisplayName + ex.Message);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("PartBody " + partBody.Name + " not in Part Document");
                }
            }
            else
            {
                derBodies.Add(partBody);
            }
            return derBodies;
        }
        /// <summary>
        /// Finds the non-ModelParameter expression.
        /// </summary>
        /// <param name="modelParameter">Parameter that expression is being requested of</param>
        /// <param name="allModelParameters">Collection that Parameter belongs to</param>
        /// <returns>First Expression not in allModelParameters collection</returns>

        public static string finalExpression(this ModelParameter modelParameter, ModelParameters allModelParameters)
        {
            string finalExpressionString = modelParameter.Expression;
            if (modelParameter.Expression.Substring(0, 1) == "d")
            {
                foreach (ModelParameter referencedParam in allModelParameters)
                {
                    if (referencedParam.Name == modelParameter.Expression)
                    {
                        finalExpressionString = finalExpression(referencedParam, allModelParameters);
                    }
                }
            }
            return finalExpressionString;
        }

        public static List<ComponentOccurrence> allNestedComponentOccurences(this AssemblyDocument asemDocument)
        {
            List<ComponentOccurrence> compOccurs = new List<ComponentOccurrence>();
            foreach(ComponentOccurrence compOcc in asemDocument.ComponentDefinition.Occurrences)
            {
                if (compOcc.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    compOccurs.Add(compOcc);
                }
                else if (compOcc.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    AssemblyDocument compAsemDoc = (AssemblyDocument)compOcc.Definition.Document;
                    compOccurs.AddRange(compAsemDoc.allNestedComponentOccurences());
                }
            }
            return compOccurs;
        }

        /// <summary>
        /// takes 2 faces as argument and returns true if the faces are on parallel planes
        /// </summary>
        /// <param name="face1"></param>
        /// <param name="face2"></param>
        /// <returns>Ture if faces are on parallel planes</returns>
        public static bool IsParallelTo(this Face face1, Face face2)
        {
            SurfaceEvaluator surfaceEval1 = face1.Evaluator;
            SurfaceEvaluator surfaceEval2 = face2.Evaluator;
            Box2d paramRangeBox = surfaceEval1.ParamRangeRect;
            double[] parmsCenter1 = new double[2];
            double[] normals1 = new double[3];
            double[] parmsCenter2 = new double[2];
            double[] normals2 = new double[3];
            double U = paramRangeBox.MinPoint.X;
            double V = paramRangeBox.MinPoint.X;
            parmsCenter1[0] = (U + V) / 2;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[1] = (U + V) / 2;
            try
            {
                surfaceEval1.GetNormal(ref parmsCenter1, ref normals1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Is Normal Error" + ex.Message + ex.StackTrace);
            }
            paramRangeBox = surfaceEval2.ParamRangeRect;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[0] = (U + V) / 2;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[1] = (U + V) / 2;
            try
            {
                surfaceEval2.GetNormal(ref parmsCenter2, ref normals2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Is Normal Error" + ex.Message + ex.StackTrace);
            }
            for (int i = 0; i < 3; i++)
            {
                normals1[i] = Math.Round(Math.Abs(normals1[i]), 4);
                normals2[i] = Math.Round(Math.Abs(normals2[i]), 4);
            }
            return Enumerable.SequenceEqual(normals1, normals2);
        }


        public static bool IsParallelTo(this Face face1, Plane plane)
        {
            SurfaceEvaluator surfaceEval1 = face1.Evaluator;
            SurfaceEvaluator surfaceEval2 = plane.Evaluator;
            Box2d paramRangeBox = surfaceEval1.ParamRangeRect;
            double[] parmsCenter1 = new double[2];
            double[] normals1 = new double[3];
            double[] parmsCenter2 = new double[2];
            double[] normals2 = new double[3];
            double U = paramRangeBox.MinPoint.X;
            double V = paramRangeBox.MinPoint.X;
            parmsCenter1[0] = (U + V) / 2;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[1] = (U + V) / 2;
            try
            {
                surfaceEval1.GetNormal(ref parmsCenter1, ref normals1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Is Normal Error" + ex.Message + ex.StackTrace);
            }
            paramRangeBox = surfaceEval2.ParamRangeRect;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[0] = (U + V) / 2;
            U = paramRangeBox.MinPoint.X;
            V = paramRangeBox.MinPoint.X;
            parmsCenter1[1] = (U + V) / 2;
            try
            {
                surfaceEval2.GetNormal(ref parmsCenter2, ref normals2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Is Normal Error" + ex.Message + ex.StackTrace);
            }
            for (int i = 0; i < 3; i++)
            {
                normals1[i] = Math.Round(Math.Abs(normals1[i]), 4);
                normals2[i] = Math.Round(Math.Abs(normals2[i]), 4);
            }
            return Enumerable.SequenceEqual(normals1, normals2);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="face">input face</param>
        /// <returns>The Perimeter EdgeLoop</returns>
        public static EdgeLoop PerimeterLoop(this Face face)
        {
            foreach (EdgeLoop edgeLoop in face.EdgeLoops)
            {
                if (edgeLoop.IsOuterEdgeLoop)
                    return edgeLoop;
            }
            return null;
        }
        
        public static void addCustomViewLabel(this DrawingView drawingView, int index)
        {
            Sheet oSheet = drawingView.Parent;
            DrawingDocument drawingDoc = (DrawingDocument)oSheet.Parent;
            DrawingViewLabel viewLabel = drawingView.Label;
            double labelX = drawingView.Left+.5;
            double labelY = drawingView.Top - drawingView.Height-1.2;
            Point2d viewPosition = AddinGlobal.InventorApp.TransientGeometry.CreatePoint2d(labelX, labelY);
            string[] viewPrompts = new string[3];
            viewPrompts[0] = drawingView.Name;
            viewPrompts[1] = drawingView.ScaleString;
            if (drawingView.ViewType == DrawingViewTypeEnum.kStandardDrawingViewType
                || drawingView.ViewType == DrawingViewTypeEnum.kProjectedDrawingViewType)
            {
                switch (drawingView.Camera.ViewOrientationType)
                {
                    case ViewOrientationTypeEnum.kTopViewOrientation:
                        viewPrompts[2] = "PLAN VIEW";
                        break;
                    case ViewOrientationTypeEnum.kBottomViewOrientation:
                        viewPrompts[2] = "BOTTOM VIEW";
                        break;
                    case ViewOrientationTypeEnum.kFrontViewOrientation:
                        viewPrompts[2] = "FRONT ELEVATION";
                        break;
                    case ViewOrientationTypeEnum.kRightViewOrientation:
                    case ViewOrientationTypeEnum.kLeftViewOrientation:
                        viewPrompts[2] = "SIDE ELEVATION";
                        break;
                    case ViewOrientationTypeEnum.kBackViewOrientation:
                        viewPrompts[2] = "BACK ELEVATION";
                        break;
                    default:
                        viewPrompts[2] = "ISOMETRIC VIEW";
                        break;
                }
            }
            else if (drawingView.ViewType == DrawingViewTypeEnum.kSectionDrawingViewType)
            {
                switch (drawingView.Camera.ViewOrientationType)
                {
                    case ViewOrientationTypeEnum.kTopViewOrientation:
                        viewPrompts[2] = "PLAN SECTION";
                        break;
                    case ViewOrientationTypeEnum.kFrontViewOrientation:
                        viewPrompts[2] = "ELEVATED SECTION";
                        break;
                    case ViewOrientationTypeEnum.kRightViewOrientation:
                    case ViewOrientationTypeEnum.kLeftViewOrientation:
                        viewPrompts[2] = "CROSS SECTION";
                        break;
                    case ViewOrientationTypeEnum.kBackViewOrientation:
                        viewPrompts[2] = "BACK ELEVATED SECTION";
                        break;
                    default:
                        viewPrompts[2] = "SECTION";
                        break;
                }
            }
            else if (drawingView.ViewType == DrawingViewTypeEnum.kDetailDrawingViewType)
            {
                viewPrompts[2] = "DETAIL";
            }
            SketchedSymbolDefinition hcmViewLabelDef = drawingDoc.SketchedSymbolDefinitions[_viewLabelTag];
            SketchedSymbol hcmViewLabel = oSheet.SketchedSymbols.Add(hcmViewLabelDef, viewPosition, 0, 1, viewPrompts);
            byte[] labelRefKey = new byte[10];
            hcmViewLabel.GetReferenceKey(ref labelRefKey);
            drawingView.updateViewAttribute("ViewData", "LabelRefKey", ValueTypeEnum.kByteArrayType,
                labelRefKey);
            drawingView.ShowLabel = false;
        }

        public static void updateViewLabel(this DrawingView drawingView)
        {
            Sheet oSheet = drawingView.Parent;
            DrawingDocument drawDoc = (DrawingDocument)oSheet.Parent;
            Inventor.Application app = (Inventor.Application)oSheet.Application;
            if (drawingView.ShowLabel)
            {
                drawingView.addCustomViewLabel(100);
            }
            else if (drawingView.AttributeSets.NameIsUsed["ViewData"]
                && drawingView.AttributeSets["ViewData"].NameIsUsed["LabelRefKey"])
            {
                try
                {
                    byte[] labelSymbolindex = (byte[])drawingView.AttributeSets["ViewData"]["LabelRefKey"].Value;
                    object matchType = new object();
                    SketchedSymbol viewLabelSymbol = (SketchedSymbol)drawDoc.ReferenceKeyManager.BindKeyToObject(
                        ref labelSymbolindex, 0, out matchType);
                    Inventor.TextBox viewNameTextBox = viewLabelSymbol.Definition.Sketch.TextBoxes[2];
                    Inventor.TextBox viewScaleTextBox = viewLabelSymbol.Definition.Sketch.TextBoxes[3];
                    viewLabelSymbol.SetPromptResultText(viewNameTextBox, drawingView.Name);
                    viewLabelSymbol.SetPromptResultText(viewScaleTextBox, drawingView.ScaleString);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error updating Label " + ex.Message + " " + ex.StackTrace);
                }
            }
        }

        public static void updateCalloutSymbol(this SectionDrawingView oSectionView)
        {
            DrawingView parentView = oSectionView.ParentView;
            Sheet parentSheet = parentView.Parent;
            Inventor.Application app = (Inventor.Application)parentSheet.Application;
            DrawingDocument parentDoc = (DrawingDocument)parentSheet.Parent;
            
            if (oSectionView.AttributeSets.NameIsUsed["ViewData"]
                && oSectionView.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
            {
                byte[] calloutRefKey = (byte[])oSectionView.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                object matchType = new object();
                SketchedSymbol calloutSymbol = (SketchedSymbol)parentDoc.ReferenceKeyManager.BindKeyToObject(
                    ref calloutRefKey, 0, out matchType);
                Inventor.TextBox viewNameBox = calloutSymbol.Definition.Sketch.TextBoxes[2];
                Inventor.TextBox pageNumBox = calloutSymbol.Definition.Sketch.TextBoxes[3];
                int pageNumIndex = oSectionView.Parent.Name.IndexOf(":") + 1;
                //string itemNum = (string)parentDoc.PropertySets["Inventor User Defined Properties"]["ITEM CODE"].Value;
                string pageNumber = oSectionView.Parent.Name.Substring(pageNumIndex); //itemNum + "." + oSectionView.Parent.Name.Substring(pageNumIndex);
                calloutSymbol.SetPromptResultText(viewNameBox, oSectionView.Name);
                calloutSymbol.SetPromptResultText(pageNumBox, pageNumber);
            }
        }

        public static void updateCalloutSymbol(this DetailDrawingView oDetailView)
        {
            DrawingView parentView = oDetailView.ParentView;
            Sheet parentSheet = parentView.Parent;
            DrawingDocument parentDoc = (DrawingDocument)parentSheet.Parent;

            if (oDetailView.AttributeSets.NameIsUsed["ViewData"]
                && oDetailView.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
            {
                try
                {                                                                      
                    byte[] calloutRefKey = (byte[])oDetailView.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                    object matchType = new object();
                    SketchedSymbol calloutSymbol = (SketchedSymbol)parentDoc.ReferenceKeyManager.BindKeyToObject(
                        ref calloutRefKey, 0, out matchType);
                    //MessageBox.Show("CalloutName " + calloutSymbol.Name);
                    Inventor.TextBox viewNameBox = calloutSymbol.Definition.Sketch.TextBoxes[2];
                    Inventor.TextBox pageNumBox = calloutSymbol.Definition.Sketch.TextBoxes[3];
                    int pageNumIndex = oDetailView.Parent.Name.IndexOf(":") + 1;
                    //string itemNum = (string)parentDoc.PropertySets["Inventor User Defined Properties"]["ITEM CODE"].Value;
                    string pageNumber = oDetailView.Parent.Name.Substring(pageNumIndex); //itemNum + "." + oDetailView.Parent.Name.Substring(pageNumIndex);
                    calloutSymbol.SetPromptResultText(viewNameBox, oDetailView.Name);
                    calloutSymbol.SetPromptResultText(pageNumBox, pageNumber);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Callout Update Error: " + ex.Message + " " + ex.StackTrace);
                }
            }
        }

        public static void addDetailViewSymbol(this DetailDrawingView newDetail)
        {
            
            DrawingView addedDrawingView = (DrawingView)newDetail;
            Sheet oSheet = addedDrawingView.Parent;
            DrawingView parentView = newDetail.ParentView;
            Sheet parentSheet = parentView.Parent;
            DrawingDocument drawDoc = (DrawingDocument)oSheet.Parent;
            Inventor.Application app = (Inventor.Application)oSheet.Application;
            string[] symbolPropts = new string[2];
            symbolPropts[0] = addedDrawingView.Name;
            int pageNumIndex = oSheet.Name.IndexOf(":") + 1;
            string pageNumber = oSheet.Name.Substring(pageNumIndex);
            symbolPropts[1] = pageNumber;
            DetailDrawingView addedDetail = (DetailDrawingView)addedDrawingView;
            if (!addedDetail.CircularFence)
            {
                Point2d fencePickCorner = addedDetail.FenceCornerOne;
                int xPos = 1;
                int yPos = 1;
                if (fencePickCorner.X < addedDetail.FenceCenter.X)
                    xPos = -1;
                if (fencePickCorner.Y < addedDetail.FenceCenter.Y)
                    yPos = -1;
                Point2d midLeaderPoint = app.TransientGeometry.CreatePoint2d(fencePickCorner.X + (xPos * 1),
                    fencePickCorner.Y + (yPos * 1));
                Point2d endLeaderPoint = app.TransientGeometry.CreatePoint2d(fencePickCorner.X + (xPos * 2),
                    fencePickCorner.Y + (yPos * 1));
                ObjectCollection leaderPoints = app.TransientObjects.CreateObjectCollection();
                leaderPoints.Add(endLeaderPoint);
                leaderPoints.Add(midLeaderPoint);
                leaderPoints.Add(fencePickCorner);
                SketchedSymbolDefinition detailSymbolDef = drawDoc.SketchedSymbolDefinitions[_detailCallout];
                SketchedSymbol newDetailCallout = parentSheet.SketchedSymbols.AddWithLeader(detailSymbolDef,
                    leaderPoints, 0, 1, symbolPropts, true, false);
                newDetailCallout.Leader.ArrowheadType = ArrowheadTypeEnum.kNoneArrowheadType;
                byte[] calloutRefKey = new byte[10];
                newDetailCallout.GetReferenceKey(ref calloutRefKey);
                addedDrawingView.updateViewAttribute("ViewData", "CalloutRefKey", ValueTypeEnum.kByteArrayType,
                    calloutRefKey);
                byte[] viewRefKey = new byte[10];
                newDetail.GetReferenceKey(ref viewRefKey);
                newDetailCallout.updateSymbolAttribute("SymbolData", "ViewRefKey", ValueTypeEnum.kByteArrayType,
                    calloutRefKey);
            }

        }

        public static void addSectionViewSymbol(this SectionDrawingView addedSection)
        {
            DrawingView drawView = (DrawingView)addedSection;
            DrawingView parentView = addedSection.ParentView;
            Vector sectionVector = addedSection.Camera.Eye.VectorTo(addedSection.Camera.Target);
            Vector parentVector = parentView.Camera.Eye.VectorTo(parentView.Camera.Target);
            UnitVector sectionUpVector = addedSection.Camera.UpVector;
            UnitVector parentUpVector = parentView.Camera.UpVector;
            Vector crossVector = sectionVector.CrossProduct(parentVector);
            UnitVector crossUpVector = sectionUpVector.CrossProduct(parentUpVector);
            Sheet oSheet = parentView.Parent;
            Sheet sectionSheet = addedSection.Parent;
            DrawingDocument drawDoc = (DrawingDocument)oSheet.Parent;
            Vector crossParenttoUp = parentVector.CrossProduct(parentUpVector.AsVector());
            double sectionToCrossParentUp = sectionVector.AngleTo(crossParenttoUp);
            double upAngle = parentUpVector.AngleTo(sectionUpVector);
            double sectionToParentUpAngle = sectionVector.AngleTo(parentUpVector.AsVector());
            string sectionCalloutName = _sectionTagPrefix;
            DrawingSketch sectionSketch = addedSection.SectionLineSketch;
            SketchLine sectionLine = sectionSketch.SketchLines[1]; //This is not always correct!! 
            double sectionLineLength = sectionLine.Length;
            //double crossUpTotal = crossUpVector.X + crossUpVector.Y + crossUpVector.Z;
            if (upAngle == 0)
            {
                sectionCalloutName += "V";

                foreach (SketchLine sketchline in sectionSketch.SketchLines)
                {
                    if (sketchline.Geometry.Direction.Y > .1 || sketchline.Geometry.Direction.Y < -.1)
                    {
                        sectionLine = sketchline;
                        break;
                    }
                }
                if (sectionLine.Geometry.Direction.Y < 0)
                {
                    sectionCalloutName += "U";
                }
                else
                {
                    sectionCalloutName += "D";
                }
                if (sectionToCrossParentUp == 0)
                {
                    sectionCalloutName += "R";
                }
                else
                {
                    sectionCalloutName += "L";
                }
            }
            else
            {
                sectionCalloutName += "H";
                foreach (SketchLine sketchline in sectionSketch.SketchLines)
                {
                    if (sketchline.Geometry.Direction.X > .1 || sketchline.Geometry.Direction.X < -.1)
                    {
                        sectionLine = sketchline;
                        break;
                    }
                }
                if (sectionToParentUpAngle == 0)
                {
                    sectionCalloutName += "U";
                }
                else
                {
                    sectionCalloutName += "D";
                }
                if (sectionLine.Geometry.Direction.X > 0)
                {
                    sectionCalloutName += "L";
                }
                else
                {
                    sectionCalloutName += "R";
                }
            }

            //if (crossUpTotal > .1)
            //{
            //    sectionCalloutName += "H";
            //    foreach (SketchLine sketchline in sectionSketch.SketchLines)
            //    {
            //        if (sketchline.Geometry.Direction.X > .1 || sketchline.Geometry.Direction.X < -.1)
            //        {
            //            sectionLine = sketchline;
            //            break;
            //        }
            //    }
            //    if (sectionLine.Geometry.Direction.X > 0)
            //    {
            //        sectionCalloutName += "LD";
            //    }
            //    else
            //    {
            //        sectionCalloutName += "RD";
            //    }
            //}

            //else if (crossUpTotal < -.1)
            //{
            //    sectionCalloutName += "H";
            //    foreach (SketchLine sketchline in sectionSketch.SketchLines)
            //    {
            //        if (sketchline.Geometry.Direction.X > .1 || sketchline.Geometry.Direction.X < -.1)
            //        {
            //            sectionLine = sketchline;
            //            break;
            //        }
            //    }
            //    if (sectionLine.Geometry.Direction.X > 0)
            //    {
            //        sectionCalloutName += "LU";
            //    }
            //    else
            //    {
            //        sectionCalloutName += "RU";
            //    }
            //}
            //else
            //{
            //    double crossTotal = crossVector.X + crossVector.Y + crossVector.Z;
            //    double parentUpTotal = parentUpVector.X + parentUpVector.Y + parentUpVector.Z;
            //    sectionCalloutName += "V";

            //    foreach (SketchLine sketchline in sectionSketch.SketchLines)
            //    {
            //        if (sketchline.Geometry.Direction.Y > .1 || sketchline.Geometry.Direction.Y < -.1)
            //        {
            //            sectionLine = sketchline;
            //            break;
            //        }
            //    }
            //    if (sectionLine.Geometry.Direction.Y < 0)
            //    {
            //        sectionCalloutName += "U";
            //    }
            //    else
            //    {
            //        sectionCalloutName += "D";
            //    }

            //    if ((crossTotal * parentUpTotal) > 0)
            //    {
            //        sectionCalloutName += "R";
            //    }
            //    else
            //    {
            //        sectionCalloutName += "L";
            //    }
            //}
            Point2d sectionLineSketchPoint, sectionLineSheetPoint;
            sectionLineSketchPoint = sectionLine.StartSketchPoint.Geometry;
            sectionLineSheetPoint = sectionSketch.SketchToSheetSpace(sectionLineSketchPoint);

            try
            {
                SketchedSymbolDefinition symbolDef = drawDoc.SketchedSymbolDefinitions[sectionCalloutName];
                int pageNumIndex = sectionSheet.Name.IndexOf(":") + 1;
                //string itemNum = GetStringProperty((Document)drawDoc, "Inventor User Defined Properties", "ITEM CODE");
                string pageNumber = sectionSheet.Name.Substring(pageNumIndex); //itemNum + "." + sectionSheet.Name.Substring(pageNumIndex);
                string[] symbolPromptString = new string[2] { addedSection.Name, pageNumber };
                ObjectCollection leaderPoints = AddinGlobal.InventorApp.TransientObjects.CreateObjectCollection();
                GeometryIntent geoIntent = oSheet.CreateGeometryIntent(sectionLine);
                leaderPoints.Add(sectionLineSheetPoint);
                leaderPoints.Add(geoIntent);
                SketchedSymbol sectionSymbol = oSheet.SketchedSymbols.AddWithLeader(symbolDef, leaderPoints, 0, 1, 
                    symbolPromptString); 
                byte[] calloutRefKey = new byte[10];
                sectionSymbol.LeaderVisible = false;
                sectionSymbol.SymbolClipping = false;
                sectionSymbol.LeaderClipping = false;
                sectionSymbol.GetReferenceKey(ref calloutRefKey);
                drawView.updateViewAttribute("ViewData", "CalloutRefKey", ValueTypeEnum.kByteArrayType,
                    calloutRefKey);
            }
            catch
            {
                MessageBox.Show("Could not find Sketched Symbol " + sectionCalloutName);
            }
            
        
        }
        
        public static void DeleteSectionViewSymbol(this SectionDrawingView sectionView)
        {
            Sheet oSheet = sectionView.Parent;
            DrawingDocument oDrawDoc = (DrawingDocument)oSheet.Parent;
            object matchType, outOb, outContext;
            if (sectionView.AttributeSets.NameIsUsed["ViewData"]
            && sectionView.AttributeSets["ViewData"].NameIsUsed["CalloutRefKey"])
            {
                byte[] calloutRefKey = (byte[])sectionView.AttributeSets["ViewData"]["CalloutRefKey"].Value;
                if (oDrawDoc.ReferenceKeyManager.CanBindKeyToObject(calloutRefKey, 0, out outOb, out outContext))
                {
                    object viewLabelObject = oDrawDoc.ReferenceKeyManager.BindKeyToObject(calloutRefKey,
                        0, out matchType);
                    SketchedSymbol calloutLabel = (SketchedSymbol)viewLabelObject;
                    calloutLabel.Delete();
                }
            }
        }

        public static Inventor.ObjectTypeEnum GetObjectType(this object obj)
        {
            Inventor.ObjectTypeEnum objectType = (Inventor.ObjectTypeEnum)GetPropertyValue(obj, "Type");
            return objectType;
        }

        public static object GetPropertyValue(object obj, string propertyName)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");
            try
            {
                return obj.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, obj, null);
            }
            catch (Exception)
            {
                return null;
            }
        }


        //public static RenderStyle getModelAppearance(this SketchedSymbol symbol, Inventor.Application app)
        //{
        //    Sheet oSheet = symbol.Parent;
        //    app = (Inventor.Application)oSheet.Application;
        //    LeaderNode endNode = symbol.Leader.AllNodes[symbol.Leader.AllNodes.Count];
        //    GeometryIntent geoIntent = endNode.AttachedEntity;
        //    DrawingCurve drawingCurve = null;
        //    Document modelDoc = (Document)oSheet.DrawingViews[1].ReferencedDocumentDescriptor.ReferencedDocument;
        //    RenderStyle appearance = null;
        //    int finishColorFaceInt = 3;
        //    Face finishFace = null;
        //    try
        //    {
        //        drawingCurve = (DrawingCurve)geoIntent.Geometry;
        //    }
        //    catch
        //    {
        //        return appearance;
        //    }
        //    object modelGeoObj = drawingCurve.ModelGeometry;
        //    StyleSourceTypeEnum styleSourceType;
        //    if (modelGeoObj is Face)
        //    {
        //        finishFace = (Face)modelGeoObj;
        //        //appearance = geoFace.GetRenderStyle(out styleSourceType);
        //    }
        //    else if(modelGeoObj is Edge)
        //    {
        //        Edge intentEdge = (Edge)modelGeoObj;
        //        Inventor.Color redColor = app.TransientObjects.CreateColor(255, 10, 50);
        //        Inventor.Color originalSymbolColor, originalEdgeColor;
        //        bool tagFaceSet = false;
        //        //Change this logic to ask user which appearance if 2 differents
        //        List<RenderStyle> edgeRenderStyles = new List<RenderStyle>();
        //        foreach (Face face in intentEdge.Faces)
        //        {
        //            edgeRenderStyles.Add(face.GetRenderStyle(out styleSourceType));
        //        }
        //        if (edgeRenderStyles.Distinct().Count() > 1)
        //        {
        //            if (symbol.AttributeSets.NameIsUsed["SymbolData"])
        //            {
        //                if (symbol.AttributeSets["SymbolData"].NameIsUsed["FaceForFinish"])
        //                {
        //                    tagFaceSet = true;
        //                    finishColorFaceInt = (int)symbol.AttributeSets["SymbolData"]["FaceForFinish"].Value;
        //                    finishFace = intentEdge.Faces[finishColorFaceInt];
        //                }
        //            }
        //            if (!tagFaceSet)
        //            {
        //                oSheet.Activate();
        //                originalSymbolColor = symbol.Color;
        //                symbol.Color = redColor;
        //                symbol.Scale = symbol.Scale * 2;
        //                originalEdgeColor = drawingCurve.Color;
        //                drawingCurve.Color = redColor;
        //                Camera activeCam = app.ActiveView.Camera;
        //                Inventor.Point eyePoint = app.TransientGeometry.CreatePoint(symbol.Position.X, symbol.Position.Y,
        //                    app.ActiveView.Camera.Eye.Z);
        //                Inventor.Point targetPoint = app.TransientGeometry.CreatePoint(symbol.Position.X, symbol.Position.Y,
        //                    app.ActiveView.Camera.Target.Z);
        //                activeCam.Eye = eyePoint;
        //                activeCam.Target = targetPoint;
        //                activeCam.SetExtents(oSheet.Width / 4, oSheet.Height / 4);
        //                activeCam.Apply();
        //                //SelectFinishForm selectFinForm = new SelectFinishForm(edgeRenderStyles);
        //                selectFinForm.ShowDialog();
        //                finishColorFaceInt = selectFinForm.selection + 1;
        //                finishFace = intentEdge.Faces[finishColorFaceInt];
        //                symbol.updateSymbolAttribute("SymbolData", "FaceForFinish", ValueTypeEnum.kIntegerType,
        //                    finishColorFaceInt);
        //                symbol.Color = originalSymbolColor;
        //                symbol.Scale = symbol.Scale / 2;
        //                drawingCurve.Color = originalEdgeColor;
        //                app.ActiveView.Fit();
        //            }
        //        }
        //        else
        //        {
        //            finishFace = intentEdge.Faces[1];
        //        }
        //    }
        //    if (finishFace != null)
        //    {
        //        appearance = finishFace.GetRenderStyle(out styleSourceType);
        //    }
        //    return appearance;
        //}

        public static void setFaceDataAttributes(this SurfaceBody body, Face partFace, Edge lengthEdge)
        {
            partFace.updateFaceAttribute("FaceData", "FaceName",
                            ValueTypeEnum.kStringType, "Face1");
            List<Edge> edges = new List<Edge>();
            foreach (Edge edge in partFace.Edges)
            {
                edges.Add(edge);
            }
            edges = edges.OrderBy(e => (e.length())).ToList();
            foreach (Face face in body.Faces)
            {
                if (face != partFace)
                {
                    if (face.IsParallelTo(partFace))
                    {
                        face.updateFaceAttribute("FaceData", "FaceName",
                            ValueTypeEnum.kStringType, "Face2");
                    }
                    else
                    {
                        if (edges.Count == 4)
                        {
                            foreach (Edge edge in face.Edges)
                            {
                                if (edge == edges[0])
                                {
                                    face.updateFaceAttribute("FaceData", "FaceName",
                                        ValueTypeEnum.kStringType, "WidthEdge2");
                                    break;
                                }
                                if (edge == edges[1])
                                {
                                    face.updateFaceAttribute("FaceData", "FaceName",
                                        ValueTypeEnum.kStringType, "WidthEdge1");
                                    break;
                                }
                                if (edge == edges[2])
                                {
                                    face.updateFaceAttribute("FaceData", "FaceName",
                                        ValueTypeEnum.kStringType, "LengthEdge2");
                                    break;
                                }
                                if (edge == edges[3])
                                {
                                    face.updateFaceAttribute("FaceData", "FaceName",
                                        ValueTypeEnum.kStringType, "LengthEdge1");
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (Edge edge in face.Edges)
                            {
                                if (edge == lengthEdge)
                                {
                                    face.updateFaceAttribute("FaceData", "FaceName",
                                        ValueTypeEnum.kStringType, "LengthEdge1");
                                    break;
                                }
                                
                            }
                        }
                    }
                }
            }
        }


        private static string GetStringProperty(Document document, string propSetName, string propName)
        {
            object propSet;
            try
            {
                if (document.PropertySets.PropertySetExists(propSetName, out propSet))
                {
                    return (string) ((PropertySet) propSet)[propName].Value;
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
    }


    public class HCM_Finish
    {
        private string fullName;
        private string calloutName;
        private string description1;
        private string description2;
        private string description3;

        public string FullName
        {
            get
            {
                return fullName;
            }
            set
            {
                fullName = value;
            }
        }

        public string CalloutName
        {
            get
            {
                return calloutName;
            }
        }
        public string Description1
        {
            get
            {
                return description1;
            }
            set
            {
                description1 = value;
            }
        }

        public string Description2
        {
            get
            {
                return description2;
            }
            set
            {
                description2 = value;
            }
        }

        public string Description3
        {
            get
            {
                return description3;
            }
            set
            {
                description3 = value;
            }
        }


        private void setFullName(string fName)
        {
            fullName = fName;
            setcalloutName();
        }

        private void setcalloutName()
        {
            if (fullName.Length > 4)
            {
                calloutName = fullName.Substring(4);
            }
        }

        public void setdescription1(string description)
        {
            description1 = description;
        }

        public void setdescription2(string description)
        {
            description2 = description;
        }

        public void setdescription3(string description)
        {
            description3 = description;
        }

        public HCM_Finish() //blank constructor
        {
            
        }
        public HCM_Finish(string FullName) //blank constructor
        {
            setFullName(FullName);
        }
        public HCM_Finish(string FullName, string Description1, string Description2, string Description3) //blank constructor
        {
            setFullName(FullName);
            setdescription1(Description1);
            setdescription2(Description2);
            setdescription3(Description3);
        }

    }


}
