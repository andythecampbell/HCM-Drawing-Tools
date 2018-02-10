using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Timers;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;

namespace HCMToolsInventorAddIn
{
    class PartOptions
    {
        public string edgeBandFeatureName = "Removed For EdgeBand";
        public void CutSizePartsOn(Document oDoc)
        {
            Inventor.Application app = AddinGlobal.InventorApp;
            TransactionManager transMan = app.TransactionManager;
            PartDocument oPartDoc;
            List<PartDocument> allPartDocs = new List<PartDocument>();
            Transaction trans = transMan.StartTransactionForDocumentOpen("Show Cut Size");
            if (oDoc is PartDocument)
            {
                oPartDoc = (PartDocument)oDoc;
                removeBanding(oPartDoc);
            }
            else if (oDoc is AssemblyDocument)
            {
                AssemblyDocument oAsemDoc = (AssemblyDocument)oDoc;
                allPartDocs = oAsemDoc.ReferencedNestedPartDocuements();
                foreach(PartDocument oSinglePartDoc in allPartDocs)
                {
                    removeBanding(oSinglePartDoc);
                }
            }
            trans.End();
        }

        public void CutSizePartsOff(Document oDoc)
        {
            Inventor.Application app = AddinGlobal.InventorApp;
            TransactionManager transMan = app.TransactionManager;
            PartDocument oPartDoc;
            List<PartDocument> allPartDocs = new List<PartDocument>();
            Transaction trans = transMan.StartTransactionForDocumentOpen("Show Full Size");
            if (oDoc is PartDocument)
            {
                oPartDoc = (PartDocument)oDoc;
                supressBanding(oPartDoc);
            }
            else if (oDoc is AssemblyDocument)
            {
                AssemblyDocument oAsemDoc = (AssemblyDocument)oDoc;
                allPartDocs = oAsemDoc.ReferencedNestedPartDocuements();
                foreach (PartDocument oSinglePartDoc in allPartDocs)
                {
                    supressBanding(oSinglePartDoc);
                }
            }
            trans.End();
        }

        private void removeBanding(PartDocument oPartDoc)
        {
            foreach (ThickenFeature thicken in oPartDoc.ComponentDefinition.Features.ThickenFeatures)
            {
                if (thicken.Name.StartsWith(edgeBandFeatureName))
                {
                    string distance = thicken.Distance.Expression;
                    if (distance.Contains(":"))
                    {
                        distance = distance.Substring(0, distance.LastIndexOf(":"));
                    }
                    RenderStyle oRenderStyle = null;
                    try
                    {
                        oRenderStyle = oPartDoc.RenderStyles[distance];
                    }
                    catch
                    {
                        MessageBox.Show("Could not find appearance " + distance);
                        continue;
                    }
                    if(thicken.InputFaces is FaceCollection)
                    {
                        FaceCollection oInputFaces = (FaceCollection)thicken.InputFaces;
                        foreach (object bandedFaceOB in oInputFaces)
                        {
                            Face bandedFace = null;
                            if (bandedFaceOB is Face)
                            {
                                bandedFace = (Face)bandedFaceOB;
                            }
                            else
                            {
                                continue;
                            }
                            try
                            {
                                StyleSourceTypeEnum styleSource;
                                bandedFace.GetRenderStyle(out styleSource);
                                if (styleSource == StyleSourceTypeEnum.kOverrideRenderStyle)
                                {
                                    byte[] refKey = new byte[10];
                                    bandedFace.GetRenderStyle(out styleSource).GetReferenceKey(refKey, 0);
                                    bandedFace.updateFaceAttribute("FaceData", "NominalAppearance", ValueTypeEnum.kByteArrayType, refKey);
                                }
                                bandedFace.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, oRenderStyle);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Unable to apply appearance in " + oPartDoc.DisplayName + "\n" + ex.Message);
                                ErrorLog.LogError(ex, (Document)oPartDoc, "Set Renderstyle in Remove Banding");
                            }
                        }
                    }
                    thicken.Suppressed = false;
                }
            }
        }

        private void supressBanding(PartDocument oPartDoc)
        {
            foreach (ThickenFeature thicken in oPartDoc.ComponentDefinition.Features.ThickenFeatures)
            {
                if (thicken.Name.StartsWith(edgeBandFeatureName))
                {
                    thicken.Suppressed = true;
                    RenderStyle partStyle = null;
                    
                    if (thicken.InputFaces is FaceCollection)
                    {
                        foreach (Face oFace in (FaceCollection)thicken.InputFaces)
                        {
                            if (findNominalAppearance(oFace) != null)
                            {
                                partStyle = findNominalAppearance(oFace);
                                try
                                {
                                    oFace.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, partStyle);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Unable to apply previous apperance");
                                    ErrorLog.LogError(ex, (Document)oPartDoc, "suppressBanding");
                                }
                            }
                            else
                            {
                                partStyle = oPartDoc.ActiveRenderStyle;
                            }
                            
                        }
                    }
                }
            }
        }

        private RenderStyle findNominalAppearance(Face inputFace)
        {
            RenderStyle nominalAppearance = null;
            Document doc = (Document)inputFace.Parent.ComponentDefinition.Document;
            if (inputFace.AttributeSets.NameIsUsed["FaceData"] &&
                inputFace.AttributeSets["FaceData"].NameIsUsed["NominalAppearance"])
            {
                try
                {
                    byte[] refKey = new byte[10];
                    object matchType;
                    refKey = (byte[])inputFace.AttributeSets["FaceData"]["NominalAppearance"].Value;
                    object styleOb = doc.ReferenceKeyManager.BindKeyToObject(ref refKey, 0, out matchType);
                    if (styleOb is RenderStyle)
                    {
                        nominalAppearance = (RenderStyle)styleOb;
                    }
                }
                catch (Exception ex)
                {
                    ErrorLog.LogError(ex, doc, "FindNominalAppearance");
                }
            }
            return nominalAppearance;
        }
    }
}


