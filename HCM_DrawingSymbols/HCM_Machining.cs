
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;

using Inventor;

namespace HCMToolsInventorAddIn
{
    class HCM_Machining
    {
        private Inventor.Application _invApp;
        private Inventor.Document oDoc;

        public void MachineParts()
        {

             Inventor.AssemblyDocument asmDoc;
            string PartDisplayName;

            Inventor.PartDocument oPrtDoc;
            ObjectCollection AllMachineParts;
            int DocCount;
            int MachinePartsCount;
            Object MachinePartSurfaceBodyProxyObject;

            //Sets the running Inventor Application
            _invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            Inventor.Document oDoc = _invApp.ActiveDocument;
            AllMachineParts = _invApp.TransientObjects.CreateObjectCollection();
            AllMachineParts.Clear();


            DocCount = 0;
            MachinePartsCount = 0;

            //If the active document is an AssemblyDocument, Delete all previous "Machined" CombineFeatures and "Machined-Parts" Bodies
            //Create two lists: 1 for all machine- parts and 1 for all other parts (excluding hardware parts)
            if (oDoc is AssemblyDocument)
            {
                asmDoc = (AssemblyDocument)oDoc;
                string asmFullFileName = asmDoc.FullFileName;
                List<PartDocument> allPartDocs = asmDoc.ReferencedNestedPartDocuements();
                List<PartDocument> referencedPartDocs = asmDoc.ReferencedNestedPartDocuements();
                List<PartDocument> machineParts = asmDoc.ReferencedNestedPartDocuements();
                int allPartsDocsCount = referencedPartDocs.Count;

                for (int i = 0; i < allPartsDocsCount;i++)
                {
                    oPrtDoc = (PartDocument) allPartDocs[i];
                    PartDisplayName = oPrtDoc.DisplayName;
                    PropertySet propertyset = oPrtDoc.PropertySets["Inventor Document Summary Information"];
                    Property Category = oPrtDoc.PropertySets["Inventor Document Summary Information"].ItemByPropId[(int)PropertiesForDocSummaryInformationEnum.kCategoryDocSummaryInformation];


                    //If the part document Display name has more than 8 characters, check if the first 8 are "Machine-", if so, remove it from the referenced parts list
                    if (PartDisplayName.Count() >= 8)
                    {
                        if (PartDisplayName.Substring(0, 8).Equals("Machine-"))
                        {
                            referencedPartDocs.Remove(oPrtDoc);
                            MachinePartsCount++;
                        }
                        else
                        {
                            if (Category.Value.Equals("H"))
                            {
                                referencedPartDocs.Remove(oPrtDoc);
                                machineParts.Remove(oPrtDoc);
                            }
                            else
                            {
                                machineParts.Remove(oPrtDoc);
                                DocCount++;
                            }
                        }
                    }
                    else 
                    {
                        if(Category.Equals("H"))
                        {
                            referencedPartDocs.Remove(oPrtDoc);
                            machineParts.Remove(oPrtDoc);
                        }
                        else
                        {
                            machineParts.Remove(oPrtDoc);
                            DocCount++;
                        }
                    }
                }

                //Delete previous "Machined-" CombinedFeatures and "Machine-Parts" Bodies in referencedPartDocs
                foreach (PartDocument refPrtDoc in referencedPartDocs)
                {
                    foreach (CombineFeature combFeature in refPrtDoc.ComponentDefinition.Features.CombineFeatures)
                    {
                        if (combFeature.Name.Equals("Machined"))
                        {
                            combFeature.Delete();
                        }
                    }
                    foreach (NonParametricBaseFeature machineBodies in refPrtDoc.ComponentDefinition.Features.NonParametricBaseFeatures)
                    {
                        if (machineBodies.Name.Equals("Machined-Parts"))
                        {
                            machineBodies.Delete();
                        }
                    }
                }

                //If there is at least 1 "Machine-" part, Create a Derived Part, use all Machine- Parts as Solids, exclude every other part
                if (MachinePartsCount > 0)
                {
                    //Create a new part document for the machine- parts
                    PartDocument DerivedPartDocument = (PartDocument)_invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, _invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), false);
                    String DerivedPartDocumentFullFileName = asmDoc.FullFileName.Substring(0, asmDoc.FullFileName.LastIndexOf('\\')) + asmFullFileName.Substring(asmFullFileName.LastIndexOf("\\")) + "pDerived_Machine_Parts.ipt";
                    if (System.IO.File.Exists(DerivedPartDocumentFullFileName))
                    {
                        System.IO.File.Delete(DerivedPartDocumentFullFileName);
                    }
                    PartComponentDefinition DerivedPrtDocDef = (PartComponentDefinition)DerivedPartDocument.ComponentDefinition;
                    DerivedAssemblyDefinition DerivedAsmDef = (DerivedAssemblyDefinition)DerivedPrtDocDef.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(asmDoc.FullFileName);

                    
                    for (int i = 1; i < DerivedAsmDef.Occurrences.Count+1; i++)
                    {
                        int oCompOccNameCount = DerivedAsmDef.Occurrences[i].Name.Count();

                        if (oCompOccNameCount >= 8)
                        {
                            if (DerivedAsmDef.Occurrences[i].Name.Substring(0, 8).Equals("Machine-"))
                            {
                                DerivedAsmDef.Occurrences[i].InclusionOption = DerivedComponentOptionEnum.kDerivedIncludeAll;

                            }
                            else
                            {
                                DerivedAsmDef.Occurrences[i].InclusionOption= DerivedComponentOptionEnum.kDerivedExcludeAll;
                            }
                        }
                        else
                        {
                            DerivedAsmDef.Occurrences[i].InclusionOption = DerivedComponentOptionEnum.kDerivedExcludeAll;
                        }
                    }

                    DerivedAsmDef.InclusionOption = DerivedComponentOptionEnum.kDerivedIndividualDefined;
                    DerivedAsmDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams;
                    DerivedAsmDef.IncludeAllTopLeveliMateDefinitions= DerivedComponentOptionEnum.kDerivedExcludeAll;
                    DerivedAsmDef.IncludeAllTopLevelParameters = DerivedComponentOptionEnum.kDerivedExcludeAll;
                    DerivedAsmDef.IncludeAllTopLevelSketches = DerivedComponentOptionEnum.kDerivedExcludeAll;
                    DerivedAsmDef.ReducedMemoryMode = true;
                    DerivedAsmDef.SetHolePatchingOptions(DerivedHolePatchEnum.kDerivedPatchNone);

                    //'Creates a part file including all "Machine-" Parts as a 1 solid body file, Breaks the link to the referenced assembly then saves a copy (Copy will be deleted at the end)
                    DerivedAssemblyComponent DerivedAsmDoc = DerivedPrtDocDef.ReferenceComponents.DerivedAssemblyComponents.Add(DerivedAsmDef);
                    DerivedAsmDoc.BreakLinkToFile();
                    _invApp.SilentOperation.Equals(true);
                    DerivedPartDocument.SaveAs(DerivedPartDocumentFullFileName, false);
                    _invApp.SilentOperation.Equals(false);


                    Inventor.TransientGeometry OTG = _invApp.TransientGeometry;
                    Inventor.Matrix MachinePartsPosition = OTG.CreateMatrix();
                    asmDoc.ComponentDefinition.Occurrences.Add(DerivedPartDocumentFullFileName, MachinePartsPosition);

                    foreach (ComponentOccurrence oCompOcc in asmDoc.ComponentDefinition.Occurrences)
                    {
                        if (oCompOcc.Name.Equals(DerivedPartDocument.DisplayName + ":1"))
                        {
                            oCompOcc.SurfaceBodies[1].Name = "Machine-Parts";
                            Inventor.SurfaceBody MachinePartsSolid = oCompOcc.SurfaceBodies[1];
                            Inventor.PartComponentDefinition MachinePartDef = (PartComponentDefinition)oCompOcc.Definition;
                            oCompOcc.CreateGeometryProxy(MachinePartsSolid, out MachinePartSurfaceBodyProxyObject);
                            AllMachineParts.Add(MachinePartSurfaceBodyProxyObject);
                        }
                    }

                    referencedPartDocs.Remove(DerivedPartDocument);
                    CopyBody(asmDoc, referencedPartDocs, DerivedPartDocument, AllMachineParts);

                    asmDoc.ComponentDefinition.Occurrences.ItemByName[DerivedPartDocument.DisplayName + ":1"].Delete();
                   
                    foreach (Document Doc in _invApp.Documents)
                    {
                        if(Doc.DisplayName.Equals(DerivedPartDocument.DisplayName))
                        {
                            Doc.Close();
                        }
                    }
                    System.IO.File.Delete(DerivedPartDocumentFullFileName);

                    MessageBox.Show("Machining Complete");
                }
                else if (oDoc is PartDocument)
                {
                    MessageBox.Show("Active document must be an assembly document");

                }
            }
        }

        private void CopyBodies(Inventor.ComponentOccurrences oCompOccs, Inventor.PartDocument DerivedPartDocument, ObjectCollection AllMachineParts)
        {
            Inventor.PartComponentDefinition oPartDef;
            Inventor.NonParametricBaseFeatureDefinition CopiedFeatureDef;
            Inventor.NonParametricBaseFeature CopiedFeature;
            Inventor.ComponentOccurrences oCompSubOccs;


            foreach (ComponentOccurrence oCompOcc in oCompOccs)
            {
                if (oCompOcc.DefinitionDocumentType.Equals(DocumentTypeEnum.kAssemblyDocumentObject))
                {
                    oCompSubOccs = (ComponentOccurrences)oCompOcc.SubOccurrences;
                    CopyBodies(oCompSubOccs, DerivedPartDocument, AllMachineParts);
                }
                else if (oCompOcc.DefinitionDocumentType.Equals(DocumentTypeEnum.kPartDocumentObject))
                {
                    int oCompOccName = (int)oCompOcc.Name.Count();
                    if (oCompOccName >= 8)
                    {
                        if (oCompOcc.Name.Substring(0, 8).Equals("Machine-"))
                        {
                        }
                        else
                        {

                            if (oCompOcc.Name.Equals(DerivedPartDocument.FullFileName + ":1"))
                            {
                            }
                            else if (oCompOcc.Name.Substring(oCompOcc.Name.Count() - 2, 2).Equals(":1"))
                            {
                                oPartDef = (Inventor.PartComponentDefinition)oCompOcc.Definition;
                                CopiedFeatureDef = oPartDef.Features.NonParametricBaseFeatures.CreateDefinition();
                                CopiedFeatureDef.BRepEntities = AllMachineParts;
                                CopiedFeatureDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType;
                                CopiedFeatureDef.TargetOccurrence = oCompOcc;
                                CopiedFeatureDef.IsAssociative = false;
                                CopiedFeature = oPartDef.Features.NonParametricBaseFeatures.AddByDefinition(CopiedFeatureDef);
                                oCompOcc.SurfaceBodies[oCompOcc.SurfaceBodies.Count].Name = "Machined-Parts";
                            }
                            else
                            {
                            }
                        }
                    }
                    else
                    {
                        if (oCompOcc.Name.Substring(oCompOcc.Name.Count() - 2, 2).Equals(":1"))
                        {
                            oPartDef = (Inventor.PartComponentDefinition)oCompOcc.Definition;
                            CopiedFeatureDef = oPartDef.Features.NonParametricBaseFeatures.CreateDefinition();
                            CopiedFeatureDef.BRepEntities = AllMachineParts;
                            CopiedFeatureDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType;
                            CopiedFeatureDef.TargetOccurrence = oCompOcc;
                            CopiedFeatureDef.IsAssociative = false;
                            CopiedFeature = oPartDef.Features.NonParametricBaseFeatures.AddByDefinition(CopiedFeatureDef);
                            oCompOcc.SurfaceBodies[oCompOcc.SurfaceBodies.Count].Name = "Machined-Parts";
                        }
                    }

                }

            }

        }
        private void CopyBody(Inventor.AssemblyDocument asmDoc, List<PartDocument> referencedPartDocs, Inventor.PartDocument DerivedPartDocument, ObjectCollection AllMachineParts)
        {
            Inventor.PartComponentDefinition oPartDef;
            Inventor.NonParametricBaseFeatureDefinition CopiedFeatureDef;
            Inventor.NonParametricBaseFeature CopiedFeature;
            Inventor.ObjectCollection Machining;

            Machining = _invApp.TransientObjects.CreateObjectCollection();
            Machining.Clear();


            for (int i = 0; i < referencedPartDocs.Count(); i++)
            {
                PartDocument oPartDoc = (PartDocument)referencedPartDocs[i];
                Document oDoc = (Document) referencedPartDocs[i];
                Inventor.ComponentOccurrence oPrtOcc = (ComponentOccurrence) asmDoc.ComponentDefinition.Occurrences.ItemByName[oDoc.DisplayName+":1"];
                oPartDef = (Inventor.PartComponentDefinition)oPrtOcc.Definition;
                CopiedFeatureDef = oPartDef.Features.NonParametricBaseFeatures.CreateDefinition();
                CopiedFeatureDef.BRepEntities = AllMachineParts;
                CopiedFeatureDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType;
                CopiedFeatureDef.TargetOccurrence = oPrtOcc;
                CopiedFeatureDef.IsAssociative = false;
                CopiedFeature = oPartDef.Features.NonParametricBaseFeatures.AddByDefinition(CopiedFeatureDef);
                //oPrtOcc.SurfaceBodies[oPrtOcc.SurfaceBodies.Count].Name = "Machined-Parts";

                Machining.Clear();
                Machining.Add(oPartDoc.ComponentDefinition.SurfaceBodies[oPartDoc.ComponentDefinition.SurfaceBodies.Count]);
                oPartDoc.ComponentDefinition.Features.CombineFeatures.Add(oPartDoc.ComponentDefinition.SurfaceBodies[1], Machining, PartFeatureOperationEnum.kCutOperation);
                oPartDoc.ComponentDefinition.Features.CombineFeatures[oPartDoc.ComponentDefinition.SurfaceBodies.Count].Name = "Machined";
                oPartDoc.ComponentDefinition.Features.NonParametricBaseFeatures[oPartDoc.ComponentDefinition.Features.NonParametricBaseFeatures.Count].Name = "Machined-Parts";
        
            }      
        }      
    }   
}

