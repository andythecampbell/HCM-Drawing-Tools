using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Inventor;

namespace HCMToolsInventorAddIn
{
    //This 
    class RotateGrainButton : Button
    {
        public RotateGrainButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId,
            string description, string toolTip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, toolTip, standardIcon, largeIcon, buttonDisplayType)
        {
            //intentionally blank
        }

        public RotateGrainButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId,
            string description, string toolTip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, toolTip, buttonDisplayType)
        {
            //intentionally blank
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            List<RotatedRenderstlye> rotatedAppearances = new List<RotatedRenderstlye>();
            try
            {
                Inventor.Application app = AddinGlobal.InventorApp;
                Document oDoc = app.ActiveEditDocument;
                SelectSet selected =  oDoc.SelectSet;
                RenderStyle oldAppearance;
                StyleSourceTypeEnum styleSource;
                bool rotatedExists = false;
                RotatedRenderstlye rotatedStyle = null;
                PartDocument partDoc;
                List<Face> facesToRotate = new List<Face>();

                foreach (object ob in selected)
                {
                    if (!(ob is Face) || ob is Edge)
                    {
                        MessageBox.Show("Selected Item Not A Face");
                        selected.Remove(ob);
                    }
                }
                if (selected.Count == 0)
                {
                    MessageBox.Show("No Faces Selected");
                    return;
                }
                
                foreach (Face loopVar in selected)
                {
                    if (loopVar is FaceProxy) //checks for the selected face is a proxy used for in place edit in assembly
                    {
                        FaceProxy faceProxy = (FaceProxy)loopVar;
                        facesToRotate.Add(faceProxy.NativeObject);
                    }
                    else
                    {
                        facesToRotate.Add(loopVar);
                    }
                }


                foreach (Face oFace in facesToRotate)
                {
                    if (oFace.Parent.ComponentDefinition.Document is PartDocument)
                    {
                        partDoc = (PartDocument)oFace.Parent.ComponentDefinition.Document;
                    }
                    else
                    {
                        MessageBox.Show("Face Parent Not Part Doc");
                        continue;
                    }
                    oldAppearance = oFace.GetRenderStyle(out styleSource);
                    foreach (RotatedRenderstlye rotatedAppearance in rotatedAppearances)
                    {
                        if (rotatedAppearance.originalRenderStyle.Name == oldAppearance.Name
                            && oldAppearance.Parent == rotatedAppearance.oDocument)
                        {
                            rotatedExists = true;
                            rotatedStyle = rotatedAppearance;
                        }
                        
                    }
                    if (!rotatedExists)
                        {
                            rotatedStyle = new RotatedRenderstlye(partDoc, oldAppearance);
                            rotatedAppearances.Add(rotatedStyle);
                        }
                    oFace.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, rotatedStyle.oRenderstyle);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                //ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, "Grain Rotate");
            }
        }
        
    }

    class RotatedRenderstlye 
    {
        public RenderStyle oRenderstyle { get; set;}
        public Document oDocument{ get; set;}
        public RenderStyle originalRenderStyle { get; set; }

        public RotatedRenderstlye(PartDocument partDocument, RenderStyle originalRenderStyle)
        {
            oRenderstyle = copyAndRotateTexture(originalRenderStyle);
        }
        
        private RenderStyle copyAndRotateTexture(RenderStyle oAppearance)
        {
            
            RenderStyle newAppearance = oAppearance.Copy(oAppearance.Name + " (Rotated)");
            try
            {
                if (newAppearance.TextureRotation == 90)
                {
                    newAppearance.TextureRotation = 0;
                }
                else
                {
                    newAppearance.TextureRotation = 90;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Rotate\nDoes the Appearance Have a Texture?");
                //ErrorLog.LogError(ex, AddinGlobal.InventorApp.ActiveDocument, " Grain Rotate ");
            }
            return newAppearance;
        }
    }
}
