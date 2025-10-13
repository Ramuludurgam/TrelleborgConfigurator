using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

namespace Configurator
{
    public class SupportBracket(SldWorks swApp, string destPath)
    {
        private SldWorks swApp = swApp;
        private string destinationPath = destPath;

        private Feature swCavityFeat;

        public void CreateSupportBracket()
        {
            ModelDoc2 swModelDoc;
            SketchSegment swSketchSeg;
            SketchManager swSketchMgr;
            Sketch swMySketch;
            FeatureManager swFeatMgr;
            Feature swFeat;
            Feature swChamFeat;


            int i;
            object vArrSketchSegs;

            //const double Width = 800;        // in mm
            //const double Height = 800;       // in mm
            const double Thickness = 25;     // in mm
                                             //const double FilletRadius = 35;  // in mm
            const double HoleDiameter = 30;  // in mm
            const double mmToMeter = 0.001;  // convert to meters
            const double chamferLength = 30; // mm

            bool boolStatus;

            swModelDoc = (ModelDoc2)swApp.NewPart();
            swSketchMgr = swModelDoc.SketchManager;
            swFeatMgr = swModelDoc.FeatureManager;

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // **** First Extrusion ***
            boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);
                vArrSketchSegs = swSketchMgr.CreateCornerRectangle(-0.475, -0.165, 0, 0, 0.165, 0);

                object[] arr = (object[])vArrSketchSegs;

                for (i = 0; i <= arr.GetUpperBound(0); i++)
                {
                    swSketchSeg = (SketchSegment)arr[i];

                    if (swSketchSeg.GetName() == "Line4" || swSketchSeg.GetName() == "Line3")
                    {
                        swSketchSeg.Select2(false, -1);
                        swModelDoc.AddDimension2(0, 0, 0);
                    }
                }

                swMySketch = (Sketch)swSketchMgr.ActiveSketch;
                swModelDoc.InsertSketch2(false);

                object[] regions = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)regions[0]).Select2(true, null);

                swFeat = swFeatMgr.FeatureExtrusion3(
                    true, false, true,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    Thickness * mmToMeter, 0,
                    false, false, false, false, 0, 0, false,
                    false, false, false, false, false, true,
                    (int)swStartConditions_e.swStartSketchPlane, 0, false);
            }

            // *** Second extrusion ***
            boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);
                vArrSketchSegs = swSketchMgr.CreateCornerRectangle(-0.475, -0.092, 0, -0.155, 0.092, 0);

                object[] arr2 = (object[])vArrSketchSegs;

                for (i = 0; i <= arr2.GetUpperBound(0); i++)
                {
                    swSketchSeg = (SketchSegment)arr2[i];

                    if (swSketchSeg.GetName() == "Line4" || swSketchSeg.GetName() == "Line3")
                    {
                        swSketchSeg.Select2(false, -1);
                        swModelDoc.AddDimension2(0, 0, 0);
                    }
                }

                swMySketch = (Sketch)swSketchMgr.ActiveSketch;
                object[] regions2 = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)regions2[0]).Select2(true, null);

                this.swCavityFeat = (Feature)swFeatMgr.FeatureCut4(
                    false, false, false,
                    (int)swEndConditions_e.swEndCondUpToNext, 1,
                    0.01, 0.01, false, false, false, false,
                    0, 0, false, false, false, false, false,
                    true, true, true, true, false, 0, 0, false, false);
            }

            // **** Hole Cut Extrude ***
            boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);
                swSketchSeg = swSketchMgr.CreateCircleByRadius(-0.375, -0.165, 0, HoleDiameter * mmToMeter / 2);
                swMySketch = (Sketch)swSketchMgr.ActiveSketch;
                swModelDoc.InsertSketch2(false);

                object[] regions3 = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)regions3[0]).Select2(true, null);

                Feature swHoleFeature = swFeatMgr.FeatureCut4(
                    false, false, false,
                    (int)swEndConditions_e.swEndCondUpToNext, 1,
                    0.01, 0.01, false, false, false, false,
                    0, 0, false, false, false, false, false,
                    true, true, true, true, false, 0, 0, false, false);
            }

            CreateFilletFeature(swModelDoc, 20 * mmToMeter);
            swChamFeat = CreateChamferFeature(swModelDoc, chamferLength * mmToMeter);
            CreateMirrorFeature(swModelDoc, swChamFeat);

            boolStatus = swModelDoc.Extension.SelectByRay(-1.68495060847818E-04, 4.33753608115239E-03, 1.68495060790974E-04,
                -0.577381545199981, -0.577287712085548, -0.577381545199979,
                1.61352307744054E-03, 1, false, 0, 0);

            boolStatus = swModelDoc.InsertAxis2(true);

            // **** Add new configuration called Right ****
            ConfigurationManager swConfigMgr = swModelDoc.ConfigurationManager;
            Configuration swConfig = swConfigMgr.AddConfiguration2("Right", "Cavity length increased", "None",
                (int)swConfigurationOptions2_e.swConfigOption_UseAlternateName, "", "Right", true);

            swModelDoc.ShowConfiguration2("Default");

            Dimension swDim;
            DisplayDimension swDispDim;

            this.swCavityFeat = (Feature)this.swCavityFeat.GetFirstSubFeature();
            swDispDim = (DisplayDimension)this.swCavityFeat.GetFirstDisplayDimension();

            while (swDispDim != null)
            {
                swDim = swDispDim.GetDimension2(0);
                swDim.SetValue3(220, (int)swInConfigurationOpts_e.swThisConfiguration, "");
                swDispDim = (DisplayDimension)swDispDim.GetNext();
            }

            swModelDoc.ShowConfiguration2("Right");
            swModelDoc.ShowConfiguration2("Default");
            swModelDoc.EditRebuild3();

            // save the file to local folder
            string fileName = "SupportBracket.SLDPRT";
            this.SavePartFile(swModelDoc, fileName);
        }

        private void CreateFilletFeature(ModelDoc2 swModelDoc, double FilletRadius)
        {
            bool boolStatus;
            swModelDoc.ClearSelection2(true);
            swModelDoc.ShowNamedView2("*Trimetric", 8);
            swModelDoc.ViewZoomtofit2();

            boolStatus = swModelDoc.Extension.SelectByRay(-0.0113804809690237, 0.0922726656970099, 0.154727378614893,
                0, -0.010110466, -0.386910217, 0.00161352307744054, 1, false, 1, 0);
            boolStatus = swModelDoc.Extension.SelectByRay(-0.0146505021914436, -0.0918130126212873, 0.154813043009142,
                0, -0.010110466, -0.386910217, 0.00161352307744054, 1, true, 1, 0);

            SimpleFilletFeatureData2 swSimpFillFeatData = (SimpleFilletFeatureData2)swModelDoc.FeatureManager.CreateDefinition((int)swFeatureNameID_e.swFmFillet);
            swSimpFillFeatData.Initialize((int)swSimpleFilletType_e.swConstRadiusFillet);

            SelectionMgr swSelMgr = (SelectionMgr)swModelDoc.SelectionManager;
            Edge[] edgesToFillet = new Edge[swSelMgr.GetSelectedObjectCount()];

            for (int i = 1; i <= swSelMgr.GetSelectedObjectCount(); i++)
            {
                edgesToFillet[i - 1] = (Edge)swSelMgr.GetSelectedObject6(i, -1);
            }

            swSimpFillFeatData.Edges = edgesToFillet;
            swSimpFillFeatData.AsymmetricFillet = false;
            swSimpFillFeatData.DefaultRadius = FilletRadius;
            swSimpFillFeatData.ConicTypeForCrossSectionProfile = (int)swFeatureFilletOptions_e.swFeatureFilletUniformRadius;
            swSimpFillFeatData.CurvatureContinuous = false;
            swSimpFillFeatData.ConstantWidth = true;
            swSimpFillFeatData.IsMultipleRadius = false;
            swSimpFillFeatData.OverflowType = (int)swFilletOverFlowType_e.swFilletOverFlowType_Default;

            swModelDoc.FeatureManager.CreateFeature(swSimpFillFeatData);
        }

        private Feature CreateChamferFeature(ModelDoc2 swModelDoc, double chamferLength)
        {
            SketchManager swSketchMgr = swModelDoc.SketchManager;
            FeatureManager swFeatMgr = swModelDoc.FeatureManager;
            bool boolStatus;

            boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

                swSketchMgr.CreateLine(-0.395, -0.092, 0, -0.475, -0.138188022, 0);
                swSketchMgr.CreateLine(-0.475, -0.138188022, 0, -0.475, -0.092, 0);
                swSketchMgr.CreateLine(-0.475, -0.092, 0, -0.395, -0.092, 0);
                Sketch swMySketch = (Sketch)swSketchMgr.ActiveSketch;

                swModelDoc.InsertSketch2(false);

                //Sketch swMySketch = (Sketch)swSketchMgr.ActiveSketch;
                object[] regions = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)regions[0]).Select2(true, null);

                Feature swChamFeat = swFeatMgr.FeatureCut4(
                    false, false, false, (int)swEndConditions_e.swEndCondThroughAll, 1,
                    0.01, 0.01, false, false, false, false,
                    0, 0, false, false, false, false, false,
                    true, true, true, true, false, 0, 0, false, false);

                return swChamFeat;
            }

            return null;
        }

        private Feature CreateMirrorFeature(ModelDoc2 swModelDoc, Feature swChamferFeat)
        {
            FeatureManager swFeatMgr = swModelDoc.FeatureManager;
            bool boolStatus;

            swModelDoc.ClearSelection2(true);
            boolStatus = swModelDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 2, null, 0);
            boolStatus = swModelDoc.Extension.SelectByID2(swChamferFeat.Name, "BODYFEATURE", 0, 0, 0, true, 1, null, 0);

            Feature swMirrFeat = swFeatMgr.InsertMirrorFeature(false, false, false, false);
            return swMirrFeat;
        }

        private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
        {
     
            string fullPath = Path.Combine(destinationPath, fileName);
            swModelDoc.SaveAsSilent(fullPath, false);
        }
    }

}