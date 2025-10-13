using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swmotionstudy;
using SolidWorks.Interop.swcommands;
using System.IO;

namespace Configurator
{
    public class BasePlate(SldWorks swApp, string destPath)
    {
        private SldWorks swApp = swApp;
        private string destinationPath = destPath;
        public void CreateBasePlate()
        {
            ModelDoc2 swModelDoc;
            SketchSegment swSketchSeg;
            SketchManager swSketchMgr;
            Sketch swMySketch;
            PartDoc swPartDoc;

            FeatureManager swFeatMgr;
            Feature swFeat;

            int i;
            object[] vArrSketchSegs = null;

            // Constants
            const double Width = 800;         // in mm
            const double Height = 800;        // in mm
            const double Thickness = 35;      // in mm
            const double FilletRadius = 35;   // in mm
            const double HoleDiameter = 52;   // in mm
            const double mmToMeter = 0.001;   // convert to meters

            bool boolStatus;

            // Get SolidWorks Application
            swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

            // Create a new part document
            swModelDoc = (ModelDoc2)swApp.NewPart();
            swSketchMgr = swModelDoc.SketchManager;
            swFeatMgr = swModelDoc.FeatureManager;

            // Select Front Plane
            boolStatus = swModelDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);

                // Create center rectangle
                vArrSketchSegs = (object[])swSketchMgr.CreateCenterRectangle(0, 0, 0, Width * mmToMeter / 2, Height * mmToMeter / 2, 0);

                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

                // Add dimensions to Line1 and Line2
                for (i = 0; i <= vArrSketchSegs.GetUpperBound(0); i++)
                {
                    swSketchSeg = (SketchSegment)vArrSketchSegs[i];

                    if (swSketchSeg.GetName() == "Line1" || swSketchSeg.GetName() == "Line2")
                    {
                        swSketchSeg.Select2(false, -1);
                        swModelDoc.AddDimension2(0, 0, 0);
                    }
                }

                swMySketch = (Sketch)swSketchMgr.ActiveSketch;
                swModelDoc.InsertSketch2(false);

                object[] sketchRegions = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)sketchRegions[0]).Select2(true, null);

                // Extrude the base plate
                swFeat = swFeatMgr.FeatureExtrusion3(true, false, true,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    Thickness * mmToMeter, 0, false, false, false, false,
                    0, 0, false, false, false, false, false, false, true,
                    (int)swStartConditions_e.swStartSketchPlane, 0, false);

                if (swFeat != null)
                {
                    CreateFilletFeature(swModelDoc, FilletRadius * mmToMeter);

                    swFeat = CreateSimpleHole(swModelDoc, HoleDiameter * mmToMeter);

                    if (swFeat != null)
                    {
                        swFeat = CreateLinearPattern(swModelDoc, swFeat);
                        if (swFeat == null)
                            System.Diagnostics.Debug.Print("Failed to create the Linear pattern feature");
                    }
                }
            }
            string fileName = "BasePlate.SLDPRT";
            this.SavePartFile(swModelDoc, fileName);
        }

        private void CreateFilletFeature(ModelDoc2 swModelDoc, double FilletRadius)
        {
            bool currHiddenEdgeSelInHLR;
            Edge swEdge;
            SimpleFilletFeatureData2 swSimpFillFeatData;
            SelectionMgr swSelMgr;

            bool boolStatus;
            int i;

            swModelDoc.ClearSelection2(true);

            swModelDoc.ShowNamedView2("*Trimetric", 8);
            swModelDoc.ViewZoomtofit2();

            currHiddenEdgeSelInHLR = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInHLR);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInHLR, true);

            // Select 4 edges by Ray
            boolStatus = swModelDoc.Extension.SelectByRay(-0.402030946619732, 0.401577459717771, -2.46568214199101E-02,
                -0.400036026779312, -0.515038074910024, -0.758094294050284,
                3.24984093911253E-03, 1, false, 1, 0);

            boolStatus = swModelDoc.Extension.SelectByRay(0.399953568098908, 0.400036064194069, -1.42885535830146E-02,
                -0.400036026779312, -0.515038074910024, -0.758094294050284,
                3.24984093911253E-03, 1, true, 1, 0);

            boolStatus = swModelDoc.Extension.SelectByRay(0.402376933526455, -0.401846191748291, -1.78884905474206E-02,
                -0.400036026779312, -0.515038074910024, -0.758094294050284,
                3.24984093911253E-03, 1, true, 1, 0);

            boolStatus = swModelDoc.Extension.SelectByRay(-0.399555786662972, -0.400345025634067, -1.99662322867198E-02,
                -0.400036026779312, -0.515038074910024, -0.758094294050284,
                3.24984093911253E-03, 1, true, 1, 0);

            swSimpFillFeatData = (SimpleFilletFeatureData2)swModelDoc.FeatureManager.CreateDefinition((int)swFeatureNameID_e.swFmFillet);

            swSimpFillFeatData.Initialize((int) swSimpleFilletType_e.swConstRadiusFillet);

            swSelMgr = (SelectionMgr)swModelDoc.SelectionManager;

            Edge[] edgesToFillet = new Edge[4];

            for (i = 1; i <= swSelMgr.GetSelectedObjectCount(); i++)
            {
                swEdge = (Edge)swSelMgr.GetSelectedObject6(i, -1);
                edgesToFillet[i - 1] = swEdge;
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
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInHLR, currHiddenEdgeSelInHLR);
        }

        private Feature CreateSimpleHole(ModelDoc2 swModelDoc, double HoleDiameter)
        {
            SketchManager swSketchMgr;
            FeatureManager swFeatMgr;
            Feature swFeat;

            bool boolStatus;

            swSketchMgr = swModelDoc.SketchManager;
            swFeatMgr = swModelDoc.FeatureManager;

            // Select reference plane
            boolStatus = swModelDoc.Extension.SelectByID2("Front Plane", "PLANE", -0.31, -0.31, 0, false, 0, null, 0);

            swFeat = swFeatMgr.HoleWizard5(
                (int)swWzdGeneralHoleTypes_e.swWzdHole,
                (int)swWzdHoleStandards_e.swStandardISO,
                (int)swWzdHoleStandardFastenerTypes_e.swStandardISODrillSizes,
                "Ø25.5",
                1,
                HoleDiameter,
                0.035,
                -1, 1, 0, 0, 0, 0, 0, 0,
                -1, -1, -1, -1, -1, "",
                false, true, true, true, true, false);

            swModelDoc.ClearSelection2(true);
            return swFeat;
        }

        private Feature CreateLinearPattern(ModelDoc2 swModelDoc, Feature swFeat)
        {
            FeatureManager swFeatMgr;
            LinearPatternFeatureData swLinearPattData;
            bool boolStatus;

            swModelDoc.ClearSelection2(true);

            swFeat.Select2(false, 4);
            boolStatus = swModelDoc.Extension.SelectByID2("", "EDGE", 0, -0.4, 0, true, 1, null, (int)swSelectOption_e.swSelectOptionDefault);
            boolStatus = swModelDoc.Extension.SelectByID2("", "EDGE", -0.4, 0, 0, true, 2, null, (int)swSelectOption_e.swSelectOptionDefault);

            swFeatMgr = swModelDoc.FeatureManager;
            swLinearPattData = (LinearPatternFeatureData)swFeatMgr.CreateDefinition((int)swFeatureNameID_e.swFmLPattern);

            swLinearPattData.D1EndCondition = 0;
            swLinearPattData.D1ReverseDirection = false;
            swLinearPattData.D1Spacing = 0.3;
            swLinearPattData.D1TotalInstances = 3;

            swLinearPattData.D2EndCondition = 0;
            swLinearPattData.D2PatternSeedOnly = false;
            swLinearPattData.D2ReverseDirection = false;
            swLinearPattData.D2Spacing = 0.3;
            swLinearPattData.D2TotalInstances = 3;
            swLinearPattData.GeometryPattern = false;
            swLinearPattData.VarySketch = false;

            swFeat = swFeatMgr.CreateFeature(swLinearPattData);
            swModelDoc.ClearSelection2(true);

            return swFeat;
        }

        private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
        {

            string fullPath = Path.Combine(destinationPath, fileName);
            swModelDoc.SaveAsSilent(fullPath, false);
        }

    }
}