using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

namespace Configurator
{
    public class Bracket(SldWorks swApp, string destPath)
    {

        private SldWorks swApp = swApp;
        private string destinationPath = destPath;
        public void CreateBracket()
        {

            ModelDoc2 swModelDoc = default;
            SketchSegment swSketchSeg = default;
            SketchManager swSketchMgr = default;
            Sketch swMySketch = default;
            PartDoc swPartDoc = default;
            FeatureManager swFeatMgr = default;
            Feature swFeat = default;

            bool boolStatus;
            bool userPrefDimStatus;
            int i;
            object vArrSketchSegs;

            // Constants
            const double Width = 750;                // in mm
            const double HeightUpHoleCenter = 475;   // in mm
            const double Thickness = 30;             // in mm
            const double BaseHeight = 30;            // in mm
            // const double FilletRadius = 35;       // in mm
            const double InnerHoleDiameter = 227;    // in mm
            const double BassHoleDiameter = 400;     // in mm
            const double mmToMeter = 0.001;          // convert mm to meters

            swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            swModelDoc = (ModelDoc2)swApp.NewPart();
            swSketchMgr = swModelDoc.SketchManager;
            swFeatMgr = swModelDoc.FeatureManager;

            boolStatus = swModelDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            if (boolStatus)
            {
                swSketchMgr.InsertSketch(true);
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

                swSketchSeg = swSketchMgr.CreateLine(-0.375, 0, 0, 0.375, 0, 0);   // Line1
                swModelDoc.AddDimension2(0, 0, 0);

                swSketchSeg = swSketchMgr.CreateLine(-0.375, 0, 0, -0.375, 0.03, 0); // Line2
                swModelDoc.AddDimension2(0, 0, 0);

                swSketchSeg = swSketchMgr.CreateLine(-0.375, 0.03, 0, -0.187915259, 0.543467915, 0); // Line3
                swSketchSeg = swSketchMgr.CreateLine(0.187915259, 0.543467915, 0, 0.375, 0.03, 0);   // Line4

                swSketchSeg = swSketchMgr.CreateLine(0.375, 0.03, 0, 0.375, 0, 0);
                swModelDoc.AddDimension2(0, 0, 0);

                swSketchSeg = swSketchMgr.CreateCircleByRadius(0, 0.475, 0, BassHoleDiameter * mmToMeter / 2);
                swModelDoc.AddDimension2(0, 0, 0);

                // swModelDoc.Extension.SelectByID2("", "SKETCHSEGMENT", 0, 0.3615, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);
                swSketchSeg.Select2(false, -1);

                swSketchMgr.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, 0, 0.3615, 0);

                swSketchSeg = swSketchMgr.CreateCircleByRadius(0, 0.475, 0, InnerHoleDiameter * mmToMeter / 2);
                swModelDoc.AddDimension2(0, 0, 0);

                swMySketch = swSketchMgr.ActiveSketch;
                swModelDoc.InsertSketch2(false);

                object[] sketchRegions = (object[])swMySketch.GetSketchRegions();
                if (sketchRegions != null && sketchRegions.Length > 0)
                {
                    ((SketchRegion)sketchRegions[1]).Select2(true,null);
                }

                swFeat = swFeatMgr.FeatureExtrusion3(
                    true, false, false,
                    (int)swEndConditions_e.swEndCondBlind,
                    (int)swEndConditions_e.swEndCondBlind,
                    Thickness * mmToMeter, 0,
                    false, false, false, false,
                    0, 0, false, false, false, false,
                    false, false, true,
                    (int)swStartConditions_e.swStartSketchPlane, 0, false
                );

                // ***Add reference Geometry***
                boolStatus = swModelDoc.Extension.SelectByRay(
                    -8.23642849272801E-02, 0.396907909693425, 1.76685275120008E-02,
                    -0.577381545199981, -0.577287712085548, -0.577381545199979,
                    2.96279342343824E-03, 2, true, 0, 0
                );

                boolStatus = swModelDoc.InsertAxis2(true);
            }

            string fileName = "Bracket.SLDPRT";
            this.SavePartFile(swModelDoc, fileName);
        }

        private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
        {
            
            string fullPath = Path.Combine(destinationPath, fileName);
            swModelDoc.SaveAsSilent(fullPath, false);
        }
    }
}