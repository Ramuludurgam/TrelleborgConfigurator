using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

public class RibBracket(SldWorks swApp, string destPath)
{

    private SldWorks swApp = swApp;
    private string destinationPath = destPath; 
    public void CreateRibBracket()
    {
        ModelDoc2 swModelDoc;
        SketchSegment swSketchSeg;
        SketchManager swSketchMgr;
        Sketch swMySketch;
        PartDoc swPartDoc;

        SelectionMgr swSelMgr;
        FeatureManager swFeatMgr;
        Feature swFeat;

        // const double Length = 0.0;
        //const double Width = 750;               // in mm
        //const double HeightUpHoleCenter = 475;  // in mm
        const double Thickness = 25;            // in mm
        //const double BaseHeight = 30;           // in mm
        const double HoleDiameter = 30;         // in mm
        const double mmToMeter = 0.001;         // conversion factor

        bool boolStatus;

        Edge[] edgesToFillet = new Edge[1];

        string[] vArrConfigs = new string[1];

        swModelDoc = (ModelDoc2)swApp.NewPart();
        swSketchMgr = swModelDoc.SketchManager;
        swFeatMgr = swModelDoc.FeatureManager;
        swSelMgr = (SelectionMgr)swModelDoc.SelectionManager;
        swPartDoc = (PartDoc)swModelDoc;

        boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

        if (boolStatus)
        {
            swSketchMgr.InsertSketch(true);

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            swSketchSeg = swSketchMgr.CreateLine(-0.025, 0, 0, -0.5, 0, 0);     // Line1
            swSketchSeg = swSketchMgr.CreateLine(-0.5, 0, 0, -0.5, 0.03, 0);    // Line2
            // swModelDoc.AddDimension2(0, 0, 0);

            swSketchSeg = swSketchMgr.CreateLine(-0.5, 0.03, 0, -0.03, 0.18, 0); // Line3
            // swModelDoc.AddDimension2(0, 0, 0);

            swSketchSeg = swSketchMgr.CreateLine(-0.03, 0.18, 0, 0, 0.18, 0);    // Line4
            swSketchSeg = swSketchMgr.CreateLine(0, 0.18, 0, 0, 0.025, 0);       // Line5
            // swModelDoc.AddDimension2(0, 0, 0);

            swSketchSeg = swSketchMgr.CreateLine(0, 0.025, 0, -0.025, 0, 0);     // Chamfer Line

            swMySketch = (Sketch)swSketchMgr.ActiveSketch;
            swModelDoc.InsertSketch2(false);

            object[] sketchRegions = (object[])swMySketch.GetSketchRegions();
            ((SketchRegion)sketchRegions[0]).Select2(true, null);

            swFeat = swFeatMgr.FeatureExtrusion3(
                true, true, false,
                (int)swEndConditions_e.swEndCondBlind,
                (int)swEndConditions_e.swEndCondBlind,
                Thickness * mmToMeter, 0,
                false, false, false, false, 0, 0,
                false, false, false, false, false, false,
                true,
                (int)swStartConditions_e.swStartSketchPlane,
                0, false);

            boolStatus = swModelDoc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault);

            swSketchMgr.InsertSketch(true);
            swSketchSeg = swSketchMgr.CreateCircleByRadius(-0.375, 0, 0, HoleDiameter * mmToMeter / 2);

            swMySketch = (Sketch)swSketchMgr.ActiveSketch;
            swModelDoc.InsertSketch2(false);

            sketchRegions = (object[])swMySketch.GetSketchRegions();
            ((SketchRegion)sketchRegions[0]).Select2(true, null);

            Feature swHoleFeature = swFeatMgr.FeatureCut4(
                false, false, false,
                (int)swEndConditions_e.swEndCondBlind,
                1,
                0.01, 0.01,
                false, false, false, false, 0, 0,
                false, false, false, false, false,
                true, true, true, true, false, 0, 0, false, false);

            string featName = swHoleFeature.Name;

            swModelDoc.ShowNamedView2("*Trimetric", 8);
            swModelDoc.ViewZoomtofit2();

            // Add reference geometry ***
            boolStatus = swModelDoc.Extension.SelectByRay(
                -4.51586678707372E-04,
                0.109735562922637,
                3.29842311430184E-04,
                0.579213816122904,
                -0.188841090161286,
                0.793000881386043,
                1.32218314322423E-03,
                1,
                false, 0, 0);

            boolStatus = swModelDoc.InsertAxis2(true);

            // ***** Working with configurations *****
            ConfigurationManager swConfigMgr = (ConfigurationManager)swModelDoc.ConfigurationManager;
            Configuration swConfig = swConfigMgr.AddConfiguration2(
                "Hole",
                "With stress relieving hole",
                "None",
                (int)swConfigurationOptions2_e.swConfigOption_UseAlternateName,
                "",
                "Hole",
                true);

            vArrConfigs[0] = "Default";
            swHoleFeature.SetSuppression2(
                (int)swFeatureSuppressionAction_e.swSuppressFeature,
                (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                vArrConfigs);
        }

        string fileName = "RibBracket.SLDPRT";
        this.SavePartFile(swModelDoc, fileName);
    }

    private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
    {
 
        string fullPath = Path.Combine(destinationPath, fileName);
        swModelDoc.SaveAsSilent(fullPath, false);
    }
}