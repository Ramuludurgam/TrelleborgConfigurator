using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

public class Bushing(SldWorks swApp, string destPath)
{

    private SldWorks swApp = swApp;
    private string destinationPath = destPath;
    public void CreateBushing()
    {
        ModelDoc2 swModelDoc = default;
        SketchSegment swSketchSeg = default;
        SketchManager swSketchMgr = default;
        Sketch swMySketch = default;
        //PartDoc swPartDoc = default;
        FeatureManager swFeatMgr = default;
        Feature swFeat = default;

        //const double Width = 750;
        //const double HeightUpHoleCenter = 475;
        //const double Thickness = 30;
        //const double BaseHeight = 30;
        //const double InnerHoleDiameter = 227;
        //const double BassHoleDiameter = 400;
        //const double mmToMeter = 0.001;

        bool boolStatus;

        swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        swModelDoc = (ModelDoc2)swApp.NewPart();
        swSketchMgr = swModelDoc.SketchManager;
        swFeatMgr = swModelDoc.FeatureManager;

        boolStatus = swModelDoc.Extension.SelectByID2(
            "Front Plane",
            "PLANE",
            0, 0, 0,
            false,
            -1,
            null,
            (int)swSelectOption_e.swSelectOptionDefault
        );

        if (boolStatus)
        {
            swSketchMgr.InsertSketch(true);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            swSketchSeg = swSketchMgr.CreateCenterLine(0, 0, 0, 0, 0.15694741, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.055, 0, 0, -0.1125, 0, 0);
            swModelDoc.AddDimension2(0, 0, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.1125, 0, 0, -0.1125, 0.055, 0);
            swModelDoc.AddDimension2(0, 0, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.1125, 0.055, 0, -0.12, 0.055, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.12, 0.055, 0, -0.12, 0.11, 0);
            swModelDoc.AddDimension2(0, 0, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.12, 0.11, 0, -0.055, 0.11, 0);
            swSketchSeg = swSketchMgr.CreateLine(-0.055, 0.11, 0, -0.055, 0, 0);

            swMySketch = swSketchMgr.ActiveSketch;
            swModelDoc.InsertSketch2(false);

            object[] regions = (object[])swMySketch.GetSketchRegions();
            if (regions != null && regions.Length > 0)
            {
                ((SketchRegion)regions[0]).Select2(true, null);
            }

            swFeat = swFeatMgr.FeatureRevolve2(
                true, true, false, false, false, false,
                0, 0,
                6.2831853071796, 0,
                false, false,
                0.01, 0.01,
                0, 0, 0,
                true, true, true
            );

            boolStatus = swModelDoc.Extension.SelectByRay(
                -0.0347532160096193, 0.0712992175607496, -0.0426287928164584,
                -0.635742010143878, -0.426124113744574, -0.643622821397456,
                0.000973637416514609, 2, true, 0, 0
            );

            boolStatus = swModelDoc.InsertAxis2(true);
            swModelDoc.ShowNamedView2 ("*Trimetric",8);
        }

        string fileName = "Bushing.SLDPRT";
        this.SavePartFile(swModelDoc, fileName);
    }
    private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
    {

        string fullPath = Path.Combine(destinationPath, fileName);
        swModelDoc.SaveAsSilent(fullPath, false);
    }
}