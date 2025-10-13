using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Runtime.InteropServices;

namespace Configurator
{

    public class RectPlate(SldWorks swApp, string destPath)
    {

        private SldWorks swApp = swApp;
        private string destinationPath = destPath;

        public void CreateRectPlate()
        {
            ModelDoc2 swModelDoc;
            SketchSegment swSketchSeg;
            SketchManager swSketchMgr;
            Sketch swMySketch;

            FeatureManager swFeatMgr;
            Feature swFeat;

            object[] vArrSketchSegs;
            int i;

            const double Width = 330;        // in mm
            const double Height = 354;       // in mm
            const double Thickness = 25;     // in mm
            const double mmToMeter = 0.001;  // conversion factor

            bool boolstatus;

            // Create new part document
            swModelDoc = (ModelDoc2)swApp.NewPart();

            swSketchMgr = swModelDoc.SketchManager;
            swFeatMgr = swModelDoc.FeatureManager;

            // Select the Front Plane
            boolstatus = swModelDoc.Extension.SelectByID2(
                "Front Plane", "PLANE", 0, 0, 0, false, -1, null, (int)swSelectOption_e.swSelectOptionDefault
            );

            if (boolstatus)
            {
                // Insert sketch on the selected plane
                swSketchMgr.InsertSketch(true);

                // Create center rectangle
                vArrSketchSegs = (object[])swSketchMgr.CreateCenterRectangle(
                    0, 0, 0,
                    (Height * mmToMeter) / 2,
                    (Width * mmToMeter) / 2,
                    0
                );

                // Constrain Line1 and Line2 with dimensions
                for (i = 0; i < vArrSketchSegs.Length; i++)
                {
                    swSketchSeg = (SketchSegment)vArrSketchSegs[i];

                    // Check line names (horizontal or vertical)
                    string name = swSketchSeg.GetName();
                    if (name == "Line1" || name == "Line2")
                    {
                        swSketchSeg.Select2(false, -1);
                        swModelDoc.AddDimension2(0, 0, 0);
                    }
                }

                // Store reference to active sketch
                swMySketch = swSketchMgr.ActiveSketch;

                // Exit sketch
                swModelDoc.InsertSketch2(false);

                // Select the sketch region to extrude
                object[] sketchRegions = (object[])swMySketch.GetSketchRegions();
                ((SketchRegion)sketchRegions[0]).Select2(true, null);

                // Create extrusion
                swFeat = swFeatMgr.FeatureExtrusion3(
                    true, false, true,
                    (int)swEndConditions_e.swEndCondBlind,
                    (int)swEndConditions_e.swEndCondBlind,
                    Thickness * mmToMeter,
                    0,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, false,
                    true,
                    (int)swStartConditions_e.swStartSketchPlane,
                    0,
                    false
                );

                // Select a point on the model for axis insertion (position & direction)
                boolstatus = swModelDoc.Extension.SelectByRay(
                    -0.177185977886722, 0.0896323681564013, -0.0249018620037873,
                    -0.400036026779312, -0.515038074910024, -0.758094294050284,
                    0.00129250184632756,
                    1, false, 0, 0
                );

                // Insert axis
                boolstatus = swModelDoc.InsertAxis2(true);
            }

            string fileName = "RectPlate.SLDPRT";
            this.SavePartFile(swModelDoc, fileName);
        }
        private void SavePartFile(ModelDoc2 swModelDoc, string fileName)
        {
      
            string fullPath = Path.Combine(destinationPath, fileName);
            swModelDoc.SaveAsSilent(fullPath, false);
        }
    }

}
