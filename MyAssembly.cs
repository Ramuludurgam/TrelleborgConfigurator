using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;

namespace Configurator
{
  
    public class TrelleborgAssembly (SldWorks swApp, string DestinationPath)
    {
        private SldWorks swApp = swApp;

        private string compsPath = DestinationPath;
        private const string basePlate = "BasePlate.SLDPRT";
        private const string bracket = "Bracket.SLDPRT";
        private const string bushing = "Bushing.SLDPRT";
        private const string supportBracket = "SupportBracket.SLDPRT";
        private const string ribBracket = "RibBracket.SLDPRT";
        private const string RectPlate = "RectPlate.SLDPRT";
        public void CreateRequiredPartFiles()
        {
            BasePlate basePlate = new BasePlate(this.swApp, compsPath);
            basePlate.CreateBasePlate();

            Bracket bracket = new Bracket(this.swApp,compsPath);
            bracket.CreateBracket();

            Bushing bushing = new Bushing(this.swApp, compsPath);
            bushing.CreateBushing();

            RibBracket ribBracket = new RibBracket(this.swApp, compsPath);
            ribBracket.CreateRibBracket();

            SupportBracket supportBracket = new SupportBracket(this.swApp, compsPath);
            supportBracket.CreateSupportBracket();

            RectPlate rectPlate = new RectPlate(this.swApp, compsPath);
            rectPlate.CreateRectPlate();
        }


        public void CreateFullAssembly()
        {
            ModelDoc2 swModelDoc;
            AssemblyDoc swAsmDoc;
            Component2 swInsertedComponent;

            // Create a new assembly document
            swModelDoc = (ModelDoc2)swApp.NewAssembly();
            swAsmDoc = (AssemblyDoc)swModelDoc;

            string asmFileName = swModelDoc.GetTitle();
            string asmName = GetNameWithoutExtension(asmFileName);

            // ---- Insert BasePlate ----
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(
                System.IO.Path.Combine(compsPath, basePlate),
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0);

            string compFileName = swModelDoc.GetTitle();

            swInsertedComponent = (Component2)swAsmDoc.AddComponent5(
                System.IO.Path.Combine(compsPath, basePlate),
                0, "", false, "", 0, 0, 0);

            swApp.CloseDoc(compFileName);

            string compNameBasePlate = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", (int)swStandardViews_e.swTrimetricView);
            swModelDoc.ViewZoomtofit2();

            // ---- Insert Bracket ----
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(
                System.IO.Path.Combine(compsPath, bracket),
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0);

            compFileName = swModelDoc.GetTitle();

            swInsertedComponent = (Component2)swAsmDoc.AddComponent5(
                System.IO.Path.Combine(compsPath, bracket),
                0, "", false, "", 0, 0, 0.270209143265979);

            swApp.CloseDoc(compFileName);

            string compNameBracket = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", (int)swStandardViews_e.swTrimetricView);
            swModelDoc.ViewZoomtofit2();

            // ---- Mates between Bracket and BasePlate ----
            string plane1Info = $"Top Plane@{compNameBasePlate}@{asmName}";
            string plane2Info = $"Front Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0.165, 0);
            swModelDoc.ClearSelection2(true);

            // Coincident mate: Front Plane@BasePlate to Top Plane@Bracket
            plane1Info = $"Front Plane@{compNameBasePlate}@{asmName}";
            plane2Info = $"Top Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // Coincident mate: Right Plane@Bracket to Right Plane@BasePlate
            plane1Info = $"Right Plane@{compNameBasePlate}@{asmName}";
            plane2Info = $"Right Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // ---- Insert Bushing ----
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(
                System.IO.Path.Combine(compsPath, bushing),
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0);

            compFileName = swModelDoc.GetTitle();

            swInsertedComponent = (Component2)swAsmDoc.AddComponent5(
                System.IO.Path.Combine(compsPath, bushing),
                0, "", false, "", 0, -0.251926676, 0.498503068);

            swApp.CloseDoc(compFileName);

            string compNameBushing = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", (int)swStandardViews_e.swTrimetricView);
            swModelDoc.ViewZoomtofit2();

            // Mates between Bracket and Bushing
            plane1Info = $"Top Plane@{compNameBushing}@{asmName}";
            plane2Info = $"Front Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0.055, 0);
            swModelDoc.ClearSelection2(true);

            plane1Info = $"Axis1@{compNameBushing}@{asmName}";
            plane2Info = $"Axis1@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "AXIS", 0, 0, 0, true, 0, null, 0);
            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            plane1Info = $"Right Plane@{compNameBushing}@{asmName}";
            plane2Info = $"Right Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);
            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // ---- Insert RibBracket ----
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(
                System.IO.Path.Combine(compsPath, ribBracket),
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0);

            compFileName = swModelDoc.GetTitle();

            swInsertedComponent = (Component2)swAsmDoc.AddComponent5(
                System.IO.Path.Combine(compsPath, ribBracket),
                0, "", false, "Default", -0.3, -0.4, 0.498503068);

            swApp.CloseDoc(compFileName);

            string compNameRibBracket = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", (int)swStandardViews_e.swTrimetricView);
            swModelDoc.ViewZoomtofit2();
            swModelDoc.EditRebuild3();

            // Angle mate 8.57 degrees (converted to radians below)
            plane1Info = $"Right Plane@{compNameRibBracket}@{asmName}";
            plane2Info = $"Right Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);
            double radians = 8.57 * (Math.PI / 180.0);
            swAsmDoc.AddMate((int)swMateType_e.swMateANGLE, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, true, 0, radians);
            swModelDoc.ClearSelection2(true);

            // Distance 30mm
            plane1Info = $"Top Plane@{compNameRibBracket}@{asmName}";
            plane2Info = $"Front Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignALIGNED, false, 0.03, 0);
            swModelDoc.ClearSelection2(true);

            // Distance 0.2 (your code used 0.2 m)
            plane1Info = $"Axis1@{compNameRibBracket}@{asmName}";
            plane2Info = $"Right Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignCLOSEST, false, 0.2, 0);
            swModelDoc.ClearSelection2(true);

            // Coincident Axis1@RibBracket to Top Plane@Bracket
            plane1Info = $"Axis1@{compNameRibBracket}@{asmName}";
            plane2Info = $"Top Plane@{compNameBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignCLOSEST, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // ---- Insert SupportBracket ----
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(
                System.IO.Path.Combine(compsPath, supportBracket),
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0);

            compFileName = swModelDoc.GetTitle();

            swInsertedComponent = (Component2)swAsmDoc.AddComponent5(
                System.IO.Path.Combine(compsPath, supportBracket),
                0, "", false, "Default", -0.2, 0, 0.498503068);

            swApp.CloseDoc(compFileName);

            string compNameSupportBracket = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", (int)swStandardViews_e.swTrimetricView);
            swModelDoc.ViewZoomtofit2();
            swModelDoc.EditRebuild3();

            // Coincident - Anti Aligned Right Plane@RibBracket to Right Plane@SupportBracket
            plane1Info = $"Right Plane@{compNameRibBracket}@{asmName}";
            plane2Info = $"Right Plane@{compNameSupportBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);
            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // Coincident - Aligned Top Plane@BasePlate to Top Plane@SupportBracket
            plane1Info = $"Top Plane@{compNameBasePlate}@{asmName}";
            plane2Info = $"Top Plane@{compNameSupportBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0);
            swModelDoc.ClearSelection2(true);

            // Coincident Axis1@RibBracket to Axis1@SupportBracket
            plane1Info = $"Axis1@{compNameRibBracket}@{asmName}";
            plane2Info = $"Axis1@{compNameSupportBracket}@{asmName}";

            swModelDoc.Extension.SelectByID2(plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0);
            swModelDoc.Extension.SelectByID2(plane2Info, "AXIS", 0, 0, 0, true, 0, null, 0);

            swAsmDoc.AddMate((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignCLOSEST, false, 0, 0);
            swModelDoc.ClearSelection2(true);


            // **** Insert RectPlate.SLDPRT ***
            swModelDoc = swApp.OpenDoc6(
                compsPath + "\\" + RectPlate,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "Default", 0, 0
            );

            compFileName = swModelDoc.GetTitle();

            swInsertedComponent = swAsmDoc.AddComponent5(
                compsPath + "\\" + RectPlate,
                (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                "",
                false,
                "Default",
                -0.2, 0, 0.1
            );

            swApp.CloseDoc(compFileName);

            string compNameRectPlate = swInsertedComponent.Name2;

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDoc.ShowNamedView2("*Trimetric", 8);
            swModelDoc.ViewZoomtofit2();

            swModelDoc.EditRebuild3();

            // *** Add Mates ***

            // Mate 1: Coincident - Aligned
            // Top Plane@RectPlate
            // Top Plane@BasePlate

            plane1Info = "Top Plane@" + compNameRectPlate + "@" + asmName;
            plane2Info = "Top Plane@" + compNameBasePlate + "@" + asmName;

            bool boolstatus;

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0
            );

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0
            );

            swAsmDoc.AddMate(
                (int)swMateType_e.swMateCOINCIDENT,
                (int)swMateAlign_e.swMateAlignALIGNED,
                false, 0, 0
            );

            swModelDoc.ClearSelection2(true);

            // Mate 2: Coincident - Closest
            // Axis1@RectPlate
            // Right Plane@SupportBracket

            plane1Info = "Axis1@" + compNameRectPlate + "@" + asmName;
            plane2Info = "Right Plane@" + compNameSupportBracket + "@" + asmName;

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0
            );

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0
            );

            swAsmDoc.AddMate(
                (int)swMateType_e.swMateCOINCIDENT,
                (int)swMateAlign_e.swMateAlignCLOSEST,
                false, 0, 0
            );

            swModelDoc.ClearSelection2(true);

            // Mate 3: Distance = 197.3042957 mm
            // Axis1@RectPlate
            // Front Plane@BasePlate

            plane1Info = "Axis1@" + compNameRectPlate + "@" + asmName;
            plane2Info = "Front Plane@" + compNameBasePlate + "@" + asmName;

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane1Info, "AXIS", 0, 0, 0, false, 0, null, 0
            );

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0
            );

            swAsmDoc.AddMate(
                (int)swMateType_e.swMateDISTANCE,
                (int)swMateAlign_e.swMateAlignCLOSEST,
                false, 0.1973042957, 0
            );

            swModelDoc.ClearSelection2(true);

            // Mate 4: Angle = 84.9°
            // Front Plane@RectPlate
            // Right Plane@SupportBracket

            plane1Info = "Front Plane@" + compNameRectPlate + "@" + asmName;
            plane2Info = "Right Plane@" + compNameSupportBracket + "@" + asmName;

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane1Info, "PLANE", 0, 0, 0, false, 0, null, 0
            );

            boolstatus = swModelDoc.Extension.SelectByID2(
                plane2Info, "PLANE", 0, 0, 0, true, 0, null, 0
            );

            // Convert degrees to radians 84.90155177
            double angleRadians = 84.88 * (Math.PI / 180.0);

            swAsmDoc.AddMate(
                (int)swMateType_e.swMateANGLE,
                (int)swMateAlign_e.swMateAlignCLOSEST,
                false, 0, angleRadians
            );

            swModelDoc.ClearSelection2(true);

            this.MirrorComponents(swModelDoc);

            // Final rebuild
            swModelDoc.EditRebuild3();
        }

        private void MirrorComponents(ModelDoc2 swModelDoc)
        {
            if (swModelDoc == null) return;

            FeatureManager swFeatMgr = swModelDoc.FeatureManager;
            SelectionMgr swSelMgr = (SelectionMgr)swModelDoc.SelectionManager;
            AssemblyDoc swAsmDoc = (AssemblyDoc)swModelDoc;

            string asmFileName = swModelDoc.GetTitle();
            string asmName = GetNameWithoutExtension(asmFileName);

            // Select components and the mirror plane for first mirror operation
            swModelDoc.Extension.SelectByID2($"SupportBracket-1@{asmName}", "COMPONENT", 0, 0, 0, false, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"RibBracket-1@{asmName}", "COMPONENT", 0, 0, 0, true, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"Right Plane@Bracket-1@{asmName}", "PLANE", 0, 0, 0, true, 1, null, 0);

            int selCount = swSelMgr.GetSelectedObjectCount();
            if (selCount <= 1) return;

            // Build Component2 array
            Component2[] comps = new Component2[selCount - 1];
            for (int i = 1; i < selCount; i++)
            {
                comps[i - 1] = (Component2)swSelMgr.GetSelectedObjectsComponent4(i, 2);
            }

            // orientations array
            swMirrorComponentOrientation2_e[] orientations = new swMirrorComponentOrientation2_e[comps.Length];
            for (int i = 0; i < orientations.Length; i++)
            {
                orientations[i] = (swMirrorComponentOrientation2_e)swMirrorComponentOrientation2_e.swOrientation_MirroredAndFlippedX_MirroredY;
            }

            Feature swPlaneFeat = (Feature)swSelMgr.GetSelectedObject6(1, 1);
            RefPlane swRefMirrPlane = (RefPlane)swPlaneFeat.GetSpecificFeature2();
           
            // Create first mirror feature
            var mirrorDef = (MirrorComponentFeatureData)swFeatMgr.CreateDefinition((int)swFeatureNameID_e.swFmMirrorComponent);
            mirrorDef.MirrorPlane = swRefMirrPlane;
            mirrorDef.ComponentsToInstanceAlignToComponentOrigin = (object)comps;
            mirrorDef.MirrorType = (int)swMirrorComponentMirrorType_e.swMirrorType_CenterOfBoundingBox;
            mirrorDef.ComponentOrientationsAlignToComponentOrigin = (object)orientations;

            swFeatMgr.CreateFeature(mirrorDef);
            swModelDoc.ClearSelection2(true);

            // Change configuration of mirrored SupportBracket-2 to "Right"
            try
            {
                Component2 compMirSupport = swAsmDoc.GetComponentByName("SupportBracket-2");
                if (compMirSupport != null)
                {
                    compMirSupport.ReferencedConfiguration = "Right";
                }
            }
            catch { /* ignore if not present */ }

            // Prepare second mirror group selection
            swModelDoc.Extension.SelectByID2($"Bracket-1@{asmName}", "COMPONENT", 0, 0, 0, false, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"Bushing-1@{asmName}", "COMPONENT", 0, 0, 0, true, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"RibBracket-1@{asmName}", "COMPONENT", 0, 0, 0, true, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"RibBracket-2@{asmName}", "COMPONENT", 0, 0, 0, true, 2, null, 0);
            swModelDoc.Extension.SelectByID2($"Top Plane@BasePlate-1@{asmName}", "PLANE", 0, 0, 0, true, 1, null, 0);

            selCount = swSelMgr.GetSelectedObjectCount();
            if (selCount <= 1) return;

            Component2[] comps2 = new Component2[selCount - 1];
            for (int i = 1; i < selCount; i++)
            {
                comps2[i - 1] = (Component2)swSelMgr.GetSelectedObjectsComponent4(i, 2);
            }

            swMirrorComponentOrientation2_e[] orientations2 = new swMirrorComponentOrientation2_e[comps2.Length];
            for (int i = 0; i < orientations2.Length; i++)
                orientations2[i] = (swMirrorComponentOrientation2_e)swMirrorComponentOrientation2_e.swOrientation_MirroredAndFlippedX_MirroredY;
            swPlaneFeat = (Feature)swSelMgr.GetSelectedObject6(1, 1);
            swRefMirrPlane = (RefPlane)swPlaneFeat.GetSpecificFeature2();
            var mirrorDef2 = (MirrorComponentFeatureData)swFeatMgr.CreateDefinition((int)swFeatureNameID_e.swFmMirrorComponent);
            mirrorDef2.MirrorPlane = swRefMirrPlane;
            mirrorDef2.ComponentsToInstanceAlignToComponentOrigin = (object)comps2;
            mirrorDef2.MirrorType = (int)swMirrorComponentMirrorType_e.swMirrorType_CenterOfBoundingBox;
            mirrorDef2.ComponentOrientationsAlignToComponentOrigin = (object)orientations2;

            swFeatMgr.CreateFeature(mirrorDef2);
            swModelDoc.ClearSelection2(true);

            // Set referenced configurations for mirrored rib brackets (if present)
            try
            {
                Component2 rib3 = swAsmDoc.GetComponentByName("RibBracket-3");
                if (rib3 != null) rib3.ReferencedConfiguration = "Hole";

                Component2 rib4 = swAsmDoc.GetComponentByName("RibBracket-4");
                if (rib4 != null) rib4.ReferencedConfiguration = "Hole";
            }
            catch { /* ignore */ }

            swModelDoc.EditRebuild3();
            swModelDoc.ClearSelection2(true);
            swModelDoc.EditRebuild3();

            string fileName = "TrelleBorgAsm.SLDASM";
            this.SaveAssembly(swModelDoc, fileName);
        }

        public void CreateDrawingFile()
        {
            ModelDoc2 swModelDoc;
            DrawingDoc swDrawingDoc;
            string drawingTemplate;
            string drawingFilePath;
            bool boolStatus;
            int errors = default;
            string fullPathName;

            // Activate assembly document or get the assembly model doc handle
            swApp.ActivateDoc3("TrelleBorgAsm.SLDASM", true, 0, ref errors);

            swModelDoc = (ModelDoc2)swApp.ActiveDoc;

            if (swModelDoc == null)
            {
                MessageBox.Show("Failed to get open assembly document: " + "TrelleBorgAsm.SLDASM", "Error", button: MessageBoxButton.OK, icon: MessageBoxImage.Exclamation);
                return;
            }

            // Create a new drawing document using default template
            drawingTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
            swDrawingDoc = (DrawingDoc)swApp.NewDocument(drawingTemplate, 0, 0, 0);

            if (swDrawingDoc == null)
            {
                MessageBox.Show("Failed to create drawing document.", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            fullPathName = swModelDoc.GetPathName();

            // Insert the standard 3 views (Front, Top, Right)
            boolStatus = swDrawingDoc.Create3rdAngleViews2(fullPathName);

            if (!boolStatus)
            {
                MessageBox.Show("Failed to create three standard views.", "Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            ((ModelDoc2)swDrawingDoc).EditRebuild3();

            //Add isometric view
            swDrawingDoc.CreateDrawViewFromModelView3(fullPathName, "*Isometric", 0.2, 0.16, 0.0);

            //Save the drawing in the same folder as the assembly
            drawingFilePath = fullPathName.Substring(0, fullPathName.LastIndexOf('.')) + ".SLDDRW";

            boolStatus = ((ModelDoc2)swDrawingDoc).SaveAs(drawingFilePath);

            if (boolStatus)
            {
                MessageBox.Show("Drawing saved as: " + drawingFilePath, "Saved", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                MessageBox.Show("Drawing was not saved.", "Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }
        
        private string GetNameWithoutExtension(string filename)
        {
            if (string.IsNullOrEmpty(filename)) return filename;
            int dot = filename.LastIndexOf('.');
            if (dot > 0) return filename.Substring(0, dot);
            return filename;
        }

        private void SaveAssembly(ModelDoc2 swModelDoc, string fileName)
        {
            string fullPath = Path.Combine(compsPath, fileName);
            swModelDoc.SaveAsSilent(fullPath,false);
        }
    }
}
