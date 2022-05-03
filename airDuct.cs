using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;

namespace airDucts
{
	class airDuct
	{
		
		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		public void createPrVozd(double dlin, double shir, double vys,double zazor, IModelDoc2 Part)
		{
			//double zazor = 0.002;
			
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, zazor * 5, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0.0, dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0.0, -dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0.0, -dlin / 2, 0, 0.0);
			Part.ClearSelection2(true);


			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2, 0, 0.0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point1@Исходная точка", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, zazor * 5, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2 , 0, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз1");
			myDimension1.SystemValue = zazor;
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, zazor * 10, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0.0, dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0.0, -dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0.0, -dlin / 2,zazor, 0.0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, zazor *10, vys, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", -dlin / 2, zazor, vys, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension2 = Part.Parameter("D1@Эскиз2");
			myDimension2.SystemValue = zazor;
			Part.SketchManager.InsertSketch(true);

			// Select the sketches for the lofted bend feature
			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", 0, 0, 0, true, 1, null, 0);

			// Insert a lofted bend feature with two bends
			feat = Part.FeatureManager.InsertSheetMetalLoftedBend2(0, zazor, false, 0.0007366, true,
				(int)swLoftedBendFacetOptions_e.swBendsPerTransitionSegment, 0, 2, 0, 0);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();

		}

		public void createKrVozd(double diam, double vys,  double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			//swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam / 2, 0, 0, 0, diam / 2, 0, 1);

			Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension = Part.Parameter("D1@Эскиз1");
			myDimension.SystemValue = diam / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension2 = Part.Parameter("D2@Эскиз1");
			myDimension2.SystemValue = zazor;

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			//swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam / 2, 0, 0, 0, diam / 2, 0, 1);

			Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз2");
			myDimension3.SystemValue = diam / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, vys, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, vys, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, vys, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam / 2, vys, true, 0, null, 0);

			Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз2");
			myDimension4.SystemValue = zazor;

			Part.ClearSelection2(true);
			Part.SetPickMode();

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", 0, 0, 0, true, 1, null, 0);

			feat = Part.FeatureManager.InsertSheetMetalLoftedBend(0, zazor);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();
		}


	}
}
