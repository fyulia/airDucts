using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace airDucts
{

	class Perehod
	{

		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		DisplayDimension DisplayDimension = default(DisplayDimension);
		Dimension swDim = default(Dimension);

		public void createPrPrPerehod(double dlin,double dlin1, double shir,double shir1, double vys, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, zazor * 5, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0.0, dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0.0, -dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0.0, -dlin / 2, 0, 0.0);


			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2, 0, 0.0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point1@Исходная точка", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, zazor * 5, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2, 0, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз1");
			myDimension1.SystemValue = zazor;
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, zazor * 5, 0, -dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, shir1 / 2, 0, dlin1 / 2, shir1 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, shir1 / 2, 0.0, dlin1 / 2, -shir1 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, -shir1 / 2, 0.0, -dlin1 / 2, -shir1 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, -shir1 / 2, 0.0, -dlin1 / 2, 0, 0.0);
			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin1 / 2, zazor * 5, vys, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", -dlin1 / 2, 0, vys, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension2 = Part.Parameter("D1@Эскиз2");
			myDimension2.SystemValue = zazor;
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", dlin / 2, -zazor, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", dlin1 / 2, -zazor, vys, true, 1, null, 0);

			feat = Part.FeatureManager.InsertSheetMetalLoftedBend2(0, zazor, false, 0.001, true,
					(int)swLoftedBendFacetOptions_e.swBendsPerTransitionSegment, 0.0005, 2, 0.005, 30.00001286);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();
		}

		public void createPrPerehod(double dlin, double shir, double diam, double vys, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, zazor * 5, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0.0, dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0.0, -dlin / 2, -shir / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0.0, -dlin / 2, 0, 0.0);
			Part.SetPickMode();
			Part.ClearSelection2(true);
			//Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2, 0, 0.0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point1@Исходная точка", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, zazor * 5, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -dlin / 2, 0, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз1");
			myDimension1.SystemValue = zazor;



			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, shir / 2, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", dlin / 2, shir / 2, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", dlin / 2, -shir / 2, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point4", "SKETCHPOINT", -dlin / 2, -shir / 2, 0, true, 0, null, 0);

			skSegment = (SketchSegment)Part.SketchManager.CreateFillet(zazor, 1);
			skSegment = (SketchSegment)Part.SketchManager.CreateFillet(zazor, 1);
			skSegment = (SketchSegment)Part.SketchManager.CreateFillet(zazor, 1);
			skSegment = (SketchSegment)Part.SketchManager.CreateFillet(zazor, 1);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam / 2, 0, 0, 0, diam / 2, 0, 1);

			Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз2");
			myDimension3.SystemValue = diam / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension2 = Part.Parameter("D2@Эскиз2");
			myDimension2.SystemValue = zazor;

			Part.SketchManager.InsertSketch(true);
			Part.ClearSelection2(true);


			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", -dlin / 2, 0, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", -diam / 2, 0, vys, true, 1, null, 0);

			//feat = Part.FeatureManager.InsertSheetMetalLoftedBend(0, zazor);


			//feat = Part.FeatureManager.InsertSheetMetalLoftedBend2(0, zazor, false, 0.001, true, (int)swLoftedBendFacetOptions_e.swChordTolerance, 0.0005, 0, 0, 0);

			//Part.ClearSelection2(true);

			//Part.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayAnnotations, true);

			//boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			//Part.BlankRefGeom();
		}

		public void createKrPerehod(double diam1, double diam2,  double vys, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam1 / 2, 0, 0, 0, diam1 / 2, 0, 1);

			Part.AddDimension2(-diam1 * 2 / 3, diam1, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension = Part.Parameter("D1@Эскиз1");
			myDimension.SystemValue = diam1 / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam1 / 2, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam1 / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam1 / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-diam1 * 2 / 3, 0, 2.5);
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

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam2 / 2, 0, 0, 0, diam2 / 2, 0, 1);

			Part.AddDimension2(-diam2 * 2 / 3, diam2, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз2");
			myDimension3.SystemValue = diam2 / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam2 / 2, 0, vys, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, vys, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam2 / 2, 0, vys, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam2 / 2, vys, true, 0, null, 0);

			Part.AddDimension2(-diam2 * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз2");
			myDimension4.SystemValue = zazor;

			Part.ClearSelection2(true);
			Part.SetPickMode();

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", -diam1 / 2, 0, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", -diam2 / 2, 0, 0, true, 1, null, 0);

			feat = Part.FeatureManager.InsertSheetMetalLoftedBend(0, zazor);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);

			Part.BlankRefGeom();
		}
	}
}
