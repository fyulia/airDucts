using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace airDucts
{
	class Troinik
	{
		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		object vskLines;

		public void createTroinikPr(double dlin1, double shir1, double vys1, double dlin2, double shir2, double vys2, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, -shir1 / 2, 0, -dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, shir1 / 2, 0, dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, shir1 / 2, 0, dlin1 / 2, -shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, -shir1 / 2, 0, -dlin1 / 2 + zazor * 5, -shir1 / 2, 0);

			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin1 / 2, -shir1 / 2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -dlin1 / 2 + zazor * 5, -shir1 / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin1 * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз1");
			myDimension1.SystemValue = zazor;
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys1, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, -shir1 / 2, 0, -dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, shir1 / 2, 0, dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, shir1 / 2, 0, dlin1 / 2, -shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, -shir1 / 2, 0, -dlin1 / 2 + zazor * 5, -shir1 / 2, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin1 / 2, -shir1 / 2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -dlin1 / 2 + zazor * 5, -shir1 / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin1 * 2 / 3, 0, 2.5);
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


			boolstatus = Part.Extension.SelectByID2("Справа", "PLANE", 0, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Неизвестный", "POINTREF", dlin1 / 2 + zazor, 0, vys1, true, 1, null, 0);

			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(1, 0, 4, 0, 0, 0);
			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin2 / 2, 0, 0, -dlin2 / 2, shir2 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin2 / 2, shir2 / 2, 0, dlin2 / 2, shir2 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin2 / 2, shir2 / 2, 0.0, dlin2 / 2, -shir2 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin2 / 2, -shir2 / 2, 0.0, -dlin2 / 2, -shir2 / 2, 0.0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin2 / 2, -shir2 / 2, 0.0, -dlin2 / 2, -zazor * 5, 0.0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", dlin2 / 2, 0, 0, false, 0, null, 0);
			Part.SelectMidpoint();
			boolstatus = Part.Extension.SelectByID2("Point1@Исходная точка", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
		

			Part.AddDimension2(-dlin2 * 2 / 3, 0, 0);
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз5");
			myDimension3.SystemValue = -vys1 / 2 + dlin2 / 2;
			Part.ClearSelection2(true);
			Part.SetPickMode();

			boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -vys1 / 2 + dlin2 / 2, 0, 0, false, 0, null, 0);
			Part.AddDimension2(-dlin2 * 2, 0, 0);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз5");
			myDimension4.SystemValue = shir2;
			Part.ClearSelection2(true);
			

			boolstatus = Part.Extension.SelectByID2("Line4", "SKETCHSEGMENT", -vys1 / 2 + dlin2 / 2, -shir2 / 2, 0, false, 0, null, 0);
			Part.AddDimension2(-dlin2 * 2, 0, 0);
			Part.ClearSelection2(true);

			var myDimension5 = Part.Parameter("D3@Эскиз5");
			myDimension5.SystemValue = dlin2;
			Part.ClearSelection2(true);
			

			boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -vys1 / 2 + dlin2 / 2, 0, 0, true, 0, null, 0);
			Part.SelectMidpoint();
			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -vys1 / 2 - dlin2 / 2, 0, 0, true, 0, null, 0);

			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -vys1 / 2 - dlin2 / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", -vys1 / 2 - dlin2 / 2, -zazor * 5, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin2 * 2, 0, 0);
			Part.ClearSelection2(true);

			var myDimension6 = Part.Parameter("D4@Эскиз5");
			myDimension6.SystemValue = zazor;
			Part.ClearSelection2(true);
			Part.SetPickMode();

			boolstatus = Part.Extension.SelectByRay(dlin1 / 2, shir1 / 2 + zazor, vys1 / 2, -1, 0, 0, 9.74884895444939E-04, 1, false, 0, 0);
			Part.SelectMidpoint();
			boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", dlin1 / 2 + zazor, shir2 / 2, vys1 / 2, true, 0, null, 0);
			
			Part.AddDimension2(-dlin2 * 2, 0, 0);
			Part.ClearSelection2(true);

			var myDimension7 = Part.Parameter("D5@Эскиз5");
			myDimension7.SystemValue = shir1 / 2 + zazor - shir2 / 2;
			Part.ClearSelection2(true);
			Part.SetPickMode();


			boolstatus = Part.Extension.SelectByID2("Эскиз5", "SKETCH", 0, 0, 0, false, 0, null, 0);
			CustomBendAllowance customBendAllowanceData;
			customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(zazor, true, 0.0007366, vys2, 0, false, 0, 0, 1, customBendAllowanceData, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);


			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			vskLines = Part.SketchManager.CreateCenterRectangle(-vys1 / 2, 0, 0, -vys1 / 2 - dlin2 / 2, shir2 / 2, 0);
			Part.SketchManager.InsertSketch(true);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз8", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, false, 2, 0, 0.01, 0.01, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
			Part.SelectionManager.EnableContourSelection = false;

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();



		}

		public void createTroinikPrKr(double dlin, double shir, double vys1, double diam, double vys2, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0, dlin / 2, -shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0, -dlin / 2 + zazor * 5, -shir / 2, 0);

			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, -shir / 2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -dlin / 2 + zazor * 5, -shir / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-dlin * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз1");
			myDimension1.SystemValue = zazor;
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys1, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, -shir / 2, 0, -dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin / 2, shir / 2, 0, dlin / 2, shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, shir / 2, 0, dlin / 2, -shir / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin / 2, -shir / 2, 0, -dlin / 2 + zazor * 5, -shir / 2, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin / 2, -shir / 2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -dlin / 2 + zazor * 5, -shir / 2, 0, true, 0, null, 0);

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

			boolstatus = Part.Extension.SelectByID2("Справа", "PLANE", 0, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Неизвестный", "POINTREF", dlin / 2 + zazor, 0, vys1, true, 1, null, 0);

			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(1, 0, 4, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(-vys1/2, 0, 0, -vys1/2-diam / 2, 0, 0, -vys1/2, diam / 2, 0, 1);

			Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз5");
			myDimension3.SystemValue = diam / 2;


			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -vys1 / 2 - diam / 2, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", -vys1/2, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -vys1 / 2 - diam / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", -vys1/2, diam / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз5");
			myDimension4.SystemValue = zazor;

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз5", "SKETCH", 0, 0, 0, false, 0, null, 0);
			CustomBendAllowance customBendAllowanceData;
			customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(zazor, false, 0.0007366, vys2, 0.01, false, 0, 0, 1, customBendAllowanceData, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateCircle(-vys1/2, 0, 0, -vys1/2 + diam/2, 0, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз8", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, false, 2, 0, 0.01, 0.01, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();

		}

		public void createTroinikKr(double diam1,  double vys1, double diam2, double vys2, double zazor, IModelDoc2 Part)
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
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys1, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, -diam1 / 2, 0, 0, 0, diam1 / 2, 0, 1);

			Part.AddDimension2(-diam1 * 2 / 3, diam1, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension3 = Part.Parameter("D1@Эскиз2");
			myDimension3.SystemValue = diam1 / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam1 / 2, 0, vys1, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, vys1, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam1 / 2, 0, vys1, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam1 / 2, vys1, true, 0, null, 0);

			Part.AddDimension2(-diam1 * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз2");
			myDimension4.SystemValue = zazor;

			Part.ClearSelection2(true);
			Part.SetPickMode();

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", -diam1/2, 0, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", -diam1/2, 0, vys1, true, 1, null, 0);

			feat = Part.FeatureManager.InsertSheetMetalLoftedBend(1, zazor);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();

			boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, true, 0, null, 0);

			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, diam1/2 + zazor, 0, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateArc(0, -vys1/2, 0, -0, -vys1/2-diam2/2, 0, -diam2/2, -vys1/2, 0, 1);

			Part.AddDimension2(-diam2 * 2 / 3, diam2, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension1 = Part.Parameter("D1@Эскиз5");
			myDimension1.SystemValue = diam2 / 2;

			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, -vys1/2, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, -vys1/2-diam2/2, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgVERTICALPOINTS2D");

			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", -diam2/2, -vys1/2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, -vys1/2-diam2/2, 0, true, 0, null, 0);
			
			Part.AddDimension2(-diam2 , diam2, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);

			var myDimension5 = Part.Parameter("D2@Эскиз5");
			myDimension5.SystemValue = zazor;
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз5", "SKETCH", 0, 0, 0, false, 0, null, 0);

			CustomBendAllowance customBendAllowanceData;
			customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(zazor, false, 0.00015, vys2, vys2/2, false, 0, 0, 0, customBendAllowanceData, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);

			Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			skSegment = Part.SketchManager.CreateCircle(0, -vys1/2, 0, 0, -vys1/2+diam2/2, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз8", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, false, 2, 0, 0.01, 0.01, false, false, false, false,
				0, 0, false, false, false, false, false, true, true,
				true, true, false, 0, 0, false, false);

			Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			skSegment = Part.SketchManager.CreateCircle(0, 0, 0, -diam1 / 2, 0, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз9", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, true, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0,
				false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);

			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();

		}
	}
}
