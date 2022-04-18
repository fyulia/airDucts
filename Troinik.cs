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

		public void createTroinikPr(double dlin1, double shir1, double vys1, double dlin2, double shir2, double vys2, double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, -shir1/2, 0, -dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(-dlin1 / 2, shir1 / 2, 0, dlin1 / 2, shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, shir1 / 2, 0, dlin1 / 2, -shir1 / 2, 0);
			skSegment = (SketchSegment)Part.SketchManager.CreateLine(dlin1 / 2, -shir1 / 2, 0, -dlin1 / 2 + zazor*4, -shir1 / 2, 0);
			
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -dlin1 / 2, -shir1 / 2, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point5", "SKETCHPOINT", -dlin1 / 2 + zazor*4, -shir1 / 2, 0, true, 0, null, 0);

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
			boolstatus = Part.Extension.SelectByID2("Неизвестный", "POINTREF", dlin1/2 + zazor, 0, vys1, true, 1, null, 0);

			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(1, 0, 4, 0, 0, 0);
			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Плоскость5", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
		}
	}
}
