using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace airDucts
{
	class Otvod
	{
		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		DisplayDimension DisplayDimension = default(DisplayDimension);
		Dimension swDim = default(Dimension);
		public void createKrOtvod(double diam,  double zazor, IModelDoc2 Part)
		{
			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0, diam, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);
			skSegment = Part.SketchManager.CreateLine(0, 0, 0, diam, 0, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, diam / 2, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", diam / 2, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgSAMELENGTH");
			Part.ClearSelection2(true);

			skSegment = Part.SketchManager.CreateTangentArc(0, diam, 0, diam, 0, 0, 4);
			Part.ClearSelection2(true);

			double a = 0.261799;
			double x1 = diam  * Math.Cos(a);
			double y1 = diam  * Math.Sin(a);
			double m = 0.01;

			skSegment = Part.SketchManager.CreateLine(0, 0, 0, x1, y1, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			//Part.ClearSelection2(true);

			//Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);

			//Part.ClearSelection2(true);
			//Part.SetPickMode();
			skSegment = Part.SketchManager.CreateLine(0, 0, 0, 0, m, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Line3@Эскиз1", "EXTSKETCHSEGMENT", 2.03215877439393E-02, 5.44515302490674E-03, 0, true, 1, null, 0);

			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);


			//skSegment = (SketchSegment)Part.SketchManager.CreateArc(diam, 0, 0, diam / 2, 0, 0, 0, diam / 2, 0, 1);

			//Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			////Part.InsertGtol();
			//Part.ClearSelection2(true);

			//var myDimension = Part.Parameter("D1@Эскиз3");
			//myDimension.SystemValue = diam / 2;

			//boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", diam / 2, 0, 0, true, 0, null, 0);
			//boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", diam, 0, 0, true, 0, null, 0);
			//Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			//Part.ClearSelection2(true);

			//boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", diam / 2, 0, 0, false, 0, null, 0);
			//boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam / 2, 0, true, 0, null, 0);

			////Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			////Part.ClearSelection2(true);

			////var myDimension2 = Part.Parameter("D2@Эскиз3");
			//myDimension2.SystemValue = zazor;








			skSegment = Part.SketchManager.CreateCircle(diam, 0, 0, diam / 2, 0, 0);

			Part.ClearSelection2(true);
			skSegment = Part.SketchManager.CreateLine(diam, 0, 0, diam / 2, 0, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", diam, 0, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", diam, 0, 0, false, 1, null, 0);
			boolstatus = Part.SketchManager.SketchOffset2(zazor / 2, true, true, 0, 0, true);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", diam, 0.001, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", diam, -0.001, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", diam / 2, 0, 0, false, 2, null, 0);
			boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0);
			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", diam / 2, 0, 0, false, 2, null, 0);
			boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0);
			Part.SketchManager.InsertSketch(true);


			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateCircle(diam, 0, 0, diam / 2, 0, 0);

			Part.ClearSelection2(true);
			skSegment = Part.SketchManager.CreateLine(diam, 0, 0, diam / 2, 0, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", diam, 0, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", diam, 0, 0, false, 1, null, 0);
			boolstatus = Part.SketchManager.SketchOffset2(zazor / 2, true, true, 0, 0, true);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", diam, 0.001, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", diam, -0.001, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);

			double x2 = diam / 2 * Math.Cos(a);
			double y2 = diam / 2 * Math.Sin(a);

			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", x2, y2, 0, false, 2, null, 0);
			boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0);
			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", x2, y2, 0, false, 2, null, 0);
			boolstatus = Part.SketchManager.SketchTrim(4, 0, 0, 0);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз3", "SKETCH", diam/2, zazor/2, 0, false, 1, null, 0);
			boolstatus = Part.Extension.SelectByID2("Эскиз4", "SKETCH", diam/2, -zazor/2, 0, true, 1, null, 0);
			Part.FeatureManager.InsertSheetMetalLoftedBend(0, zazor);
			Part.ClearSelection2(true);



		}
	}
}
