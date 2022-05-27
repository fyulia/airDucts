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
	class Zaglushka
	{

		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		object vskLines;

		public void createZaglPr(double dlin, double shir,   double vys, IModelDoc2 Part)
		{
			Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);

			vskLines = Part.SketchManager.CreateCenterRectangle(0, 0, 0, -dlin / 2, shir / 2, 0);

			Part.SetPickMode();
			Part.ClearSelection2(true);

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);

			CustomBendAllowance customBendAllowanceData1;
			customBendAllowanceData1 = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData1.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(0.002, false, 0.001, 0.02, 0.01, true, 0, 0, 1, customBendAllowanceData1,
				false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);
			Part.ClearSelection2(true);

			createRebro(Part,vys, dlin);

			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Справа", "PLANE", 0, 0, 0, false, 2, null, 0);
			boolstatus = Part.Extension.SelectByID2("Ребро-кромка1", "BODYFEATURE", 0, 0, 0, true, 1, null, 0);


			feat = Part.FeatureManager.InsertMirrorFeature(false, false, false, false);

			Part.ClearSelection2(true);
			//boolstatus = Part.Extension.SelectByID2("", "EDGE", vys / 2, 0, 0, false, 0, null, 0);

			createRebro(Part,vys, 0, shir);
			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, false, 2, null, 0);
			boolstatus = Part.Extension.SelectByID2("Ребро-кромка2", "BODYFEATURE", 0, 0, 0, true, 1, null, 0);


			feat = Part.FeatureManager.InsertMirrorFeature(false, false, false, false);

			Part.ClearSelection2(true);
		}

		void createRebro(IModelDoc2 Part,double vys, double dlin = 0, double shir = 0)
		{
			bool bValue = false;
			Edge swEdge;
			double dAngle = 0;
			double dLength = 0;
			Feature feat = default(Feature);
			Entity swEntity = default(Entity);
			Sketch swSketch = default(Sketch);
			object[] vSketchSegments = null;
			SketchLine swSketchLine = default(SketchLine);
			SketchPoint swStartPoint = default(SketchPoint);
			SketchPoint swEndPoint = default(SketchPoint);
			int nOptions = 0;
			double dSize = 0;
			double dFactor1 = 0;
			double dFactor2 = 0;
			Edge[] aFlangeEdges = new Edge[1];
			object vFlangeEdges = null;
			Sketch[] aSketchFeats = new Sketch[1];
			object vSketchFeats = null;

			object vskLines;

			// Set the angle
			dAngle = (90.0 / 180.0) * 3.1415926535897;

			dLength = 0.02;

			// Select edge for flange
			bValue = Part.Extension.SelectByID2("", "EDGE", dlin / 2, shir / 2, 0, false, 0, null, 0);
			// Get edge
			swEdge = (Edge)((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
			// Insert a sketch for an edge flange
			feat = (Feature)Part.InsertSketchForEdgeFlange(swEdge, dAngle, false);
			// Select
			bValue = feat.Select2(false, 0);
			// Start sketch editing
			Part.EditSketch();
			// Get the active sketch
			swSketch = (Sketch)Part.SketchManager.ActiveSketch;
			// Add the edge to the sketch
			// Cast edge to entity
			swEntity = (Entity)swEdge;
			// Select edge
			bValue = swEntity.Select4(false, null);
			// Use the edge in the sketch
			bValue = Part.SketchManager.SketchUseEdge(false);
			// Get the created sketch line
			vSketchSegments = (object[])swSketch.GetSketchSegments();
			swSketchLine = (SketchLine)vSketchSegments[0];
			// Get start and end point
			swStartPoint = (SketchPoint)swSketchLine.GetStartPoint2();
			swEndPoint = (SketchPoint)swSketchLine.GetEndPoint2();
			// Create additional lines to define sketch
			// Set parameters defining the sketch geometry
			dSize = swEndPoint.X - swStartPoint.X;
			dFactor1 = 0.1;
			dFactor2 = 1.25;

			Part.SetAddToDB(true);
			Part.SetDisplayWhenAdded(false);
			Part.SketchManager.CreateLine(swStartPoint.X, swStartPoint.Y, 0.0,
				swStartPoint.X, swStartPoint.Y + vys, 0.0);
			Part.SketchManager.CreateLine(swStartPoint.X, swStartPoint.Y + vys, 0.0,
				swEndPoint.X, swStartPoint.Y + vys, 0.0);
			Part.SketchManager.CreateLine(swEndPoint.X, swEndPoint.Y, 0.0,
				swEndPoint.X, swEndPoint.Y + vys, 0.0);
			// Reset
			Part.SetDisplayWhenAdded(true);
			Part.SetAddToDB(false);
			Part.SketchManager.InsertSketch(true);
			nOptions = (int)swInsertEdgeFlangeOptions_e.swInsertEdgeFlangeUseDefaultRadius +
				(int)swInsertEdgeFlangeOptions_e.swInsertEdgeFlangeUseDefaultRelief;
			aFlangeEdges[0] = swEdge;
			aSketchFeats[0] = swSketch;
			vFlangeEdges = aFlangeEdges;
			vSketchFeats = aSketchFeats;
			feat = Part.FeatureManager.InsertSheetMetalEdgeFlange2((vFlangeEdges),
				(vSketchFeats), nOptions, dAngle, 0.0,
				(int)swFlangePositionTypes_e.swFlangePositionTypeBendOutside, vys,
				(int)swSheetMetalReliefTypes_e.swSheetMetalReliefNone, 0.0, 0.0,
			0.0, (int)swFlangeDimTypes_e.swFlangeDimTypeInnerVirtualSharp, null);



		}

		public void createZaglKrBottom(double diam,double vys,double thick,IModelDoc2 Part)
		{
			Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);

			skSegment = Part.SketchManager.CreateCircle(0, 0, 0, diam / 2, 0, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);

			CustomBendAllowance customBendAllowanceData1;
			customBendAllowanceData1 = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData1.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(thick, false, 0.001, 0.02, 0.01, 
				true, 0, 0, 1, customBendAllowanceData1,
				false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);
			Part.ClearSelection2(true);

		}

		public void createZaglKrTop(double diam, double vys, double thick, IModelDoc2 Part)
		{
			Part.SketchManager.InsertSketch(true);
			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			skSegment = (SketchSegment)Part.SketchManager.CreateArc(0, 0, 0, - diam / 2, 0, 0, 0, diam / 2, 0, 1);

			Part.AddDimension2(-diam * 2 / 3, diam, 2.5);
			//Part.InsertGtol();
			Part.ClearSelection2(true);
			var myDimension3 = Part.Parameter("D1@Эскиз1");
			myDimension3.SystemValue = diam / 2;

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, true, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
			Part.SketchAddConstraints("sgHORIZONTALPOINTS2D");
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Point1", "SKETCHPOINT", -diam / 2, 0, 0, false, 0, null, 0);
			boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, diam / 2, 0, true, 0, null, 0);

			Part.AddDimension2(-diam * 2 / 3, 0, 2.5);
			Part.ClearSelection2(true);

			var myDimension4 = Part.Parameter("D2@Эскиз1");
			myDimension4.SystemValue = thick;

			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);
			CustomBendAllowance customBendAllowanceData;
			customBendAllowanceData = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(thick, true, 0.0007366, vys, 0.01, 
				false, 0, 0, 1, customBendAllowanceData, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);

			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
			refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, vys, 0, 0, 0, 0);
			Part.ClearSelection2(true);
			boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.BlankRefGeom();
		}
	}
}
