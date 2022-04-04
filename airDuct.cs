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
		 public void createPrVozd(double shir, double vys, double dlin, IModelDoc2 Part)
		{
			double zazor = 0.002;
			
			

			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
				Part.ClearSelection2(true);
				Part.SketchManager.InsertSketch(true);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, 0, 0, -shir / 2, vys / 2, 0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, vys / 2, 0, shir / 2, vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(shir / 2, vys / 2, 0.0, shir / 2, -vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(shir / 2, -vys / 2, 0.0, -shir / 2, -vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, -vys / 2, 0.0, -shir / 2, -zazor, 0.0);
				Part.ClearSelection2(true);
				Part.SketchManager.InsertSketch(true);

				boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, true, 0, null, 0);
				refPlane = (RefPlane)Part.FeatureManager.InsertRefPlane(8, dlin, 0, 0, 0, 0);
				Part.ClearSelection2(true);

				boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
				Part.SketchManager.InsertSketch(true);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, 0, 0, -shir / 2, vys / 2, 0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, vys / 2, 0, shir / 2, vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(shir / 2, vys / 2, 0.0, shir / 2, -vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(shir / 2, -vys / 2, 0.0, -shir / 2, -vys / 2, 0.0);
				skSegment = (SketchSegment)Part.SketchManager.CreateLine(-shir / 2, -vys / 2, 0.0, -shir / 2, -zazor, 0.0);
				Part.ClearSelection2(true);
				Part.SketchManager.InsertSketch(true);

				// Select the sketches for the lofted bend feature
				boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 1, null, 0);
				boolstatus = Part.Extension.SelectByID2("Эскиз2", "SKETCH", 0, 0, 0, true, 1, null, 0);

				// Insert a lofted bend feature with two bends
				feat = Part.FeatureManager.InsertSheetMetalLoftedBend2(0, 0.002, false, 0.0007366, true,
					(int)swLoftedBendFacetOptions_e.swBendsPerTransitionSegment, 0, 2, 0, 0);

				boolstatus = Part.Extension.SelectByID2("Плоскость4", "PLANE", 0, 0, 0, false, 0, null, 0);
				Part.BlankRefGeom();

		} 




	}
}
