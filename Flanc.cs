using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace airDucts
{
	class Flanc
	{

		bool boolstatus;
		Feature feat;
		RefPlane refPlane;
		SketchSegment skSegment;
		object vskLines;

		public void createPrFlanc(double dlin1, double shir1, double dlin2, double shir2, double sDlin, double sShir, int countDlin, int countShir, double thick, double dOtv, IModelDoc2 Part)
		{
			double raznDlin = dlin2 - dlin1;
			double raznShir = shir2 - shir1;

			double xBok = dlin1+raznDlin*2 - sDlin * (countDlin - 1);
			double yBok = shir1 + raznShir * 2 - sShir * (countShir - 1);


			double xOtv = -dlin1 / 2 - raznDlin + xBok/2;
			double yOtv = -shir1 / 2 - raznShir / 2;

			double xOtv2 = -dlin1 / 2 - raznDlin / 2;
			double yOtv2 = -shir1 / 2 - raznShir + yBok / 2;

			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			vskLines = Part.SketchManager.CreateCenterRectangle(0, 0, 0, dlin1/2, shir1/2, 0);
			Part.ClearSelection2(true);
			vskLines = Part.SketchManager.CreateCenterRectangle(0, 0, 0, dlin1/2+raznDlin, shir1/2+raznShir, 0);
			Part.SetPickMode();
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);


			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);

			CustomBendAllowance customBendAllowanceData1;
			customBendAllowanceData1 = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData1.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(thick, false, 0.0007366, 0.02, 0.01, false, 0, 0, 1, customBendAllowanceData1, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);

			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateCircle(xOtv, yOtv, 0, xOtv-dOtv/2, yOtv, 0);

			boolstatus = Part.SketchManager.CreateLinearSketchStepAndRepeat(countDlin, 1, sDlin, 0, 0, 0, "", false, false, false, true, false);
			Part.ClearSelection2(true);

			for (int i = 1; i <= countDlin; i++)
			{
				boolstatus = Part.Extension.SelectByID2($"Arc{i}", "SKETCHSEGMENT", xOtv + sDlin * i - dOtv/2, yOtv, 0, true, 0, null, 0);
			}
			boolstatus = Part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, true, 0, null, 0);
			Part.SketchMirror();
			Part.ClearSelection2(true);

			skSegment = Part.SketchManager.CreateCircle(xOtv2, yOtv2, 0, xOtv2 - dOtv / 2, yOtv2, 0);

			boolstatus = Part.SketchManager.CreateLinearSketchStepAndRepeat(1, countShir, 0, sShir, 0, 1.5707963267949, "", false, false, false, false, true);

			//boolstatus = Part.SketchManager.CreateLinearSketchStepAndRepeat(countShir, 1, sShir, 0, 3.1415926535898, 0, "", false, false, false, true, false);


			Part.ClearSelection2(true);
			for (int i = countDlin*2+1; i <= countDlin*2 + countShir; i++)
			{
				int n = 1;
				boolstatus = Part.Extension.SelectByID2($"Arc{i}", "SKETCHSEGMENT", xOtv2, yOtv2+sShir*n, 0, true, 0, null, 0);
				n++;
			}
			boolstatus = Part.Extension.SelectByID2("Справа", "PLANE", 0, 0, 0, true, 0, null, 0);
			Part.SketchMirror();
			Part.ClearSelection2(true);

			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз6", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, true, true, false, 0, 0, false, true);


			Part.ClearSelection2(true);



		}

		public void createKrFlanc(double diam, double dOs, double dOtv, int count, double thick,  IModelDoc2 Part)
		{
			double raznDiam = dOs - diam;
			double sOtv = dOs / count;
			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateCircle(0, 0, 0, dOs/2, 0, 0);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 1, null, 0);
			boolstatus = Part.SketchManager.SketchOffset2(raznDiam/2, true, true, 0, 0, true);
			Part.ClearSelection2(true);

			boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", dOs/2, 0, 0, false, 0, null, 0);
			Part.SketchManager.CreateConstructionGeometry();
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);

			CustomBendAllowance customBendAllowanceData1;
			customBendAllowanceData1 = Part.FeatureManager.CreateCustomBendAllowance();
			customBendAllowanceData1.KFactor = 0.5;
			feat = Part.FeatureManager.InsertSheetMetalBaseFlange2(thick, false, 0.0007366, thick, 0.01, false, 0, 0, 1, customBendAllowanceData1, false, 0, 0.0001, 0.0001, 0.5, true, false, true, true);

			boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);
			skSegment = Part.SketchManager.CreateCircle(0, dOs/2, 0, dOtv / 2, dOs/2, 0);
			boolstatus = Part.SketchManager.CreateCircularSketchStepAndRepeat(dOs/2, 4.71238898038473, count, 4.71238898038473 / 2 * 4/ count, true, "", false, false, true);
			Part.ClearSelection2(true);
			Part.SketchManager.InsertSketch(true);

			boolstatus = Part.Extension.SelectByID2("Эскиз6", "SKETCH", 0, 0, 0, false, 0, null, 0);
			feat = Part.FeatureManager.FeatureCut4(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false,
				1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, true, true, false, 0, 0, false, true);
			Part.SelectionManager.EnableContourSelection = false;

			Part.ClearSelection2(true);




		}

	}
}
