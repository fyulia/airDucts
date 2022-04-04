using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;

namespace airDucts
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			
			dataTab();
			textBox2.Text = "Переход с прямоугольного на прямоугольное сечение является фасонным элементом, позволяющим соединить воздуховоды с разным размером прямоугольных сечений.";
			}
		/// <summary>
		/// 
		/// </summary>


		ISldWorks iSwApp = null;
		IModelDoc2 Part; //Предоставляет доступ к документам SOLIDWORKS: деталям, сборкам и чертежам.
						 //ShellFeatureData swShell;
		private void dataTab()
		{
			//прям возд
			cb1_Shir.Items.Clear();
			cb1_Shir.Items.Add("100");
			cb1_Shir.Items.Add("150");
			cb1_Shir.Items.Add("250");
			cb1_Shir.Items.Add("400");
			cb1_Shir.Items.Add("500");
			cb1_Shir.Items.Add("600");
			cb1_Shir.Items.Add("800");
			cb1_Shir.Items.Add("1000");
			cb1_Shir.Items.Add("1250");
			cb1_Shir.Items.Add("1600");

			cb1_Vys.Items.Clear();
			cb1_Vys.Items.Add("150");
			cb1_Vys.Items.Add("250");
			cb1_Vys.Items.Add("300");
			cb1_Vys.Items.Add("400");
			cb1_Vys.Items.Add("500");
			cb1_Vys.Items.Add("600");
			cb1_Vys.Items.Add("800");
			cb1_Vys.Items.Add("1000");
			cb1_Vys.Items.Add("1250");
			cb1_Vys.Items.Add("1600");
			cb1_Vys.Items.Add("2000");

			cb1_Dlin.Items.Clear();
			cb1_Dlin.Items.Add("2500");
			cb1_Dlin.Items.Add("3000");
			cb1_Dlin.Items.Add("4000");
			cb1_Dlin.Items.Add("5000");
			cb1_Dlin.Items.Add("6000");
		}

		



		private void bt_PrymVoz_Click(object sender, EventArgs e)
		{

			double shir_pr, vys_pr, dlin_pr;
			shir_pr = Convert.ToDouble(cb1_Shir.Text);
			vys_pr = Convert.ToDouble(cb1_Vys.Text);
			dlin_pr = Convert.ToDouble(cb1_Dlin.Text);

			shir_pr = shir_pr / 1000;
			vys_pr = vys_pr / 1000;
			dlin_pr = dlin_pr / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;
			
			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			airDuct duct = new airDuct();
			duct.createPrVozd(shir_pr,vys_pr,dlin_pr,Part);
			//setMaterial();


		}

		void setMaterial()
		{

			//задание материала
			string swMateDB = "";
			string tempMaterial = "";
			// Получить существующие материалы
			//tempMaterial = ((PartDoc)Part).GetMaterialPropertyName2("", out swMateDB);

			//MessageBox.Show("Текущий материал деталей - {TempMaterial}" );

			//string configName = null;
			//string databaseName = null;
			//string newPropName = null;
			string configName = "По умолчанию";
			string databaseName = "D:/SolidWorks/SOLIDWORKS/lang/russian/sldmaterials/solidworks materials.sldmat";
			string newPropName = "Оцинкованная сталь";
			((PartDoc)Part).SetMaterialPropertyName2(configName, databaseName, newPropName);

		}
	}
}
