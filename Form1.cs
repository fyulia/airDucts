using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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

		private SqlConnection sqlConnection = null;
		private SqlDataAdapter sqlDataAdapter = null;
		private SqlDataAdapter sqlDataAdapter2 = null;
		private DataSet dataSet = null;
		private DataSet dataSet2 = null;

		public Form1()
		{
			InitializeComponent();

			dataTab();
			textBox2.Text = "Переход с прямоугольного на прямоугольное сечение является фасонным элементом, позволяющим соединить воздуховоды с разным размером прямоугольных сечений.";

			//SqlConnection sqlConnection = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = D:\универ\airDucts\airDucts\Database1.mdf; Integrated Security = True");
			//SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT DISTINCT dlin FROM VozdPr ", sqlConnection);

			//DataSet dataSet = new DataSet();
			//sqlDataAdapter.Fill(dataSet, "VozdPr");

			//cb1_dlin.Items.Clear();

			//for (int i = 0; i < dataSet.Tables["VozdPr"].Rows.Count; i++)
			//{
			//	cb1_dlin.Items.Add(dataSet.Tables["VozdPr"].Rows[i][1].ToString());
			//}

			//cb1_dlin.DataSource = dataSet;
			//cb1_dlin.DisplayMember = "dlin";
			//cb1_dlin.ValueMember = "id";


		}
		


		ISldWorks iSwApp = null;
		IModelDoc2 Part; //Предоставляет доступ к документам SOLIDWORKS: деталям, сборкам и чертежам.
						 //ShellFeatureData swShell;
		private void dataTab()
		{
			//прям возд
			cb51_dlin.Items.Clear();
			cb51_dlin.Items.Add("100");
			cb51_dlin.Items.Add("150");
			cb51_dlin.Items.Add("250");
			cb51_dlin.Items.Add("400");
			cb51_dlin.Items.Add("500");
			cb51_dlin.Items.Add("600");
			cb51_dlin.Items.Add("800");
			cb51_dlin.Items.Add("1000");
			cb51_dlin.Items.Add("1250");
			cb51_dlin.Items.Add("1600");

			cb51_shir.Items.Clear();
			cb51_shir.Items.Add("150");
			cb51_shir.Items.Add("250");
			cb51_shir.Items.Add("300");
			cb51_shir.Items.Add("400");
			cb51_shir.Items.Add("500");
			cb51_shir.Items.Add("600");
			cb51_shir.Items.Add("800");
			cb51_shir.Items.Add("1000");
			cb51_shir.Items.Add("1250");
			cb51_shir.Items.Add("1600");
			cb51_shir.Items.Add("2000");

			cb51_vys.Items.Clear();
			cb51_vys.Items.Add("20");
			cb51_vys.Items.Add("50");
			cb51_vys.Items.Add("100");
			
			
		}

		



		private void bt_PrymVoz_Click(object sender, EventArgs e)
		{

			double shir_pr, vys_pr, dlin_pr, zazor_pr;
			dlin_pr = Convert.ToDouble(cb51_dlin.Text);
			shir_pr = Convert.ToDouble(cb51_shir.Text);
			vys_pr = Convert.ToDouble(cb51_vys.Text);
			zazor_pr = Convert.ToDouble(cb1_zazor);

			shir_pr = shir_pr / 1000;
			vys_pr = vys_pr / 1000;
			dlin_pr = dlin_pr / 1000;
			zazor_pr = zazor_pr / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;
			
			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;
			

			airDuct duct = new airDuct();
			duct.createPrVozd(dlin_pr, shir_pr, vys_pr, zazor_pr, Part);


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

		//private void cb1_dlin_SelectedIndexChanged(object sender, EventArgs e)
		//{
		//	string str = cb1_dlin.SelectedItem.ToString();
		//	SqlConnection sqlConnection = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = D:\универ\airDucts\airDucts\Database1.mdf; Integrated Security = True");
		//	SqlCommand command = new SqlCommand("SELECT shir FROM VozdPr WHERE dlin = @str ", sqlConnection);

		//	command.Parameters.AddWithValue("@str", str);
		//	sqlDataAdapter2 = new SqlDataAdapter(command);

		//	dataSet2 = new DataSet();
		//	sqlDataAdapter2.Fill(dataSet2, "VozdPr");

		//	cb1_shir.Items.Clear();

		//	for (int i = 0; i < dataSet2.Tables["VozdPr"].Rows.Count; i++)
		//	{

		//		cb1_shir.Items.Add(dataSet2.Tables["VozdPr"].Rows[i][2].ToString());
		//	}

		//}

		private void cb1_dlin_SelectedValueChanged(object sender, EventArgs e)
		{
			
		}

		private void bt_ZaglPr_Click(object sender, EventArgs e)
		{
			double shir_pr, vys_pr, dlin_pr;
			dlin_pr = Convert.ToDouble(cb51_dlin.Text);
			shir_pr = Convert.ToDouble(cb51_shir.Text);
			vys_pr = Convert.ToDouble(cb51_vys.Text);
			

			shir_pr = shir_pr / 1000;
			vys_pr = vys_pr / 1000;
			dlin_pr = dlin_pr / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();

			Zaglushka zaglPr = new Zaglushka();
			zaglPr.createZaglPr(dlin_pr, shir_pr, vys_pr, Part);
		}

		private void bt_KrVozd_Click(object sender, EventArgs e)
		{
			double diam_kr, vys_kr, zazor_kr;
			diam_kr = Convert.ToDouble(cb2_diam.Text);
			vys_kr = Convert.ToDouble(cb2_vys.Text);
			zazor_kr = Convert.ToDouble(cb2_zazor.Text);


			diam_kr = diam_kr / 1000;
			vys_kr = vys_kr / 1000;
			zazor_kr = zazor_kr / 1000;

			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			airDuct duct2 = new airDuct();
			duct2.createKrVozd(diam_kr, vys_kr, zazor_kr, Part);
		}

		private void bt_PerehPrPr_Click(object sender, EventArgs e)
		{
			double dlin, dlin1, shir, shir1, vys, zazor;
			dlin = Convert.ToDouble(cb31_dlin.Text);
			dlin1 = Convert.ToDouble(cb31_dlin1.Text);
			shir = Convert.ToDouble(cb31_shir.Text);
			shir1 = Convert.ToDouble(cb31_shir1.Text);
			vys = Convert.ToDouble(cb31_vys.Text);
			zazor = Convert.ToDouble(cb31_zazor.Text);


			dlin = dlin / 1000;
			dlin1 = dlin1 / 1000;
			shir = shir / 1000;
			shir1 = shir1 / 1000;
			vys = vys / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();

			Perehod perehod1 = new Perehod();
			perehod1.createPrPrPerehod(dlin, dlin1, shir, shir1, vys, zazor, Part);
		}

		private void bt_PerehPr_Click(object sender, EventArgs e)
		{
			double dlin, shir, diam, vys, zazor;
			dlin = Convert.ToDouble(cb32_dlin.Text);
			shir = Convert.ToDouble(cb32_shir.Text);
			diam = Convert.ToDouble(cb32_diam.Text);
			vys = Convert.ToDouble(cb32_vys.Text);
			zazor = Convert.ToDouble(cb32_zazor.Text);


			dlin = dlin / 1000;
			shir = shir / 1000;
			diam = diam / 1000;
			vys = vys / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Perehod perehod2 = new Perehod();
			perehod2.createPrPerehod(dlin, shir, diam, vys, zazor, Part);
		}

		private void bt_PerehKr_Click(object sender, EventArgs e)
		{
			double diam1, diam2, vys, zazor;
			diam1 = Convert.ToDouble(cb33_diam1.Text);
			diam2 = Convert.ToDouble(cb33_diam2.Text);
			vys = Convert.ToDouble(cb33_vys.Text);
			zazor = Convert.ToDouble(cb33_zazor.Text);


			diam1 = diam1 / 1000;
			diam2 = diam2 / 1000;
			vys = vys / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Perehod perehod3 = new Perehod();
			perehod3.createKrPerehod(diam1, diam2, vys, zazor, Part);
		}

		private void bt_OtvodKr_Click(object sender, EventArgs e)
		{
			double diam,  zazor;
			diam = Convert.ToDouble(cb42_diam.Text);
			zazor = Convert.ToDouble(cb42_zazor.Text);


			diam = diam / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			//iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Otvod otvod1 = new Otvod();
			otvod1.createKrOtvod(diam,zazor,Part);

			//Part.SaveAs3("D:\универ\диплом доки\интерфейс\Деталь1.SLDPRT", 0, 0);

		}

		private void bt_OtvodPr_Click(object sender, EventArgs e)
		{
			double dlin,shir, dlinOtv, zazor;
			dlin = Convert.ToDouble(cb41_dlin.Text);
			shir = Convert.ToDouble(cb41_shir.Text);
			dlinOtv = Convert.ToDouble(cb41_dlinOtvod.Text);
			zazor = Convert.ToDouble(cb41_zazor.Text);


			dlin = dlin / 1000;
			shir = shir / 1000;
			dlinOtv = dlinOtv / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Otvod otvod2 = new Otvod();
			otvod2.createPrOtvod(dlin, shir, dlinOtv, zazor, Part);
		}
	}
}
