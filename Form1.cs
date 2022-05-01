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
		private SqlDataAdapter sqlDataAdapter3 = null;
		private DataSet dataSet = null;
		private DataSet dataSet2 = null;
		private DataSet dataSet3 = null;

		public Form1()
		{
			InitializeComponent();

			dataTab();
			textBox2.Text = "Переход с прямоугольного на прямоугольное сечение является фасонным элементом, позволяющим соединить воздуховоды с разным размером прямоугольных сечений.";

			
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
			dlin_pr = Convert.ToDouble(cb1_dlin.Text);
			shir_pr = Convert.ToDouble(cb1_shir.Text);
			vys_pr = Convert.ToDouble(cb1_vys.Text);
			zazor_pr = Convert.ToDouble(cb1_zazor.Text);

			shir_pr = shir_pr / 1000;
			vys_pr = vys_pr / 1000;
			dlin_pr = dlin_pr / 1000;
			zazor_pr = zazor_pr / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;
			Close();

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

		private void cb1_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			string str = cb1_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT  shir FROM vozdPr WHERE dlin = @str ", sqlConnection);

			command.Parameters.AddWithValue("@str", str);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdPr");

			cb1_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["vozdPr"].Rows.Count; i++)
			{

				cb1_shir.Items.Add(dataSet2.Tables["vozdPr"].Rows[i]["shir"].ToString());
			}



		}

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

			iSwApp.NewPart();
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

			iSwApp.NewPart();
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

			iSwApp.NewPart();
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

			iSwApp.NewPart();
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

			iSwApp.NewPart();
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

			iSwApp.NewPart();
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

		private void Form1_Load(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT DISTINCT dlin  FROM vozdPr", sqlConnection);

			DataSet dataSet = new DataSet();
			sqlDataAdapter.Fill(dataSet, "vozdPr");

			cb1_dlin.Items.Clear();
			cb1_vys.Items.Clear();

			for (int i = 0; i < dataSet.Tables["vozdPr"].Rows.Count; i++)
			{
				cb1_dlin.Items.Add(dataSet.Tables["vozdPr"].Rows[i]["dlin"].ToString());
				
			}

			SqlDataAdapter sqlDataAdapter2 = new SqlDataAdapter("SELECT vys  FROM vozdPr WHERE vys IS NOT NULL", sqlConnection);

			DataSet dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdPr");
			for (int i = 0; i < dataSet2.Tables["vozdPr"].Rows.Count; i++)
			{
				cb1_vys.Items.Add(dataSet2.Tables["vozdPr"].Rows[i]["vys"].ToString());

			}

			SqlDataAdapter sqlDataAdapter3 = new SqlDataAdapter("SELECT  diam  FROM vozdKr", sqlConnection);

			DataSet dataSet3 = new DataSet();
			sqlDataAdapter3.Fill(dataSet3, "vozdKr");

			cb2_diam.Items.Clear();
			cb2_vys.Items.Clear();

			for (int i = 0; i < dataSet3.Tables["vozdKr"].Rows.Count; i++)
			{
				cb2_diam.Items.Add(dataSet3.Tables["vozdKr"].Rows[i]["diam"].ToString());

			}

			SqlDataAdapter sqlDataAdapter4 = new SqlDataAdapter("SELECT vys  FROM vozdKr WHERE vys IS NOT NULL", sqlConnection);

			DataSet dataSet4 = new DataSet();
			sqlDataAdapter4.Fill(dataSet4, "vozdKr");
			for (int i = 0; i < dataSet4.Tables["vozdKr"].Rows.Count; i++)
			{
				cb2_vys.Items.Add(dataSet4.Tables["vozdKr"].Rows[i]["vys"].ToString());

			}

			SqlDataAdapter sqlDataAdapter5 = new SqlDataAdapter("SELECT DISTINCT  dlin1  FROM perehodPrPr", sqlConnection);

			DataSet dataSet5 = new DataSet();
			sqlDataAdapter5.Fill(dataSet5, "perehodPrPr");

			cb31_dlin.Items.Clear();

			for (int i = 0; i < dataSet5.Tables["perehodPrPr"].Rows.Count; i++)
			{
				cb31_dlin.Items.Add(dataSet5.Tables["perehodPrPr"].Rows[i]["dlin1"].ToString());

			}

			SqlDataAdapter sqlDataAdapter6 = new SqlDataAdapter("SELECT DISTINCT  dlin1  FROM perehodPrPr", sqlConnection);

			DataSet dataSet6 = new DataSet();
			sqlDataAdapter5.Fill(dataSet5, "perehodPrPr");

			cb31_dlin.Items.Clear();

			for (int i = 0; i < dataSet5.Tables["perehodPrPr"].Rows.Count; i++)
			{
				cb31_dlin.Items.Add(dataSet5.Tables["perehodPrPr"].Rows[i]["dlin1"].ToString());

			}
		}

		private void bt_TroinicPr_Click(object sender, EventArgs e)
		{
			double dlin1, shir1, dlin2, shir2, vys1, vys2, zazor;
			dlin1 = Convert.ToDouble(cb61_dlin1.Text);
			shir1 = Convert.ToDouble(cb61_shir1.Text);
			dlin2 = Convert.ToDouble(cb61_dlin2.Text);
			shir2 = Convert.ToDouble(cb61_shir2.Text);
			vys1 = Convert.ToDouble(cb61_vys1.Text);
			vys2 = Convert.ToDouble(cb61_vys2.Text);
			zazor = Convert.ToDouble(cb61_zazor.Text);


			dlin1 = dlin1 / 1000;
			shir1 = shir1 / 1000;
			dlin2 = dlin2 / 1000;
			shir2 = shir2 / 1000;
			vys1 = vys1 / 1000;
			vys2 = vys2 / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Troinik troinik = new Troinik();
			troinik.createTroinikPr(dlin1, shir1, vys1, dlin2, shir2, vys2, zazor, Part);


		}

		private void bt_TroinicPrKr_Click(object sender, EventArgs e)
		{
			double dlin, shir, vys1, diam,  vys2, zazor;
			dlin = Convert.ToDouble(cb62_dlin.Text);
			shir = Convert.ToDouble(cb62_shir.Text);
			vys1 = Convert.ToDouble(cb62_vys1.Text);
			diam = Convert.ToDouble(cb62_diam.Text);
			vys2 = Convert.ToDouble(cb62_vys2.Text);
			zazor = Convert.ToDouble(cb62_zazor.Text);


			dlin = dlin / 1000;
			shir = shir / 1000;
			diam = diam / 1000;
			vys1 = vys1 / 1000;
			vys2 = vys2 / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Troinik troinik = new Troinik();
			troinik.createTroinikPrKr(dlin, shir, vys1, diam, vys2, zazor, Part);
		}

		private void bt_TroinicKr_Click(object sender, EventArgs e)
		{
			double diam1, diam2,vys1, vys2, zazor;
			diam1 = Convert.ToDouble(cb63_diam1.Text);
			vys1 = Convert.ToDouble(cb63_vys1.Text);
			diam2 = Convert.ToDouble(cb63_diam2.Text);
			vys2 = Convert.ToDouble(cb63_vys2.Text);
			zazor = Convert.ToDouble(cb63_zazor.Text);


			diam1 = diam1 / 1000;
			diam2 = diam2 / 1000;
			vys1 = vys1 / 1000;
			vys2 = vys2 / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Troinik troinik = new Troinik();
			troinik.createTroinikKr(diam1,  vys1, diam2, vys2, zazor, Part);
		}

		private void bt_TroinicKrPr_Click(object sender, EventArgs e)
		{
			double dlin, shir, vys1, diam, vys2, zazor;
			dlin = Convert.ToDouble(cb64_dlin.Text);
			shir = Convert.ToDouble(cb64_shir.Text);
			vys1 = Convert.ToDouble(cb64_vys1.Text);
			diam = Convert.ToDouble(cb64_diam.Text);
			vys2 = Convert.ToDouble(cb64_vys2.Text);
			zazor = Convert.ToDouble(cb64_zazor.Text);


			dlin = dlin / 1000;
			shir = shir / 1000;
			diam = diam / 1000;
			vys1 = vys1 / 1000;
			vys2 = vys2 / 1000;
			zazor = zazor / 1000;


			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			iSwApp.NewPart();
			Part = iSwApp.IActiveDoc2;

			Close();
			iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

			Troinik troinik = new Troinik();
			troinik.createTroinikKrPr(dlin, shir, vys1, diam, vys2, zazor, Part);
		}

		private void cb1_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			string str1 = cb1_dlin.SelectedItem.ToString();
			string str2 = cb1_shir.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT thickness FROM vozdPr WHERE shir = @str2 AND dlin = @str1 ", sqlConnection);

			command.Parameters.AddWithValue("@str1", str1);
			command.Parameters.AddWithValue("@str2", str2);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdPr");

			cb1_zazor.Clear();

			for (int i = 0; i < dataSet2.Tables["vozdPr"].Rows.Count; i++)
			{
				cb1_zazor.Text = dataSet2.Tables["vozdPr"].Rows[i]["thickness"].ToString();
			}
		}

		private void cb2_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			string str1 = cb2_diam.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT thick FROM vozdKr WHERE diam = @str1  ", sqlConnection);

			command.Parameters.AddWithValue("@str1", str1);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdKr");

			cb2_zazor.Clear();

			for (int i = 0; i < dataSet2.Tables["vozdKr"].Rows.Count; i++)
			{
				cb2_zazor.Text = dataSet2.Tables["vozdKr"].Rows[i]["thick"].ToString();
			}
		}

		private void cb31_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			string str = cb31_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT  shir1 FROM perehodPrPr WHERE dlin1 = @str ", sqlConnection);

			command.Parameters.AddWithValue("@str", str);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{

				cb31_shir.Items.Add(dataSet2.Tables["perehodPrPr"].Rows[i]["shir1"].ToString());
			}

		}

		private void cb31_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dlin2 FROM perehodPrPr WHERE dlin1 = "+ cb31_dlin.SelectedItem.ToString() +" AND shir1 = " + cb31_shir.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_dlin1.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{

				cb31_dlin1.Items.Add(dataSet2.Tables["perehodPrPr"].Rows[i]["dlin2"].ToString());
			}
		}

		private void cb31_dlin1_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir2,vys FROM perehodPrPr WHERE dlin1 = " + cb31_dlin.SelectedItem.ToString() + " AND shir1 = " + cb31_shir.SelectedItem.ToString() + " AND dlin2 = " + cb31_dlin1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_shir1.Items.Clear();
			cb31_vys.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{

				cb31_shir1.Items.Add(dataSet2.Tables["perehodPrPr"].Rows[i]["shir2"].ToString());
				cb31_vys.Items.Add(dataSet2.Tables["perehodPrPr"].Rows[i]["vys"].ToString());
			}

			
		}

		private void cb31_shir1_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS;Initial Catalog=airDuct;Integrated Security=True");
			

			SqlCommand command1 = new SqlCommand("SELECT DISTINCT thick FROM perehodPrPr WHERE shir1 = " + cb31_shir.SelectedItem.ToString() + " AND dlin2 = " + cb31_dlin1.SelectedItem.ToString() + " AND shir2 = " + cb31_shir1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command1);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_zazor.Clear();
			

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{

				cb31_zazor.Text = dataSet2.Tables["perehodPrPr"].Rows[i]["thick"].ToString();
			}
		}
	}
}
