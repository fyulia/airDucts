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
		private DataSet dataSet = null;

		public Form1()
		{
			InitializeComponent();
			
			//dataTab();
			textBox2.Text = "Переход с прямоугольного на прямоугольное сечение является фасонным элементом, позволяющим соединить воздуховоды с разным размером прямоугольных сечений.";

			SqlConnection sqlConnection = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = D:\универ\airDucts\airDucts\Database1.mdf; Integrated Security = True");
			SqlDataAdapter da = new SqlDataAdapter("SELECT DISTINCT dlin FROM VozdPr ", sqlConnection);
			
			DataSet dataSet = new DataSet();
			da.Fill(dataSet, "VozdPr");

			cb1_dlin.Items.Clear();
			
			for (int i = 0; i < dataSet.Tables["VozdPr"].Rows.Count; i++)
			{
					cb1_dlin.Items.Add(dataSet.Tables["VozdPr"].Rows[i][0].ToString());
			}
			
		
		
		}
		


		ISldWorks iSwApp = null;
		IModelDoc2 Part; //Предоставляет доступ к документам SOLIDWORKS: деталям, сборкам и чертежам.
						 //ShellFeatureData swShell;
		private void dataTab()
		{
			//прям возд
			cb1_dlin.Items.Clear();
			cb1_dlin.Items.Add("100");
			cb1_dlin.Items.Add("150");
			cb1_dlin.Items.Add("250");
			cb1_dlin.Items.Add("400");
			cb1_dlin.Items.Add("500");
			cb1_dlin.Items.Add("600");
			cb1_dlin.Items.Add("800");
			cb1_dlin.Items.Add("1000");
			cb1_dlin.Items.Add("1250");
			cb1_dlin.Items.Add("1600");

			cb1_shir.Items.Clear();
			cb1_shir.Items.Add("150");
			cb1_shir.Items.Add("250");
			cb1_shir.Items.Add("300");
			cb1_shir.Items.Add("400");
			cb1_shir.Items.Add("500");
			cb1_shir.Items.Add("600");
			cb1_shir.Items.Add("800");
			cb1_shir.Items.Add("1000");
			cb1_shir.Items.Add("1250");
			cb1_shir.Items.Add("1600");
			cb1_shir.Items.Add("2000");

			cb1_vys.Items.Clear();
			cb1_vys.Items.Add("2500");
			cb1_vys.Items.Add("3000");
			cb1_vys.Items.Add("4000");
			cb1_vys.Items.Add("5000");
			cb1_vys.Items.Add("6000");
		}

		



		private void bt_PrymVoz_Click(object sender, EventArgs e)
		{

			double shir_pr, vys_pr, dlin_pr;
			shir_pr = Convert.ToDouble(cb1_dlin.Text);
			vys_pr = Convert.ToDouble(cb1_shir.Text);
			dlin_pr = Convert.ToDouble(cb1_vys.Text);

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

		private void cb1_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			
		}

		private void cb1_dlin_SelectedValueChanged(object sender, EventArgs e)
		{
			string str = cb1_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = D:\универ\airDucts\airDucts\Database1.mdf; Integrated Security = True");
			SqlCommand command = new SqlCommand("SELECT shir FROM VozdPr WHERE dlin = @str ", sqlConnection);

			command.Parameters.AddWithValue("@str", str);
			sqlDataAdapter = new SqlDataAdapter(command);

			dataSet = new DataSet();
			sqlDataAdapter.Fill(dataSet, "VozdPr");

			cb1_shir.Items.Clear();

			for (int i = 0; i < dataSet.Tables["VozdPr"].Rows.Count; i++)
			{

				cb1_shir.Items.Add(dataSet.Tables["VozdPr"].Rows[i][1].ToString());
			}
		}
	}
}
