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
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Runtime.InteropServices;
using System.IO;
//using System.Windows;

namespace airDucts
{
	public partial class Interface : Form
	{

		private SqlConnection sqlConnection = null;
		private SqlDataAdapter sqlDataAdapter = null;
		private SqlDataAdapter sqlDataAdapter2 = null;
		private SqlDataAdapter sqlDataAdapter3 = null;
		private DataSet dataSet = null;
		private DataSet dataSet2 = null;
		private DataSet dataSet3 = null;

		int longstatus;
		bool boolstatus;
		int lErrors;
		int lWarnings;

		string assPut;
			string assPut2;
		string nameActiveAss;
		public Interface()
		{
			InitializeComponent();	

			textBox2.Text = "Переход с прямоугольного на прямоугольное сечение является фасонным элементом, позволяющим соединить воздуховоды с разным размером прямоугольных сечений.";

			svoistva();
			

			iSwApp = new SldWorks();
			iSwApp.Visible = true;

			IModelDoc2 Part = iSwApp.IActiveDoc2;
			assPut = Part.GetPathName();
			 nameActiveAss = Path.GetFileNameWithoutExtension(assPut);

			string name = Part.GetTitle();
			//путь где хранится сброка
			assPut2 = assPut.Substring(0, assPut.Length - name.Length - 1);
			//MessageBox.Show(assPut2);
		}

		public void svoistva()
		{
			cb1_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb1_shir.KeyPress += (sender, e) => e.Handled = true;
			cb1_vys.KeyPress += (sender, e) => e.Handled = true;

			cb2_diam.KeyPress += (sender, e) => e.Handled = true;
			cb2_vys.KeyPress += (sender, e) => e.Handled = true;

			cb31_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb31_dlin1.KeyPress += (sender, e) => e.Handled = true;
			cb31_shir.KeyPress += (sender, e) => e.Handled = true;
			cb31_shir1.KeyPress += (sender, e) => e.Handled = true;

			cb32_diam.KeyPress += (sender, e) => e.Handled = true;
			cb32_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb32_shir.KeyPress += (sender, e) => e.Handled = true;

			cb33_diam1.KeyPress += (sender, e) => e.Handled = true;
			cb33_diam2.KeyPress += (sender, e) => e.Handled = true;

			cb41_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb41_shir.KeyPress += (sender, e) => e.Handled = true;

			cb42_diam.KeyPress += (sender, e) => e.Handled = true;

			cb51_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb51_shir.KeyPress += (sender, e) => e.Handled = true;

			cb52_diam.KeyPress += (sender, e) => e.Handled = true;

			cb61_dlin1.KeyPress += (sender, e) => e.Handled = true;
			cb61_dlin2.KeyPress += (sender, e) => e.Handled = true;
			cb61_shir1.KeyPress += (sender, e) => e.Handled = true;
			cb61_shir2.KeyPress += (sender, e) => e.Handled = true;

			cb62_diam.KeyPress += (sender, e) => e.Handled = true;
			cb62_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb62_shir.KeyPress += (sender, e) => e.Handled = true;

			cb63_diam1.KeyPress += (sender, e) => e.Handled = true;
			cb63_diam2.KeyPress += (sender, e) => e.Handled = true;

			cb64_diam.KeyPress += (sender, e) => e.Handled = true;
			cb64_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb64_shir.KeyPress += (sender, e) => e.Handled = true;

			cb71_dlin.KeyPress += (sender, e) => e.Handled = true;
			cb71_shir.KeyPress += (sender, e) => e.Handled = true;

			cb72_diam.KeyPress += (sender, e) => e.Handled = true;

		}

		ISldWorks iSwApp = null;
		IModelDoc2 Part; //Предоставляет доступ к документам SOLIDWORKS: деталям, сборкам и чертежам.
						 //ShellFeatureData swShell;


		IAssemblyDoc assembly;
		
		string tmpPath = null;

		
		int errors = 0;
		int warnings = 0;
		ModelDocExtension swModelDocExt = default(ModelDocExtension);
		Feature swFeature;
		FeatureManager swFeatureManager = default(FeatureManager);
		CircularPatternFeatureData swFeatData;
		Configuration swConfig = default(Configuration);
		SketchSegment skSegment;
		//данные


		private void Form1_Load(object sender, EventArgs e)
		{			
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
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

			SqlDataAdapter sqlDataAdapter6 = new SqlDataAdapter("SELECT DISTINCT  dlin  FROM perehodPr", sqlConnection);

			DataSet dataSet6 = new DataSet();
			sqlDataAdapter6.Fill(dataSet6, "perehodPr");

			cb32_dlin.Items.Clear();

			for (int i = 0; i < dataSet6.Tables["perehodPr"].Rows.Count; i++)
			{
				cb32_dlin.Items.Add(dataSet6.Tables["perehodPr"].Rows[i]["dlin"].ToString());
			}

			SqlDataAdapter sqlDataAdapter7 = new SqlDataAdapter("SELECT DISTINCT  diam1  FROM perehodKr", sqlConnection);

			DataSet dataSet7 = new DataSet();
			sqlDataAdapter7.Fill(dataSet7, "perehodKr");

			cb33_diam1.Items.Clear();

			for (int i = 0; i < dataSet7.Tables["perehodKr"].Rows.Count; i++)
			{
				cb33_diam1.Items.Add(dataSet7.Tables["perehodKr"].Rows[i]["diam1"].ToString());
			}

			SqlDataAdapter sqlDataAdapter8 = new SqlDataAdapter("SELECT DISTINCT  dlin  FROM otvodPr", sqlConnection);

			DataSet dataSet8 = new DataSet();
			sqlDataAdapter8.Fill(dataSet8, "otvodPr");

			cb41_dlin.Items.Clear();

			for (int i = 0; i < dataSet8.Tables["otvodPr"].Rows.Count; i++)
			{
				cb41_dlin.Items.Add(dataSet8.Tables["otvodPr"].Rows[i]["dlin"].ToString());
			}

			SqlDataAdapter sqlDataAdapter9 = new SqlDataAdapter("SELECT diam  FROM otvodKr", sqlConnection);

			DataSet dataSet9 = new DataSet();
			sqlDataAdapter9.Fill(dataSet9, "otvodKr");

			cb42_diam.Items.Clear();

			for (int i = 0; i < dataSet9.Tables["otvodKr"].Rows.Count; i++)
			{
				cb42_diam.Items.Add(dataSet9.Tables["otvodKr"].Rows[i]["diam"].ToString());
			}

			SqlDataAdapter sqlDataAdapter10 = new SqlDataAdapter("SELECT DISTINCT dlin  FROM zaglPr", sqlConnection);

			DataSet dataSet10 = new DataSet();
			sqlDataAdapter10.Fill(dataSet10, "zaglPr");

			cb51_dlin.Items.Clear();

			for (int i = 0; i < dataSet10.Tables["zaglPr"].Rows.Count; i++)
			{
				cb51_dlin.Items.Add(dataSet10.Tables["zaglPr"].Rows[i]["dlin"].ToString());
			}

			SqlDataAdapter sqlDataAdapter11 = new SqlDataAdapter("SELECT  diam  FROM zaglKr", sqlConnection);

			DataSet dataSet11 = new DataSet();
			sqlDataAdapter11.Fill(dataSet11, "zaglKr");

			cb52_diam.Items.Clear();

			for (int i = 0; i < dataSet11.Tables["zaglKr"].Rows.Count; i++)
			{
				cb52_diam.Items.Add(dataSet11.Tables["zaglKr"].Rows[i]["diam"].ToString());
			}

			SqlDataAdapter sqlDataAdapter12 = new SqlDataAdapter("SELECT DISTINCT dlin1  FROM troinicPrPr", sqlConnection);

			DataSet dataSet12 = new DataSet();
			sqlDataAdapter12.Fill(dataSet12, "troinicPrPr");

			cb61_dlin1.Items.Clear();

			for (int i = 0; i < dataSet12.Tables["troinicPrPr"].Rows.Count; i++)
			{
				cb61_dlin1.Items.Add(dataSet12.Tables["troinicPrPr"].Rows[i]["dlin1"].ToString());
			}

			SqlDataAdapter sqlDataAdapter13 = new SqlDataAdapter("SELECT DISTINCT dlin  FROM troinicPrKr", sqlConnection);

			DataSet dataSet13 = new DataSet();
			sqlDataAdapter13.Fill(dataSet13, "troinicPrKr");

			cb62_dlin.Items.Clear();

			for (int i = 0; i < dataSet13.Tables["troinicPrKr"].Rows.Count; i++)
			{
				cb62_dlin.Items.Add(dataSet13.Tables["troinicPrKr"].Rows[i]["dlin"].ToString());

			}

			SqlDataAdapter sqlDataAdapter14 = new SqlDataAdapter("SELECT DISTINCT diam1  FROM troinicKrKr", sqlConnection);

			DataSet dataSet14 = new DataSet();
			sqlDataAdapter14.Fill(dataSet14, "troinicKrKr");

			cb63_diam1.Items.Clear();

			for (int i = 0; i < dataSet14.Tables["troinicKrKr"].Rows.Count; i++)
			{
				cb63_diam1.Items.Add(dataSet14.Tables["troinicKrKr"].Rows[i]["diam1"].ToString());

			}

			SqlDataAdapter sqlDataAdapter15 = new SqlDataAdapter("SELECT DISTINCT diam  FROM troinicKrPr", sqlConnection);

			DataSet dataSet15 = new DataSet();
			sqlDataAdapter15.Fill(dataSet15, "troinicKrPr");

			cb64_diam.Items.Clear();

			for (int i = 0; i < dataSet15.Tables["troinicKrPr"].Rows.Count; i++)
			{
				cb64_diam.Items.Add(dataSet15.Tables["troinicKrPr"].Rows[i]["diam"].ToString());

			}

			SqlDataAdapter sqlDataAdapter16 = new SqlDataAdapter("SELECT DISTINCT dlin1  FROM flanPr", sqlConnection);

			DataSet dataSet16 = new DataSet();
			sqlDataAdapter16.Fill(dataSet16, "flanPr");

			cb71_dlin.Items.Clear();

			for (int i = 0; i < dataSet16.Tables["flanPr"].Rows.Count; i++)
			{
				cb71_dlin.Items.Add(dataSet16.Tables["flanPr"].Rows[i]["dlin1"].ToString());

			}

			SqlDataAdapter sqlDataAdapter17 = new SqlDataAdapter("SELECT DISTINCT diam  FROM flanKr", sqlConnection);

			DataSet dataSet17 = new DataSet();
			sqlDataAdapter17.Fill(dataSet17, "flanKr");

			cb72_diam.Items.Clear();

			for (int i = 0; i < dataSet17.Tables["flanKr"].Rows.Count; i++)
			{
				cb72_diam.Items.Add(dataSet17.Tables["flanKr"].Rows[i]["diam"].ToString());

			}
		}



		private void cb1_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			
			cb1_shir.Enabled = true;
			cb1_shir.Text = "";


			string str = cb1_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
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

		private void cb1_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb1_vys.Enabled = true;

			

			string str1 = cb1_dlin.SelectedItem.ToString();
			string str2 = cb1_shir.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT thickness FROM vozdPr WHERE shir = @str2 AND dlin = @str1 ", sqlConnection);

			command.Parameters.AddWithValue("@str1", str1);
			command.Parameters.AddWithValue("@str2", str2);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdPr");

			cb1_zazor.Clear();
			cb1_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["vozdPr"].Rows.Count; i++)
			{
				cb1_zazor.Text = dataSet2.Tables["vozdPr"].Rows[i]["thickness"].ToString();
				cb1_thick.Text = dataSet2.Tables["vozdPr"].Rows[i]["thickness"].ToString();
			}
		}

		private void cb2_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb2_vys.Text = "";
			string str1 = cb2_diam.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT thick FROM vozdKr WHERE diam = @str1  ", sqlConnection);

			command.Parameters.AddWithValue("@str1", str1);
			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "vozdKr");

			cb2_zazor.Clear();
			cb2_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["vozdKr"].Rows.Count; i++)
			{
				cb2_zazor.Text = dataSet2.Tables["vozdKr"].Rows[i]["thick"].ToString();
				cb2_thick.Text = dataSet2.Tables["vozdKr"].Rows[i]["thick"].ToString();
			}
		}

		private void cb31_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb31_shir.Enabled = true;
			cb31_shir.Text = "";
			cb31_dlin1.Text = "";
			cb31_shir1.Text = "";
			cb31_vys.Clear();
			cb31_zazor.Clear();
			cb31_thick.Clear();
			string str = cb31_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir1 FROM perehodPrPr WHERE dlin1 = @str ", sqlConnection);

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
			cb31_dlin1.Enabled = true;
			
			cb31_dlin1.Text = "";
			cb31_shir1.Text = "";
			cb31_vys.Clear();
			cb31_zazor.Clear();
			cb31_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dlin2 FROM perehodPrPr WHERE dlin1 = " + cb31_dlin.SelectedItem.ToString() + " " +
				"AND shir1 = " + cb31_shir.SelectedItem.ToString() + " ", sqlConnection);

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
			cb31_shir1.Enabled = true;
			
			cb31_shir1.Text = "";
			cb31_vys.Clear();
			cb31_zazor.Clear();
			cb31_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir2,vys FROM perehodPrPr WHERE dlin1 = " + cb31_dlin.SelectedItem.ToString() + " AND shir1 = " + cb31_shir.SelectedItem.ToString() + " AND dlin2 = " + cb31_dlin1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_shir1.Items.Clear();
			cb31_vys.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{
				cb31_shir1.Items.Add(dataSet2.Tables["perehodPrPr"].Rows[i]["shir2"].ToString());
				cb31_vys.Text=(dataSet2.Tables["perehodPrPr"].Rows[i]["vys"].ToString());
			}
		}

		private void cb31_shir1_SelectedIndexChanged(object sender, EventArgs e)
		{
			
			
			cb31_zazor.Clear();
			cb31_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command1 = new SqlCommand("SELECT DISTINCT thick FROM perehodPrPr WHERE shir1 = " + cb31_shir.SelectedItem.ToString() + " AND dlin2 = " + cb31_dlin1.SelectedItem.ToString() + " AND shir2 = " + cb31_shir1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command1);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPrPr");

			cb31_zazor.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPrPr"].Rows.Count; i++)
			{
				cb31_zazor.Text = dataSet2.Tables["perehodPrPr"].Rows[i]["thick"].ToString();
				cb31_thick.Text = dataSet2.Tables["perehodPrPr"].Rows[i]["thick"].ToString();
			}
		}

		private void cb32_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb32_shir.Enabled = true;
			cb32_shir.Text = "";
			cb32_diam.Text = "";
			cb32_vys.Clear();
			cb32_zazor.Clear();
			cb32_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir FROM perehodPr WHERE dlin = " + cb32_dlin.SelectedItem.ToString() + "  ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPr");

			cb32_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPr"].Rows.Count; i++)
			{
				cb32_shir.Items.Add(dataSet2.Tables["perehodPr"].Rows[i]["shir"].ToString());
			}
		}

		private void cb32_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb32_diam.Enabled = true;
		
			cb32_diam.Text = "";
			cb32_vys.Clear();
			cb32_zazor.Clear(); 
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT diam,thick FROM perehodPr WHERE dlin = " + cb32_dlin.SelectedItem.ToString() + " AND shir = " + cb32_shir.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPr");

			cb32_diam.Items.Clear();
			cb32_zazor.Clear();
			cb32_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPr"].Rows.Count; i++)
			{
				cb32_diam.Items.Add(dataSet2.Tables["perehodPr"].Rows[i]["diam"].ToString());
				cb32_zazor.Text = (dataSet2.Tables["perehodPr"].Rows[i]["thick"].ToString());
				cb32_thick.Text = (dataSet2.Tables["perehodPr"].Rows[i]["thick"].ToString());
			}
		}

		private void cb32_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT vys FROM perehodPr WHERE dlin = " + cb32_dlin.SelectedItem.ToString() + " AND shir = " + cb32_shir.SelectedItem.ToString() + " AND diam = " + cb32_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodPr");

			cb32_vys.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodPr"].Rows.Count; i++)
			{
				cb32_vys.Text = (dataSet2.Tables["perehodPr"].Rows[i]["vys"].ToString());
			}
		}

		private void cb33_diam1_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb33_diam2.Enabled = true;
			cb33_diam2.Text = "";
			cb33_vys.Clear();
			cb33_zazor.Clear();
			cb33_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT diam2 FROM perehodKr WHERE diam1 = " + cb33_diam1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodKr");

			cb33_diam2.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodKr"].Rows.Count; i++)
			{
				cb33_diam2.Items.Add(dataSet2.Tables["perehodKr"].Rows[i]["diam2"].ToString());
			}
		}

		private void cb33_diam2_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT vys,thick FROM perehodKr WHERE diam1 = " + cb33_diam1.SelectedItem.ToString() + " AND diam2 = " + cb33_diam2.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "perehodKr");

			cb33_vys.Clear();
			cb33_zazor.Clear();
			cb33_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["perehodKr"].Rows.Count; i++)
			{
				cb33_vys.Text = (dataSet2.Tables["perehodKr"].Rows[i]["vys"].ToString());
				cb33_zazor.Text = (dataSet2.Tables["perehodKr"].Rows[i]["thick"].ToString());
				cb33_thick.Text = (dataSet2.Tables["perehodKr"].Rows[i]["thick"].ToString());
			}
		}

		private void cb41_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb41_shir.Enabled = true;
			cb41_shir.Text = "";
			cb41_dlinOtvod.Clear();
			cb41_zazor.Clear();
			cb41_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir FROM otvodPr WHERE dlin = " + cb41_dlin.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "otvodPr");

			cb41_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["otvodPr"].Rows.Count; i++)
			{
				cb41_shir.Items.Add(dataSet2.Tables["otvodPr"].Rows[i]["shir"].ToString());
			}
		}

		private void cb41_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dlinOtv,thick FROM otvodPr WHERE dlin = " + cb41_dlin.SelectedItem.ToString() + " AND shir = " + cb41_shir.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "otvodPr");

			cb41_dlinOtvod.Clear();
			cb41_zazor.Clear();
			cb41_thick.Clear();
			for (int i = 0; i < dataSet2.Tables["otvodPr"].Rows.Count; i++)
			{

				cb41_dlinOtvod.Text = (dataSet2.Tables["otvodPr"].Rows[i]["dlinOtv"].ToString());
				cb41_zazor.Text = (dataSet2.Tables["otvodPr"].Rows[i]["thick"].ToString());
				cb41_thick.Text = (dataSet2.Tables["otvodPr"].Rows[i]["thick"].ToString());
			}
		}

		private void cb42_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick FROM otvodKr WHERE diam = " + cb42_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "otvodKr");

			cb42_zazor.Clear();
			cb42_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["otvodKr"].Rows.Count; i++)
			{
				cb42_zazor.Text = (dataSet2.Tables["otvodKr"].Rows[i]["thick"].ToString());
				cb42_thick.Text = (dataSet2.Tables["otvodKr"].Rows[i]["thick"].ToString());
			}
		}

		private void cb51_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb51_shir.Enabled = true;
			cb51_shir.Text = "";
			cb51_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir,thick FROM zaglPr WHERE dlin = " + cb51_dlin.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "zaglPr");

			cb51_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["zaglPr"].Rows.Count; i++)
			{
				cb51_shir.Items.Add(dataSet2.Tables["zaglPr"].Rows[i]["shir"].ToString());
				cb51_thick.Text = dataSet2.Tables["zaglPr"].Rows[i]["thick"].ToString();
			}
		}

		private void cb61_dlin1_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb61_shir1.Enabled = true;
			cb61_shir1.Text = "";
			cb61_dlin2.Text = "";
			cb61_shir2.Text = "";
			cb61_vys1.Clear();
			cb61_zazor.Clear();
			cb61_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir1 FROM troinicPrPr WHERE dlin1 = " + cb61_dlin1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrPr");

			cb61_shir1.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrPr"].Rows.Count; i++)
			{
				cb61_shir1.Items.Add(dataSet2.Tables["troinicPrPr"].Rows[i]["shir1"].ToString());
			}
		}

		private void cb61_shir1_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb61_dlin2.Enabled = true;
			
			cb61_dlin2.Text = "";
			cb61_shir2.Text = "";
			cb61_vys1.Clear();
			cb61_zazor.Clear();
			cb61_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dlin2 FROM troinicPrPr WHERE dlin1 = " + cb61_dlin1.SelectedItem.ToString() + " AND shir1 = " + cb61_shir1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrPr");

			cb61_dlin2.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrPr"].Rows.Count; i++)
			{
				cb61_dlin2.Items.Add(dataSet2.Tables["troinicPrPr"].Rows[i]["dlin2"].ToString());
			}
		}

		private void cb62_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb62_shir.Enabled = true;
			cb62_shir.Text = "";
			cb62_diam.Text = "";
			cb62_vys1.Clear();
			cb62_zazor.Clear();
			cb62_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir FROM troinicPrKr WHERE dlin = " + cb62_dlin.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrKr");

			cb62_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrKr"].Rows.Count; i++)
			{
				cb62_shir.Items.Add(dataSet2.Tables["troinicPrKr"].Rows[i]["shir"].ToString());
			}
		}

		private void cb62_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb62_diam.Enabled = true;
		
			cb62_diam.Text = "";
			cb62_vys1.Clear();
			cb62_zazor.Clear();
			cb62_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT diam FROM troinicPrKr WHERE dlin = " + cb62_dlin.SelectedItem.ToString() + " AND shir = " + cb62_shir.SelectedItem.ToString() + "", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrKr");

			cb62_diam.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrKr"].Rows.Count; i++)
			{
				cb62_diam.Items.Add(dataSet2.Tables["troinicPrKr"].Rows[i]["diam"].ToString());
			}
		}

		private void cb62_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick,vys FROM troinicPrKr WHERE dlin = " + cb62_dlin.SelectedItem.ToString() + " AND shir = " + cb62_shir.SelectedItem.ToString() + " AND diam = " + cb62_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrKr");

			cb62_vys1.Clear();
			cb62_zazor.Clear();
			cb62_thick.Clear();
			for (int i = 0; i < dataSet2.Tables["troinicPrKr"].Rows.Count; i++)
			{
				cb62_zazor.Text = (dataSet2.Tables["troinicPrKr"].Rows[i]["thick"].ToString());
				cb62_thick.Text = (dataSet2.Tables["troinicPrKr"].Rows[i]["thick"].ToString());
				cb62_vys1.Text = (dataSet2.Tables["troinicPrKr"].Rows[i]["vys"].ToString());

			}

		}

		private void cb63_diam1_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb63_diam2.Enabled = true;
			cb63_diam2.Text = "";
			cb63_vys1.Clear();
			cb63_zazor.Clear();
			cb63_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT diam2 FROM troinicKrKr WHERE diam1 = " + cb63_diam1.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicKrKr");

			cb63_diam2.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicKrKr"].Rows.Count; i++)
			{
				cb63_diam2.Items.Add(dataSet2.Tables["troinicKrKr"].Rows[i]["diam2"].ToString());
			}
		}

		private void cb63_diam2_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick,vys1 FROM troinicKrKr WHERE diam1 = " + cb63_diam1.SelectedItem.ToString() + " AND diam2 = " + cb63_diam2.SelectedItem.ToString() + "  ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicKrKr");

			cb63_vys1.Clear();
			cb63_zazor.Clear();
			cb63_thick.Clear();
			for (int i = 0; i < dataSet2.Tables["troinicKrKr"].Rows.Count; i++)
			{
				cb63_zazor.Text = (dataSet2.Tables["troinicKrKr"].Rows[i]["thick"].ToString());
				cb63_thick.Text = (dataSet2.Tables["troinicKrKr"].Rows[i]["thick"].ToString());
				cb63_vys1.Text = (dataSet2.Tables["troinicKrKr"].Rows[i]["vys1"].ToString());
			}
		}

		private void cb64_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb64_dlin.Enabled = true;
			cb64_dlin.Text = "";
			cb64_shir.Text = "";
			cb64_vys1.Clear();
			cb64_zazor.Clear();
			cb64_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dlin FROM troinicKrPr WHERE diam = " + cb64_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicKrPr");

			cb64_dlin.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicKrPr"].Rows.Count; i++)
			{
				cb64_dlin.Items.Add(dataSet2.Tables["troinicKrPr"].Rows[i]["dlin"].ToString());
			}
		}

		private void cb64_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb64_shir.Enabled = true;
			cb64_shir.Text = "";
			cb64_vys1.Clear();
			cb64_zazor.Clear();
			cb64_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir FROM troinicKrPr WHERE diam = " + cb64_diam.SelectedItem.ToString() + " AND dlin = " + cb64_dlin.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicKrPr");

			cb64_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicKrPr"].Rows.Count; i++)
			{
				cb64_shir.Items.Add(dataSet2.Tables["troinicKrPr"].Rows[i]["shir"].ToString());
			}
		}

		private void cb64_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb64_vys1.Enabled = true;
			cb64_zazor.Enabled = true;

			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick,vys FROM troinicKrPr WHERE diam = " + cb64_diam.SelectedItem.ToString() + " AND dlin = " + cb64_dlin.SelectedItem.ToString() + " AND shir = " + cb64_shir.SelectedItem.ToString() + "  ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicKrPr");

			cb64_vys1.Clear();
			cb64_zazor.Clear();
			cb64_thick.Clear();
			for (int i = 0; i < dataSet2.Tables["troinicKrPr"].Rows.Count; i++)
			{
				cb64_zazor.Text = (dataSet2.Tables["troinicKrPr"].Rows[i]["thick"].ToString());
				cb64_thick.Text = (dataSet2.Tables["troinicKrPr"].Rows[i]["thick"].ToString());
				cb64_vys1.Text = (dataSet2.Tables["troinicKrPr"].Rows[i]["vys"].ToString());
			}
		}

		private void cb71_dlin_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb71_shir.Enabled = true;
			cb71_shir.Text = "";
			cb71_dOtv.Clear();
			cb71_thick.Clear();
			string str = cb71_dlin.SelectedItem.ToString();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT  shir1 FROM flanPr WHERE dlin1 =" + cb71_dlin.SelectedItem.ToString() + "", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "flanPr");

			cb71_shir.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["flanPr"].Rows.Count; i++)
			{
				cb71_shir.Items.Add(dataSet2.Tables["flanPr"].Rows[i]["shir1"].ToString());
			}
		} //фланец прямоугольный

		private void cb71_shir_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dOtv,thick FROM flanPr WHERE dlin1 = " + cb71_dlin.SelectedItem.ToString() + " AND shir1 = " + cb71_shir.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "flanPr");
			cb71_dOtv.Clear();
			cb71_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["flanPr"].Rows.Count; i++)
			{
				cb71_dOtv.Text = (dataSet2.Tables["flanPr"].Rows[i]["dOtv"].ToString());
				cb71_thick.Text = (dataSet2.Tables["flanPr"].Rows[i]["thick"].ToString());
			}
		} 

		private void cb72_diam_SelectedIndexChanged(object sender, EventArgs e)
		{
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT dOtv,thick FROM flanKr WHERE diam = " + cb72_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "flanKr");
			cb72_dOtv.Clear();
			cb72_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["flanKr"].Rows.Count; i++)
			{
				cb72_dOtv.Text = (dataSet2.Tables["flanKr"].Rows[i]["dOtv"].ToString());
				cb72_thick.Text = (dataSet2.Tables["flanKr"].Rows[i]["thick"].ToString());
			}
		} //фланец круглый


		private void cb52_diam_SelectedIndexChanged(object sender, EventArgs e)
		{ 
			
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick FROM zaglKr WHERE diam = " + cb52_diam.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "zaglKr");

			cb52_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["zaglKr"].Rows.Count; i++)
			{
				cb52_thick.Text = (dataSet2.Tables["zaglKr"].Rows[i]["thick"].ToString());
			}
		}

		private void cb61_shir2_SelectedIndexChanged(object sender, EventArgs e)
		{

			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT thick,vys1 FROM troinicPrPr WHERE dlin1 = " + cb61_dlin1.SelectedItem.ToString() + " AND shir1 = " + cb61_shir1.SelectedItem.ToString() + " AND dlin2 = " + cb61_dlin2.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrPr");

			cb61_vys1.Clear();
			cb61_zazor.Clear();
			cb61_thick.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrPr"].Rows.Count; i++)
			{
				cb61_zazor.Text = (dataSet2.Tables["troinicPrPr"].Rows[i]["thick"].ToString());
				cb61_thick.Text = (dataSet2.Tables["troinicPrPr"].Rows[i]["thick"].ToString());
				cb61_vys1.Text = (dataSet2.Tables["troinicPrPr"].Rows[i]["vys1"].ToString());
			}
		}

		private void cb61_dlin2_SelectedIndexChanged(object sender, EventArgs e)
		{
			cb61_shir2.Enabled = true;
			
			cb61_shir2.Text = "";
			cb61_vys1.Clear();
			cb61_zazor.Clear();
			cb61_thick.Clear();
			SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
			SqlCommand command = new SqlCommand("SELECT DISTINCT shir2 FROM troinicPrPr WHERE dlin1 = " + cb61_dlin1.SelectedItem.ToString() + " AND shir1 = " + cb61_shir1.SelectedItem.ToString() + " AND dlin2 = " + cb61_dlin2.SelectedItem.ToString() + " ", sqlConnection);

			sqlDataAdapter2 = new SqlDataAdapter(command);

			dataSet2 = new DataSet();
			sqlDataAdapter2.Fill(dataSet2, "troinicPrPr");

			cb61_shir2.Items.Clear();

			for (int i = 0; i < dataSet2.Tables["troinicPrPr"].Rows.Count; i++)
			{
				cb61_shir2.Items.Add(dataSet2.Tables["troinicPrPr"].Rows[i]["shir2"].ToString());
			}
		}


		void setMaterial()
		{

			//задание материала
			
			string configName = "По умолчанию";
			string databaseName = "D:/SolidWorks/SOLIDWORKS/lang/russian/sldmaterials/solidworks materials.sldmat";
			string newPropName = "Оцинкованная сталь";
			((PartDoc)Part).SetMaterialPropertyName2(configName, databaseName, newPropName);

		}

		//детали

		private void bt_PrymVoz_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb1_dlin.Text) || string.IsNullOrEmpty(cb1_shir.Text) || string.IsNullOrEmpty(cb1_vys.Text))
			{
				MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{

				assembly = (AssemblyDoc)Part;
				double shir_pr, vys_pr, dlin_pr, zazor_pr;
				dlin_pr = Convert.ToDouble(cb1_dlin.Text);
				shir_pr = Convert.ToDouble(cb1_shir.Text);
				vys_pr = Convert.ToDouble(cb1_vys.Text);
				zazor_pr = Convert.ToDouble(cb1_zazor.Text);

				string name = $"Воздуховод прямоугольного сечения_{shir_pr}x{dlin_pr}";

				shir_pr = shir_pr / 1000;
				vys_pr = vys_pr / 1000;
				dlin_pr = dlin_pr / 1000;
				zazor_pr = zazor_pr / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
				airDuct duct = new airDuct();
				duct.createPrVozd(dlin_pr, shir_pr, vys_pr, zazor_pr, Part);
				setMaterial();
				savePart(name);
				string assPath = Part.GetPathName();
				//MessageBox.Show(assPath);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина: {dlin_pr * 1000}  Ширина: {shir_pr * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина: {dlin_pr * 1000}  Ширина: {shir_pr * 1000}";
				}


				addComponent2(assPath);
				//bt_razv1.Visible = true;
			}
		}

		private void bt_ZaglPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb51_dlin.Text) || string.IsNullOrEmpty(cb51_shir.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}else if (Convert.ToInt32(cb51_vys.Text) < 20 || Convert.ToInt32(cb51_vys.Text) > 100)
			{
				MessageBox.Show("Высота заглушки должна быть в диапазоне от 20 до 100 мм");
			}
			else
			{
				double shir_pr, vys_pr, dlin_pr;
				dlin_pr = Convert.ToDouble(cb51_dlin.Text);
				shir_pr = Convert.ToDouble(cb51_shir.Text);
				vys_pr = Convert.ToDouble(cb51_vys.Text);

				string name = $"Заглушка прямоугольного сечения_{dlin_pr}x{shir_pr}";

				shir_pr = shir_pr / 1000;
				vys_pr = vys_pr / 1000;
				dlin_pr = dlin_pr / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
				//Close();

				Zaglushka zaglPr = new Zaglushka();
				zaglPr.createZaglPr(dlin_pr, shir_pr, vys_pr, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина: {dlin_pr * 1000}  Ширина: {shir_pr * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина: {dlin_pr * 1000}  Ширина: {shir_pr * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_KrVozd_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb2_diam.Text) || string.IsNullOrEmpty(cb2_vys.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double diam_kr, vys_kr, zazor_kr;
				diam_kr = Convert.ToDouble(cb2_diam.Text);
				vys_kr = Convert.ToDouble(cb2_vys.Text);
				zazor_kr = Convert.ToDouble(cb2_zazor.Text);

				string name = $"Воздуховод круглого сечения_{diam_kr}";

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

				setMaterial();
				savePart(name);
				
				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр: {diam_kr * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр: {diam_kr * 1000}";
				}
				string assPath = Part.GetPathName();
				
				addComponent2(assPath);
			}
			//Close();
		}

		private void bt_PerehPrPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb31_dlin.Text) || string.IsNullOrEmpty(cb31_dlin1.Text) || string.IsNullOrEmpty(cb31_shir.Text) || string.IsNullOrEmpty(cb31_shir1.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double dlin, dlin1, shir, shir1, vys, zazor;
				dlin = Convert.ToDouble(cb31_dlin.Text);
				dlin1 = Convert.ToDouble(cb31_dlin1.Text);
				shir = Convert.ToDouble(cb31_shir.Text);
				shir1 = Convert.ToDouble(cb31_shir1.Text);
				vys = Convert.ToDouble(cb31_vys.Text);
				zazor = Convert.ToDouble(cb31_zazor.Text);

				string name = $"Переход прямоугольного сечения_{dlin}x{shir}_{dlin1}x{shir1}";

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
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				//Close();

				Perehod perehod1 = new Perehod();
				perehod1.createPrPrPerehod(dlin, dlin1, shir, shir1, vys, zazor, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина 1: {dlin * 1000}  Ширина 1: {shir * 1000}" + '\r' + '\n' + $"Длина 2: {dlin1 * 1000}  Ширина 2: {shir1 * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина 1: {dlin * 1000}  Ширина 1: {shir * 1000}" + '\r' + '\n' + $"Длина 2: {dlin1 * 1000}  Ширина 2: {shir1 * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_PerehPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb32_diam.Text) || string.IsNullOrEmpty(cb32_dlin.Text) || string.IsNullOrEmpty(cb32_shir.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double dlin, shir, diam, vys, zazor;
				dlin = Convert.ToDouble(cb32_dlin.Text);
				shir = Convert.ToDouble(cb32_shir.Text);
				diam = Convert.ToDouble(cb32_diam.Text);
				vys = Convert.ToDouble(cb32_vys.Text);
				zazor = Convert.ToDouble(cb32_zazor.Text);

				string name = $"Переход прямоугольного сечения_{diam}_{dlin}x{shir}";

				dlin = dlin / 1000;
				shir = shir / 1000;
				diam = diam / 1000;
				vys = vys / 1000;
				zazor = zazor / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Perehod perehod2 = new Perehod();
				perehod2.createPrPerehod(dlin, shir, diam, vys, zazor, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin * 1000}  Ширина : {shir * 1000}" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина 1: {dlin * 1000}  Ширина 1: {shir * 1000}" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_PerehKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb33_diam1.Text) || string.IsNullOrEmpty(cb33_diam2.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double diam1, diam2, vys, zazor;
				diam1 = Convert.ToDouble(cb33_diam1.Text);
				diam2 = Convert.ToDouble(cb33_diam2.Text);
				vys = Convert.ToDouble(cb33_vys.Text);
				zazor = Convert.ToDouble(cb33_zazor.Text);

				string name = $"Переход круглого сечения_{diam1}_{diam2}";

				diam1 = diam1 / 1000;
				diam2 = diam2 / 1000;
				vys = vys / 1000;
				zazor = zazor / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Perehod perehod3 = new Perehod();
				perehod3.createKrPerehod(diam1, diam2, vys, zazor, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр 1 : {diam1 * 1000}" + '\r' + '\n' + $"Диаметр 2 : {diam2 * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр 1 : {diam1 * 1000}" + '\r' + '\n' + $"Диаметр 2 : {diam2 * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_OtvodKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb42_diam.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double diam, zazor;
				diam = Convert.ToDouble(cb42_diam.Text);
				zazor = Convert.ToDouble(cb42_zazor.Text);

				string name = $"Отвод круглого сечения_{diam}_90";
				string name2 = $"Отвод круглого сечения_{diam}_90_сборка";


				diam = diam / 1000;
				zazor = zazor / 1000;

				double xAss = diam + diam / 2;
				double yAss = diam;

				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Otvod otvod1 = new Otvod();
				otvod1.createKrOtvod(diam, zazor, Part);
				setMaterial();
				savePart(name);
				 
				//получаем имя детали
				tmpPath = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(tmpPath);

				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name2 + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name2 + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name2 + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name2 + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name2 + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name2 + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name2 + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				//имя сборки
				string assPath = Part.GetPathName();
				string nameSbor = Path.GetFileNameWithoutExtension(assPath);


				//добавляем компонент в сборку
				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(tmpPath, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(tmpPath);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);



				createCircPattern(assPath, nameSbor, osnName2);

				//boolstatus = Part.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplaySketches, false);
				//boolstatus = Part.Extension.InsertScene("\\scenes\\01 basic scenes\\00 3 point faded.p2s");

				
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}


				addComponent2Assembly(assPath);
			}
		}

		private void bt_OtvodPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb41_dlin.Text) || string.IsNullOrEmpty(cb41_shir.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double dlin, shir, dlinOtv, zazor;
				dlin = Convert.ToDouble(cb41_dlin.Text);
				shir = Convert.ToDouble(cb41_shir.Text);
				dlinOtv = Convert.ToDouble(cb41_dlinOtvod.Text);
				zazor = Convert.ToDouble(cb41_zazor.Text);

				string name = $"Отвод прямоугольного сечения_{dlin}x{shir}_90";
				string name2 = $"Отвод прямоугольного сечения_{dlin}x{shir}_90_сборка";


				dlin = dlin / 1000;
				shir = shir / 1000;
				dlinOtv = dlinOtv / 1000;
				zazor = zazor / 1000;
				double xAss = (dlinOtv + shir / 2 + zazor) / 2;
				double yAss = dlinOtv / 2;

				iSwApp = new SldWorks();
				iSwApp.Visible = true;
				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Otvod otvod2 = new Otvod();
				otvod2.createPrOtvod(dlin, shir, dlinOtv, zazor, Part);
				setMaterial();
				savePart(name);
				//получаем имя детали
				tmpPath = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(tmpPath);

				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name2 + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name2 + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name2 + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name2 + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name2 + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name2 + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name2 + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				//имя сборки
				string assPath = Part.GetPathName();
				string nameSbor = Path.GetFileNameWithoutExtension(assPath);


				//добавляем компонент в сборку
				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(tmpPath, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(tmpPath);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				createCircPattern(assPath, nameSbor, osnName2);

				
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
				
				
				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin * 1000} Ширина : {shir * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin * 1000} Ширина : {shir * 1000}";
				}
				

				addComponent2Assembly(assPath);
			}
		}
		public void createCircPattern(string assemblyPath, string assemblyName, string partName)
		{
			int mateSelMark;
			int errorCode1 = 0;
			int mateError;
			//Активация документа со сборкой
			assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assemblyPath, true,
				(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
			Part = (ModelDoc2)assembly;
			//Иницииализация объекта для выбора элементов компонентов 
			swModelDocExt = Part.Extension;
			mateSelMark = 1;
			swFeatureManager = (FeatureManager)Part.FeatureManager;

			//Создание кругового массива элементов
			Part.ClearSelection2(true);
			swModelDocExt.SelectByID2(partName + "-1@" + assemblyName, "COMPONENT", 0, 0, 0,
				true, 1, null, 0);
			swModelDocExt.SelectByID2("Line1@Эскиз2@" + partName + "-1@" + assemblyName, "EXTSKETCHSEGMENT", 0, 0, 0,
				true, 2, null, 0);

			swFeature = (Feature)swFeatureManager.FeatureCircularPattern5(6, 0.26179938779915,
				true, "NULL",false,false,false,false,false,false,0,0, "NULL",false);
			Part.ForceRebuild3(false);
			//Part.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

		}

		private void bt_TroinicPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb61_dlin1.Text) || string.IsNullOrEmpty(cb61_dlin2.Text) 
				|| string.IsNullOrEmpty(cb61_shir1.Text) || string.IsNullOrEmpty(cb61_shir2.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else if (Convert.ToInt32(cb61_vys2.Text) > 100)
			{
				MessageBox.Show("Высота врезки должна быть не больше 100 мм");
			}
			else
			{
				double dlin1, shir1, dlin2, shir2, vys1, vys2, zazor;
				dlin1 = Convert.ToDouble(cb61_dlin1.Text);
				shir1 = Convert.ToDouble(cb61_shir1.Text);
				dlin2 = Convert.ToDouble(cb61_dlin2.Text);
				shir2 = Convert.ToDouble(cb61_shir2.Text);
				vys1 = Convert.ToDouble(cb61_vys1.Text);
				vys2 = Convert.ToDouble(cb61_vys2.Text);
				zazor = Convert.ToDouble(cb61_zazor.Text);

				string name = $"Тройник прямоугольного сечения с прямоугольной врезкой_{dlin1}x{shir1}_{dlin2}x{shir2}";
				string osn = $"Основание1_{dlin1}x{shir1}";
				string vrez = $"Врезка1_{dlin2}x{shir2}";

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

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik = new Troinik();
				troinik.createTroinikPr(dlin1, shir1, vys1, dlin2, shir2, vys2, zazor, Part);
				setMaterial();
				savePart(osn);
				string osnName = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(osnName);

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik1 = new Troinik();
				troinik1.createVrezPr(dlin2, shir2, vys2, zazor, Part);
				setMaterial();
				savePart(vrez);
				string vrezName = Part.GetPathName();
				string vrezName2 = Path.GetFileNameWithoutExtension(vrezName);


				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();
				//Сохраняем сборку
				if (File.Exists(assPut2 + "\\" + name + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm"), "*" 
						+ Path.GetExtension(assPut2 + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm") + "\\" 
							+ Path.GetFileNameWithoutExtension(assPut2 + "\\" + name + ".sldasm") + "_" + i 
							+ Path.GetExtension(assPut2 + "\\" + name + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
						(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				string assPath = Part.GetPathName();
				string name2 = Path.GetFileNameWithoutExtension(assPath);
				MessageBox.Show(assPath);
				MessageBox.Show(name2);

				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, 
					(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(osnName, 0, "", false, "", 0, 0, 0);
				component = assembly.AddComponent5(vrezName, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(osnName);
				iSwApp.CloseDoc(vrezName);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				int mateSelMark;
				int errorCode1 = 0;
				int mateError;
				//Активация документа со сборкой
				assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assPath, true,
					(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
				Part = (ModelDoc2)assembly;
				//Иницииализация объекта для выбора элементов компонентов 
				swModelDocExt = Part.Extension;
				mateSelMark = 1;

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Справа@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Плоскость5@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Сверху@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Сверху@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0,
					false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Спереди@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Спереди@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateDISTANCE,
					(int)swMateAlign_e.swMateAlignALIGNED, false, -vys1/2, -vys1 / 2, 
					-vys1 / 2, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина 1 : {dlin1 * 1000} Ширина 1 : {shir1 * 1000}" + '\r' + '\n' + $"Длина 2 : {dlin2 * 1000} Ширина 2 : {shir2 * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина 1 : {dlin1 * 1000} Ширина 1 : {shir1 * 1000}" + '\r' + '\n' + $"Длина 2 : {dlin2 * 1000} Ширина 2 : {shir2 * 1000}";
				}


				addComponent2Assembly(assPath);
			}
		}

		private void bt_TroinicPrKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb62_diam.Text) || string.IsNullOrEmpty(cb62_dlin.Text) || string.IsNullOrEmpty(cb62_shir.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else if (Convert.ToInt32(cb62_vys2.Text) > 100)
			{
				MessageBox.Show("Высота врезки должна быть не больше 100 мм");
			}
			else
			{
				double dlin, shir, vys1, diam, vys2, zazor;
				dlin = Convert.ToDouble(cb62_dlin.Text);
				shir = Convert.ToDouble(cb62_shir.Text);
				vys1 = Convert.ToDouble(cb62_vys1.Text);
				diam = Convert.ToDouble(cb62_diam.Text);
				vys2 = Convert.ToDouble(cb62_vys2.Text);
				zazor = Convert.ToDouble(cb62_zazor.Text);

				string name = $"Тройник прямоугольного сечения с круглой врезкой_{dlin}x{shir}_{diam}";
				string osn = $"Основание2_{dlin}x{shir}";
				string vrez = $"Врезка2_{diam}";

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

				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik = new Troinik();
				troinik.createTroinikPrKr(dlin, shir, vys1, diam, vys2, zazor, Part);
				setMaterial();
				savePart(osn);
				string osnName = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(osnName);

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik1 = new Troinik();
				troinik1.createVrezKr(diam, vys2, zazor, Part);
				setMaterial();
				savePart(vrez);
				string vrezName = Part.GetPathName();
				string vrezName2 = Path.GetFileNameWithoutExtension(vrezName);


				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				string assPath = Part.GetPathName();
				string name2 = Path.GetFileNameWithoutExtension(assPath);
				MessageBox.Show(assPath);
				MessageBox.Show(name2);

				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(osnName, 0, "", false, "", 0, 0, 0);
				component = assembly.AddComponent5(vrezName, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(osnName);
				iSwApp.CloseDoc(vrezName);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				int mateSelMark;
				int errorCode1 = 0;
				int mateError;
				//Активация документа со сборкой
				assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assPath, true,
					(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
				Part = (ModelDoc2)assembly;
				//Иницииализация объекта для выбора элементов компонентов 
				swModelDocExt = Part.Extension;
				mateSelMark = 1;

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Справа@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Плоскость5@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Сверху@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Сверху@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Спереди@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Спереди@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateDISTANCE,
					(int)swMateAlign_e.swMateAlignALIGNED, false, -vys1 / 2, -vys1 / 2, -vys1 / 2, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin * 1000} Ширина : {shir * 1000}" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin * 1000} Ширина : {shir * 1000}" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}

				addComponent2Assembly(assPath);
			}
		}

		private void bt_TroinicKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb63_diam1.Text) || string.IsNullOrEmpty(cb63_diam2.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else if (Convert.ToInt32(cb63_vys2.Text) > 100)
			{
				MessageBox.Show("Высота врезки должна быть не больше 100 мм");
			}
			else
			{
				double diam1, diam2, vys1, vys2, zazor;
				diam1 = Convert.ToDouble(cb63_diam1.Text);
				vys1 = Convert.ToDouble(cb63_vys1.Text);
				diam2 = Convert.ToDouble(cb63_diam2.Text);
				vys2 = Convert.ToDouble(cb63_vys2.Text);
				zazor = Convert.ToDouble(cb63_zazor.Text);

				string name = $"Тройник круглого сечения с круглой врезкой_{diam1}_{diam2}";
			
				string osn = $"Основание3_{diam1}";
				string vrez = $"Врезка3_{diam2}";

				diam1 = diam1 / 1000;
				diam2 = diam2 / 1000;
				vys1 = vys1 / 1000;
				vys2 = vys2 / 1000;
				zazor = zazor / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik = new Troinik();
				troinik.createTroinikKr(diam1, vys1, diam2, vys2, zazor, Part);
				setMaterial();
				savePart(osn);
				string osnName = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(osnName);

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik1 = new Troinik();
				troinik1.createVrezKr2(diam2, vys2, zazor, Part);
				setMaterial();
				savePart(vrez);
				string vrezName = Part.GetPathName();
				string vrezName2 = Path.GetFileNameWithoutExtension(vrezName);

				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				string assPath = Part.GetPathName();
				string name2 = Path.GetFileNameWithoutExtension(assPath);
				MessageBox.Show(assPath);
				MessageBox.Show(name2);

				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(osnName, 0, "", false, "", 0, 0, 0);
				component = assembly.AddComponent5(vrezName, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(osnName);
				iSwApp.CloseDoc(vrezName);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				int mateSelMark;
				int errorCode1 = 0;
				int mateError;
				//Активация документа со сборкой
				assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assPath, true,
					(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
				Part = (ModelDoc2)assembly;
				//Иницииализация объекта для выбора элементов компонентов 
				swModelDocExt = Part.Extension;
				mateSelMark = 1;

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Справа@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Плоскость5@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Сверху@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Справа@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Спереди@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Спереди@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateDISTANCE,
					(int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, vys1 / 2, vys1 / 2, vys1 / 2, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				Part.SketchManager.InsertSketch(true);
				boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
				Part.ClearSelection2(true);
				skSegment = Part.SketchManager.CreateCircle(0, 0, 0, diam1 / 2 - 0.001, 0, 0);
				Part.ClearSelection2(true);
				Part.SketchManager.InsertSketch(true);

				boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);
				swFeature = Part.FeatureManager.FeatureCut4(false, false, false, 9, 1, 0.01, 0.01, false, false, false, false,
					1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false,
					0, 0, false, false);
				Part.ClearSelection2(true);
				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр 1 : {diam1 * 1000}" + '\r' + '\n' + $"Диаметр 2 : {diam2 * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр 1 : {diam1 * 1000}" + '\r' + '\n' + $"Диаметр 2 : {diam2 * 1000}";
				}


				addComponent2Assembly(assPath);
			}
		}

		private void bt_TroinicKrPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb64_diam.Text) || string.IsNullOrEmpty(cb64_dlin.Text) || string.IsNullOrEmpty(cb64_shir.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else if (Convert.ToInt32(cb64_vys2.Text) > 100)
			{
				MessageBox.Show("Высота врезки должна быть не больше 100 мм");
			}
			else
			{
				double dlin, shir, vys1, diam, vys2, zazor;
				dlin = Convert.ToDouble(cb64_dlin.Text);
				shir = Convert.ToDouble(cb64_shir.Text);
				vys1 = Convert.ToDouble(cb64_vys1.Text);
				diam = Convert.ToDouble(cb64_diam.Text);
				vys2 = Convert.ToDouble(cb64_vys2.Text);
				zazor = Convert.ToDouble(cb64_zazor.Text);

				string name = $"Тройник круглого сечения с прямоугольной врезкой_{diam}_{dlin}x{shir}";

				string osn = $"Основание4_{diam}";
				string vrez = $"Врезка4_{dlin}x{shir}";

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

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik = new Troinik();
				troinik.createTroinikKrPr(dlin, shir, vys1, diam, vys2, zazor, Part);
				setMaterial();
				savePart(osn);
				string osnName = Part.GetPathName();
				string osnName2 = Path.GetFileNameWithoutExtension(osnName);

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;

				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Troinik troinik1 = new Troinik();
				troinik1.createVrezPr2(dlin,shir, vys2, zazor, Part);
				setMaterial();
				savePart(vrez);
				string vrezName = Part.GetPathName();
				string vrezName2 = Path.GetFileNameWithoutExtension(vrezName);

				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				string assPath = Part.GetPathName();
				string name2 = Path.GetFileNameWithoutExtension(assPath);
				MessageBox.Show(assPath);
				MessageBox.Show(name2);

				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(osnName, 0, "", false, "", 0, 0, 0);
				component = assembly.AddComponent5(vrezName, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(osnName);
				iSwApp.CloseDoc(vrezName);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				int mateSelMark;
				int errorCode1 = 0;
				int mateError;
				//Активация документа со сборкой
				assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assPath, true,
					(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
				Part = (ModelDoc2)assembly;
				//Иницииализация объекта для выбора элементов компонентов 
				swModelDocExt = Part.Extension;
				mateSelMark = 1;

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Справа@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Плоскость5@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Сверху@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Справа@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
					(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);

				//Создание зависимости плоскости 
				Part.ClearSelection2(true);
				swModelDocExt.SelectByID2("Спереди@" + vrezName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swModelDocExt.SelectByID2("Спереди@" + osnName2 + "-1@" + name2, "PLANE", 0, 0, 0,
					true, mateSelMark, null, 0);
				swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateDISTANCE,
					(int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, vys1 / 2, vys1 / 2, vys1 / 2, 0, 0, 0, 0, 0, false, false, 0, out mateError);

				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				Part.SketchManager.InsertSketch(true);
				boolstatus = Part.Extension.SelectByID2("Спереди", "PLANE", 0, 0, 0, false, 0, null, 0);
				Part.ClearSelection2(true);
				skSegment = Part.SketchManager.CreateCircle(0, 0, 0, diam / 2 - 0.001, 0, 0);
				Part.ClearSelection2(true);
				Part.SketchManager.InsertSketch(true);

				boolstatus = Part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, false, 0, null, 0);
				swFeature = Part.FeatureManager.FeatureCut4(false, false, false, 9, 1, 0.01, 0.01, false, false, false, false,
					1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false,
					0, 0, false, false);
				Part.ClearSelection2(true);
				Part.ForceRebuild3(false);
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}" + '\r' + '\n' + $"Длина : {dlin * 1000}  Ширина : {shir * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}" + '\r' + '\n' + $"Длина : {dlin * 1000}  Ширина : {shir * 1000}";
				}

				addComponent2Assembly(assPath);
			}
		}

		private void bt_FlanPr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb71_dlin.Text) || string.IsNullOrEmpty(cb71_shir.Text))
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double dlin1, shir1, dlin2, shir2, sDlin, sShir, thick, dOtv;
				int countDlin, countShir;
				dlin1 = Convert.ToDouble(cb71_dlin.SelectedItem.ToString());
				shir1 = Convert.ToDouble(cb71_shir.SelectedItem.ToString());
				dOtv = Convert.ToDouble(cb71_dOtv.Text);
				thick = Convert.ToDouble(cb71_thick.Text);

				string name = $"Фланец прямоугольного сечения_{dlin1}x{shir1}";


				dlin1 = dlin1 / 1000;
				shir1 = shir1 / 1000;
				dOtv = dOtv / 1000;
				thick = thick / 1000;
				dlin2 = 0;
				shir2 = 0;
				sDlin = 0;
				sShir = 0;
				countDlin = 0;
				countShir = 0;

				SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
				SqlCommand command = new SqlCommand("SELECT DISTINCT dlin2,shir2,sDlin,sShir,countDlin,countShir FROM flanPr WHERE dlin1 = " + cb71_dlin.SelectedItem.ToString() + " AND shir1 = " + cb71_shir.SelectedItem.ToString() + " ", sqlConnection);

				sqlDataAdapter2 = new SqlDataAdapter(command);

				dataSet2 = new DataSet();
				sqlDataAdapter2.Fill(dataSet2, "flanPr");

				for (int i = 0; i < dataSet2.Tables["flanPr"].Rows.Count; i++)
				{
					dlin2 = Convert.ToDouble(dataSet2.Tables["flanPr"].Rows[i]["dlin2"].ToString());
					shir2 = Convert.ToDouble(dataSet2.Tables["flanPr"].Rows[i]["shir2"].ToString());
					sDlin = Convert.ToDouble(dataSet2.Tables["flanPr"].Rows[i]["sDlin"].ToString());
					sShir = Convert.ToDouble(dataSet2.Tables["flanPr"].Rows[i]["sShir"].ToString());
					countDlin = Convert.ToInt32(dataSet2.Tables["flanPr"].Rows[i]["countDlin"].ToString());
					countShir = Convert.ToInt32(dataSet2.Tables["flanPr"].Rows[i]["countShir"].ToString());

				}

				dlin2 = dlin2 / 1000;
				shir2 = shir2 / 1000;
				sDlin = sDlin / 1000;
				sShir = sShir / 1000;

				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;
				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Flanc flancPr = new Flanc();
				flancPr.createPrFlanc(dlin1, shir1, dlin2, shir2, sDlin, sShir, countDlin, countShir, thick, dOtv, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin1 * 1000}  Ширина : {shir1 * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Длина : {dlin1 * 1000}  Ширина : {shir1 * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_FlanKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb72_diam.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else
			{
				double diam, dOs, thick, dOtv;
				int count;
				diam = Convert.ToDouble(cb72_diam.SelectedItem.ToString());
				dOtv = Convert.ToDouble(cb72_dOtv.Text);
				thick = Convert.ToDouble(cb72_thick.Text);

				string name = $"Фланец круглого сечения_{diam}";

				diam = diam / 1000;
				dOtv = dOtv / 1000;
				thick = thick / 1000;

				dOs = 0;
				count = 0;

				SqlConnection sqlConnection = new SqlConnection(@"Data Source=YULIAS\SQLEXPRESS;Initial Catalog=airDuct;Integrated Security=True");
				SqlCommand command = new SqlCommand("SELECT DISTINCT dOs,count FROM flanKr WHERE diam = " + cb72_diam.SelectedItem.ToString() + " ", sqlConnection);

				sqlDataAdapter2 = new SqlDataAdapter(command);

				dataSet2 = new DataSet();
				sqlDataAdapter2.Fill(dataSet2, "flanKr");

				for (int i = 0; i < dataSet2.Tables["flanKr"].Rows.Count; i++)
				{
					dOs = Convert.ToDouble(dataSet2.Tables["flanKr"].Rows[i]["dOs"].ToString());
					count = Convert.ToInt32(dataSet2.Tables["flanKr"].Rows[i]["count"].ToString());

				}

				dOs = dOs / 1000;


				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;
				//Close();
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Flanc flancKr = new Flanc();
				flancKr.createKrFlanc(diam, dOs, dOtv, count, thick, Part);
				setMaterial();
				savePart(name);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				string assPath = Part.GetPathName();

				addComponent2(assPath);
			}
		}

		private void bt_ZaglKr_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(cb52_diam.Text) )
			{
				System.Windows.Forms.MessageBox.Show("Заполните параметры для построения детали!");
			}
			else if (Convert.ToInt32(cb52_vys.Text) < 20 || Convert.ToInt32(cb52_vys.Text) > 100)
			{
				MessageBox.Show("Высота заглушки должна быть в диапазоне от 20 до 100 мм");
			}
			else
			{
				double diam, vys, thick;
				diam = Convert.ToDouble(cb52_diam.Text);
				vys = Convert.ToDouble(cb52_vys.Text);
				thick = Convert.ToDouble(cb52_thick.Text);

				string name = $"Заглушка круглого сечения_{diam}";
				string bottom = $"Заглушка круглого сечения_{diam}(bottom)";
				string top = $"Заглушка круглого сечения_{diam}(top)";

				diam = diam / 1000;
				vys = vys / 1000;
				thick = thick / 1000;

				iSwApp = new SldWorks();
				iSwApp.Visible = true;

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;
				iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

				Zaglushka zagl = new Zaglushka();
				zagl.createZaglKrBottom(diam, vys, thick, Part);
				setMaterial();
				savePart(bottom);
				string botName = Part.GetPathName();
				string botName2 = Path.GetFileNameWithoutExtension(botName);

				iSwApp.NewPart();
				Part = iSwApp.IActiveDoc2;
				Zaglushka zagl2 = new Zaglushka();
				zagl2.createZaglKrTop(diam, vys, thick, Part);
				setMaterial();
				savePart(top);
				string topName = Part.GetPathName();
				string topName2 = Path.GetFileNameWithoutExtension(topName);

				//создаем сборку 
				//Новое окно сборки
				Part = (ModelDoc2)iSwApp.NewAssembly();

				//Сохраняем сборку

				if (File.Exists(assPut2 + "\\" + name + ".sldasm"))
				{
					//Такой файл уже существует в конечной папке
					String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm"), "*" + Path.GetExtension(assPut2 + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
					for (int i = 0; i < dirsfile.Length; i++)
					{
						string newname = Path.GetDirectoryName(assPut2 + "\\" + name + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(assPut2 + "\\" + name + ".sldasm") + "_" + i + Path.GetExtension(assPut2 + "\\" + name + ".sldasm"); //Новое имя файла
						if (!File.Exists(newname))
						{
							Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
							break;
						}
					}
				}
				else
				{
					Part.SaveAs3(assPut2 + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				}
				string assPath = Part.GetPathName();
				string name2 = Path.GetFileNameWithoutExtension(assPath);

				object component;
				//Считывание открытого окна сборки
				assembly = (AssemblyDoc)iSwApp.ActiveDoc;

				//Активация документа сборки
				Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
				//Добавление компонентов в документ сборки
				component = assembly.AddComponent5(topName, 0, "", false, "", 0, 0, 0);
				component = assembly.AddComponent5(botName, 0, "", false, "", 0, 0, 0);
				//Закрытие открытых файлов компонентов в среде SolidWorks
				iSwApp.CloseDoc(topName);
				iSwApp.CloseDoc(botName);
				//Перестроение документа
				assembly.ForceRebuild();
				//Фокус камеры на компоненты
				//Part.ViewZoomtofit2();
				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);


				AddMates(assPath, name2, topName2, botName2, vys);

				Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
					(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

				if (txt_hint1.TextLength == 0 && txt_hint2.TextLength == 0)
				{
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}
				else
				{
					txt_hint2.Text = txt_hint1.Text;
					txt_hint1.Text = "Недавние размеры:" + '\r' + '\n' + $"Диаметр : {diam * 1000}";
				}


				addComponent2Assembly(assPath);
			}
		}

		
		void savePart(string name)
		{
			string trg = "C:\\Users\\YuliaS\\Desktop\\Воздуховоды";
			FolderBrowserDialog target = new FolderBrowserDialog();
			target.Description = "Выберите папку для сохранения сборки";
			target.RootFolder = System.Environment.SpecialFolder.MyComputer;
			target.SelectedPath = "C:\\Users\\YuliaS\\Desktop\\Воздуховоды";
			if (target.ShowDialog() == DialogResult.OK)
			{
				trg = target.SelectedPath;
			}

		
			string name2 = trg + "\\" + name + ".sldprt";
			if (File.Exists(name2))
			{
				//Такой файл уже существует в конечной папке
				String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(name2), "*" + Path.GetExtension(name2).Remove(0, 1)); //Поиск всех файлов в папке с расширением
				for (int i = 0; i < dirsfile.Length; i++)
				{
					string newname = Path.GetDirectoryName(name2) + "\\" + Path.GetFileNameWithoutExtension(name2)+"_" + i + Path.GetExtension(name2); //Новое имя файла
					if (!File.Exists(newname))
					{
						Part.SaveAs3(newname , (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
				(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
						break;
					}
				}
			}
			else
			{
				Part.SaveAs3(name2, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
				(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
			}


		}

		public void createAssembly(string path,string name)
		{
			iSwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
			ModelDoc2 Part;

			//Новое окно сборки
			Part = (ModelDoc2)iSwApp.NewAssembly();

			//Сохраняем сборку

			if (File.Exists(path + "\\" + name + ".sldasm"))
			{
				//Такой файл уже существует в конечной папке
				String[] dirsfile = Directory.GetFiles(Path.GetDirectoryName(path + "\\" + name + ".sldasm"), "*" + Path.GetExtension(path + "\\" + name + ".sldasm").Remove(0, 1)); //Поиск всех файлов в папке с расширением
				for (int i = 0; i < dirsfile.Length; i++)
				{
					string newname = Path.GetDirectoryName(path + "\\" + name + ".sldasm") + "\\" + Path.GetFileNameWithoutExtension(path + "\\" + name + ".sldasm")+"_" + i + Path.GetExtension(path + "\\" + name + ".sldasm"); //Новое имя файла
					if (!File.Exists(newname))
					{
						Part.SaveAs3(newname, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
				(int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen); ; //Сохранить файл с новым именем
						break;
					}
				}
			}
			else
			{
				Part.SaveAs3(path + "\\" + name + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

			}

			string assPath = Part.GetPathName();

			Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);

		}

		public void addComponent(IModelDoc2 Part, string assPath, string partPath, double x, double y, double z)
		{
			object component;
			//Считывание открытого окна сборки
			assembly = (AssemblyDoc)iSwApp.ActiveDoc;

			//Активация документа сборки
			Part = (ModelDoc2)iSwApp.ActivateDoc3(assPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
			//Добавление компонентов в документ сборки
			component = assembly.AddComponent5(partPath, 0, "", false, "", x, y, z);
			//Закрытие открытых файлов компонентов в среде SolidWorks
			iSwApp.CloseDoc(partPath);
			//Перестроение документа
			assembly.ForceRebuild();
			//Фокус камеры на компоненты
			//Part.ViewZoomtofit2();
			Part.SaveAs3(assPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
		}

		public void addComponent2(string assPath)
		{
			object component;
			//ModelDoc2 Part;
			ModelDoc2 tmpObj;
			//Закрытие открытых файлов компонентов в среде SolidWorks
			iSwApp.CloseDoc(assPath);

			//Считывание открытого окна сборки
			assembly = (AssemblyDoc)iSwApp.ActiveDoc;
			//string compName = Part.GetPathName();

			//Добавление координатных систем компонентов в массив
			//string xcoorsysnames = "Coordinate System1";
			//string mainPath = txt_target + "\\" + txt_name + ".sldasm";

			tmpObj = (ModelDoc2)iSwApp.OpenDoc6(assPath, (int)swDocumentTypes_e.swDocPART, 
				(int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

			//Активация документа сборки
			Part = (ModelDoc2)iSwApp.ActivateDoc3(assPut, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
			
			//Добавление компонентов в документ сборки
			component = assembly.AddComponent5(assPath, 0, "", false, "", 0, 0, 0); //путь к детали или сборке 
			//Закрытие открытых файлов компонентов в среде SolidWorks
			iSwApp.CloseDoc(assPath);
			//Перестроение документа
			assembly.ForceRebuild();
			//Фокус камеры на компоненты
			//Part.ViewZoomtofit2();
			Part.SaveAs3(assPut, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
			MessageBox.Show("Элемент добавлен в сборку");


		}

		public void addComponent2Assembly(string assPath)
		{
			object component;
			//ModelDoc2 Part;
			ModelDoc2 tmpObj;
			//Закрытие открытых файлов компонентов в среде SolidWorks
			iSwApp.CloseDoc(assPath);

			//Считывание открытого окна сборки
			assembly = (AssemblyDoc)iSwApp.ActiveDoc;
			//string compName = Part.GetPathName();

			//Добавление координатных систем компонентов в массив
			//string xcoorsysnames = "Coordinate System1";
			//string mainPath = txt_target + "\\" + txt_name + ".sldasm";

			tmpObj = (ModelDoc2)iSwApp.OpenDoc6(assPath, (int)swDocumentTypes_e.swDocASSEMBLY, 
				(int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

			//Активация документа сборки
			Part = (ModelDoc2)iSwApp.ActivateDoc3(assPut, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);

			//Добавление компонентов в документ сборки
			component = assembly.AddComponent5(assPath, 0, "", false, "", 0, 0, 0); //путь к детали или сборке 
																					//Закрытие открытых файлов компонентов в среде SolidWorks
			iSwApp.CloseDoc(assPath);
			//Перестроение документа
			assembly.ForceRebuild();
			//Фокус камеры на компоненты
			//Part.ViewZoomtofit2();
			//boolstatus = Part.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplaySketches, false);
			Part.SaveAs3(assPut, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
			MessageBox.Show("Элемент добавлен в сборку");

		}
		public void AddMates(string assemblyPath, string assemblyName, string topName, string bottomName, double vys)
		{
			int mateSelMark;
			int errorCode1 = 0;
			int mateError;
			//Активация документа со сборкой
			assembly = (AssemblyDoc)iSwApp.ActivateDoc3(assemblyPath, true, 
				(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
			Part = (ModelDoc2)assembly;
			//Иницииализация объекта для выбора элементов компонентов 
			swModelDocExt = Part.Extension;
			mateSelMark = 1;

			//Создание зависимости плоскости низа с плоскостью верха
			Part.ClearSelection2(true);
			swModelDocExt.SelectByID2("Спереди@" + bottomName + "-1@" + assemblyName, "PLANE", 0, 0, 0,
				true, mateSelMark, null, 0);
			swModelDocExt.SelectByID2("Плоскость4@" + topName + "-1@" + assemblyName, "PLANE", 0, 0, 0,
				true, mateSelMark, null, 0);
			swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT, 
				(int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out mateError);

			Part.ForceRebuild3(false);

			Part.ClearSelection2(true);
			swModelDocExt.SelectByID2("Point1@Исходная точка@" + bottomName + "-1@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0,
				true, mateSelMark, null, 0);
			swModelDocExt.SelectByID2("Point1@Исходная точка@" + topName + "-1@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0,
				true, mateSelMark, null, 0);
			swFeature = (Feature)assembly.AddMate5((int)swMateType_e.swMateDISTANCE, 
				(int)swMateAlign_e.swMateAlignALIGNED, false, vys, vys, vys, 0, 0, 0, 0, 0, false, false, 0, out mateError);
			Part.ForceRebuild3(false);
		}



		//кнопки далее
		private void bt_Next1_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage2"];
		}

		private void bt_Next2_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage3"];
		}

		private void bt_Next31_Click(object sender, EventArgs e)
		{
			tabControl2.SelectedTab = tabControl2.TabPages["TabPage10"];
		}

		private void bt_Next32_Click(object sender, EventArgs e)
		{
			tabControl2.SelectedTab = tabControl2.TabPages["TabPage11"];
		}

		private void bt_Next33_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage4"];
		}

		private void bt_Next41_Click(object sender, EventArgs e)
		{
			tabControl3.SelectedTab = tabControl3.TabPages["TabPage13"];
		}

		private void bt_Next42_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage5"];
		}

		private void bt_Next51_Click(object sender, EventArgs e)
		{
			tabControl4.SelectedTab = tabControl4.TabPages["TabPage15"];
		}

		private void bt_Next52_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage6"];
		}

		private void bt_Next61_Click(object sender, EventArgs e)
		{
			tabControl5.SelectedTab = tabControl5.TabPages["TabPage17"];
		}

		private void bt_Nezt62_Click(object sender, EventArgs e)
		{
			tabControl5.SelectedTab = tabControl5.TabPages["TabPage18"];
		}

		private void bt_Next63_Click(object sender, EventArgs e)
		{
			tabControl5.SelectedTab = tabControl5.TabPages["TabPage19"];
		}

		private void bt_Next64_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage8"];
		}

		private void bt_Next71_Click(object sender, EventArgs e)
		{
			tabControl6.SelectedTab = tabControl6.TabPages["TabPage20"];
		}

		private void bt_Next72_Click(object sender, EventArgs e)
		{
			tabControl1.SelectedTab = tabControl1.TabPages["TabPage1"];
		}

		
		//главная
		

		private void bt_1_Click(object sender, EventArgs e)
		{	
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage1"];
		}

		private void bt_2_Click(object sender, EventArgs e)
		{
		
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage2"];
			
		}

		private void bt_3_Click(object sender, EventArgs e)
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage3"];
			
		}

		private void bt_4_Click(object sender, EventArgs e) //переход пр 
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage3"];
				tabControl2.SelectedTab = tabControl2.TabPages["TabPage10"];
			
		}

		private void bt_5_Click(object sender, EventArgs e) //переход кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage3"];
				tabControl2.SelectedTab = tabControl2.TabPages["TabPage11"];
			
		}

		private void bt_6_Click(object sender, EventArgs e) //отвод пр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage4"];
				tabControl3.SelectedTab = tabControl3.TabPages["TabPage12"];
			
		}

		private void bt_7_Click(object sender, EventArgs e) //отвод кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage4"];
				tabControl3.SelectedTab = tabControl3.TabPages["TabPage13"];
			
		}

		private void bt_8_Click(object sender, EventArgs e) //загл пр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage5"];
				tabControl4.SelectedTab = tabControl4.TabPages["TabPage14"];
			
		}

		private void bt_9_Click(object sender, EventArgs e) //загл кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage5"];
				tabControl4.SelectedTab = tabControl4.TabPages["TabPage15"];
			
		}

		private void bt_10_Click(object sender, EventArgs e) //трой пр пр
		{
		
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage6"];
				tabControl5.SelectedTab = tabControl5.TabPages["TabPage16"];
			
		}

		private void bt_11_Click(object sender, EventArgs e) //трой пр кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage6"];
				tabControl5.SelectedTab = tabControl5.TabPages["TabPage17"];
			
		}

		private void bt_12_Click(object sender, EventArgs e) //трой кр кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage6"];
				tabControl5.SelectedTab = tabControl5.TabPages["TabPage18"];
			
		}

		private void bt_13_Click(object sender, EventArgs e)// трой кр пр
		{

				tabControl1.SelectedTab = tabControl1.TabPages["TabPage6"];
				tabControl5.SelectedTab = tabControl5.TabPages["TabPage19"];
			
		}

		private void bt_14_Click(object sender, EventArgs e) //флан пр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage8"];
				tabControl6.SelectedTab = tabControl6.TabPages["TabPage9"];
			
		}

		private void bt_15_Click(object sender, EventArgs e)//флан кр
		{
			
				tabControl1.SelectedTab = tabControl1.TabPages["TabPage8"];
				tabControl6.SelectedTab = tabControl6.TabPages["TabPage20"];
			
		}

		private void groupBox3_Enter(object sender, EventArgs e)
		{

		}

		private void textBox15_TextChanged(object sender, EventArgs e)
		{

		}

		private void label249_Click(object sender, EventArgs e)
		{

		}

		//private void bt_razv1_Click(object sender, EventArgs e)
		//{
			
		//	double shir_pr, vys_pr, dlin_pr, zazor_pr;
		//	dlin_pr = Convert.ToDouble(cb1_dlin.Text);
		//	shir_pr = Convert.ToDouble(cb1_shir.Text);
		//	string name = $"Воздуховод прямоугольного сечения_{shir_pr}x{dlin_pr}";

		//	iSwApp = new SldWorks();
		//	iSwApp.Visible = true;

		//	iSwApp.INewDocument2(name, (int)swDwgPaperSizes_e.swDwgPaperA3size,0,0);
		//	Part = iSwApp.IActiveDoc2;

		//}
	}
}
