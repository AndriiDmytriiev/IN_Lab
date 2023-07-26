

using ClosedXML.Excel;
using RDotNet;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Windows.Forms;

namespace BI_CPV_tool
{
  public class Form1 : Form
  {
    internal REngine engine;
    private PrintDocument docToPrint = new PrintDocument();
    public static string connectionString = "";
    public static bool filterIsOn = false;
    public static float intMaxPoints = 200000f;
    public static string strExcelFileName;
    public static DataGridView[] arrDataGridView = new DataGridView[1000];
    public static int intCounterDataGridViews = 0;
    public static string strFilterExclude = "";
    public static string strFileName = "";
    public static string strRscript = "";
    public static string strRpath = "";
    public static string strDataDir = "";
    public static string strOutputDir = "";
    public static string strCalcID = "";
    public static string strSQLfiltered = "";
    public static string strLaufNRDateMin = "";
    public static string strLaufNRDateMax = "";
    public static bool dateIsGood = false;
    public static bool filterOn = false;
    public static DataTable dt = (DataTable) null;
    public static string[] strArrToolTip = new string[500];
    public static FormWindowState LastWindowState;
    public static string[] strArg = new string[3]
    {
      "",
      "",
      ""
    };
    public static string strHasID = "";
    public static string strInitialDate = "";
    public static string strEndDate = "";
    public static int last = 1;
    public static string[] strArr = new string[1]{ "" };
    public static string[] strArr2 = new string[500];
    public static int ID = 0;
    public static int intIndex = 0;
    public static string NParameterTotal = "";
    public static string NStatistically = "";
    public static string PercentStatistically = "";
    public static string DoNotFitStatistically = "";
    public static string CalcID = "";
    public static string User = "";
    public static string TimePointData = "";
    public static string TimePointCalc = "";
    public static string Note = "";
    public static int Active = 0;
    public static bool IsCalcIDAvailable = false;
    public static bool IsCalcIDAvailable3 = false;
    public static bool IsCalcIDAvailable4 = false;
    public static string strOutPutPath = "";
    public static string strQuery3 = "";
    public static string strQuery4 = "";
    public static string OZID = "";
    public static string TotalN = "";
    public static string VIRT_OZID = "";
    public static string FitStatistically = "";
    public static string Additional_note = "";
    public static string Status_fit_statistically = "";
    public static string Num_VIRT_OZID_not_fit_stat = "";
    public static string Percent_of_values_status_KPI0_KPI3 = "";
    public static string Num_of_values_status_KPI0_KPI3 = "";
    public static string RelevantForDiscussion = "";
    public static string KPI0 = "";
    public static string KPI1 = "";
    public static string KPI2 = "";
    public static string KPI3 = "";
    public static string GraphID = "";
    public static string[] arr1 = new string[500];
    public static string[] arr2 = new string[500];
    public static string[] arr3 = new string[500];
    public static string strFilterLaufnr = "";
    public static string strOZID = "";
    public static bool blnPressed = false;
    public string[,] strMatrix = new string[500, 15];
    public static int[,] intParameters = new int[200, 3];
    public static string strQuery = "";
    public static string strQuery2 = "";
    public static int strRowsCount = 0;
    public static string strFilter = "";
    public static bool blnFlag = false;
    public static bool stopCalc = false;
    private Matrix transform = new Matrix();
    public static float s_dScrollValue = 1.01f;
    private double m_dZoomscale = 1.0;
    private Control ba;
    private IContainer components = (IContainer) null;
    private TabControl tabMain;
    private TabPage Main;
    private ComboBox cmbLaufnrMax;
    private ComboBox cmbLaufnrMin;
    private ComboBox cmbProductCode;
    private CheckBox chkWithoutCL;
    private Label label5;
    private CheckBox chLastN;
    private Label label4;
    private CheckBox chLastM;
    private Label label3;
    private CheckBox chkSortDate;
    private Label label2;
    private CheckBox chkAllVirtOzid;
    private CheckBox chkAllRefCpv;
    private CheckedListBox clbVirtOzid;
    private CheckedListBox clbRefCpv;
    private TextBox textBox1;
    private Button btnCalculationView;
    private PictureBox pictureBox1;
    private Button btnCalculationSearch;
    private ListBox listBox2;
    private DataGridView dataGridViewTemp;
    private TextBox txtTestSQL;
    private Button btnOpenFile;
    private RadioButton rbYears;
    private RadioButton rbMonths;
    private RadioButton rbWeeks;
    private RadioButton rbDays;
    private Button button1;
    private ListBox listBox1;
    private TextBox txtResult;
    private Label label1;
    private CheckedListBox lstCheckExclVirtOzid;
    private Label label58;
    private DataGridView dataGridViewRaw;
    private Button btnCalculation;
    private Label lblDataGridTitle;
    private CheckBox chkExclVirtOzid;
    private Label lblActiveEvo;
    private Label lblListCheckExclVirtOzid;
    private TextBox txtLastMdataPoints;
    private Label lblLastMdataPoints;
    private Label lblLastNdataPoints;
    private Button btnReset;
    private Button btnFilterData;
    private CheckBox chkLaufNr;
    private Label lblActiveLN;
    private DateTimePicker dtSortDateTo;
    private DateTimePicker dtSortDateFrom;
    private Label label6;
    private Label lblLaufNRfrom;
    private Label lblLaufNRto;
    private Label lblSortDate;
    private Label lblVirtOzid;
    private Label lblRferencedCPV;
    private Label lblProductCode;
    private TextBox txtLastNdataPoints;
    private TabPage ViewCalculation;
    private TabPage SearchCalculation;
    private Panel panel2;
    private Label label21;
    private Label label20;
    private Label label19;
    private Label label18;
    private Label label17;
    private Label label12;
    private Label label13;
    private Label label14;
    private Label label15;
    private Label label16;
    private DataGridView dataGridView2;
    private ComboBox cmbCalcID;
    private PictureBox picGraph;
    private Panel panel1;
    private Button btnSave;
    private CheckBox chkActive;
    private RadioButton rbAll;
    private RadioButton rbFitStat;
    private RadioButton rbNotFitStat;
    private Label label11;
    private Label label10;
    private TextBox txtNote;
    private TextBox txtUser;
    private TextBox txtTimePointData;
    private TextBox txtTimePointCalc;
    private TextBox txtNStatistically;
    private TextBox txtPercentStatistically;
    private TextBox txtDoNotFitStatistically;
    private Label label7;
    private Label label8;
    private Label label22;
    private Label label23;
    private Label label25;
    private Label label26;
    private Label label28;
    private Label label29;
    private Label label30;
    private DataGridView dataGridView1;
    private TextBox txtNParameterTotal;
    private Label label9;
    private Label label24;
    private Label label27;
    private Panel panelSelection;
    private Label label44;
    private ListBox lbGraph;
    private RadioButton rbKPI3;
    private RadioButton rbKPI2;
    private RadioButton rbKPI1;
    private RadioButton rbKPI0;
    private Label label45;
    private ListBox lbOzid;
    private Label label46;
    private Panel panel3;
    private GroupBox groupBox1;
    private CheckedListBox chkVirtOzid2;
    private DateTimePicker dtCalcDateTime;
    private ComboBox cmbProdID2;
    private RadioButton rbNotActive1;
    private RadioButton rbActive1;
    private ComboBox cmbCalcID2;
    private RadioButton rbAll1;
    private TextBox txtNote1;
    private Label label39;
    private Label label40;
    private Label label41;
    private Label label42;
    private GroupBox groupFilterSelection;
    private Button btnGetHistoric;
    private RadioButton rbNotFitStatF;
    private RadioButton rbFitstatF;
    private RadioButton rbAllF;
    private Panel panelButtons;
    private CheckBox chActivate;
    private CheckBox chDeActivate;
    private Button button2;
    private Button button3;
    private Panel panel4;
    private Label label48;
    private Panel panel6;
    private Panel panel17;
    private Label label31;
    private Panel panel16;
    private Label label32;
    private Panel panel15;
    private Label label33;
    private Label label34;
    private Panel panel13;
    private Label label35;
    private Panel panel12;
    private Label label36;
    private Panel panel11;
    private Label label37;
    private Panel panel10;
    private Label label38;
    private Panel panel5;
    private Label label47;
    private OpenFileDialog openFileDialog1;
    private PictureBox pictureBox2;
    private BackgroundWorker backgroundWorker2;
    private Timer timer1;
    private TextBox txProgressBar;
    private TabPage CompareCalculation;
    private Panel panel20;
    private Label label63;
    private Panel panel19;
    private Label label62;
    private Panel panel18;
    private Label label61;
    private Panel panel14;
    private Label label60;
    private Panel panel9;
    private Label label59;
    private Panel panel7;
    private Label label56;
    private Panel panel8;
    private Label label57;
    private GroupBox groupBox3;
    private CheckedListBox chkVirtOzid4;
    private DateTimePicker dtCalcDateTime4;
    private ComboBox cmbProdID4;
    private ComboBox cmbCalcID4;
    private TextBox txtNote4;
    private Label label52;
    private Label label53;
    private Label label54;
    private Label label55;
    private GroupBox groupBox2;
    private CheckedListBox chkVirtOzid3;
    private DateTimePicker dtCalcDateTime3;
    private ComboBox cmbProdID3;
    private ComboBox cmbCalcID3;
    private TextBox txtNote3;
    private Label label43;
    private Label label49;
    private Label label50;
    private Label label51;
    private Button btnCompareCalc;
    private PictureBox pictureBox4;
    private PictureBox pictureBox3;
    private Panel panel24;
    private Label label68;
    private Label label67;
    private Panel panel23;
    private Label label66;
    private Panel panel22;
    private Label label65;
    private Panel panel21;
    private Label label64;
    private Button btnZoomOut4;
    private Button btnZoom1;
    private Button btnZoomIn1;
    private Button btnZoomOut3;
    private Button btnPrint;
    private PrintDocument printDocument1;
    private Button btnPrint2;
    private PrintDocument printDocument2;
    private PrintDocument printDocument3;
    private PrintDocument printDocument4;
    private Button btnPrint4;
    private Button btnPrint3;
    private PrintDialog printDialog1;
    private Button button5;
    private Button button4;
    private Label label71;
    private Label label70;
    private Label label69;
    private Label label72;
    private Label label73;
    private Label label75;
    private Label label74;
    private Label label76;
    private Label label87;
    private Label label83;
    private Label label82;
    private Label label81;
    private Label label78;
    private Label label77;
    private Label label79;
    private Label label80;
    private Label label86;
    private Label label85;
    private Label label84;
    private Button button7;
    private Button button6;

    public Form1()
    {
      this.InitializeComponent();
      this.tabMain.SelectedIndexChanged += new EventHandler(this.tabMain_Click);
    }

    public bool IsInteger(string strNum) => int.TryParse(strNum, out int _);

    private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
    {
      try
      {
        DirectoryInfo directoryInfo1 = new DirectoryInfo(sourceDirName);
        DirectoryInfo[] directories = directoryInfo1.GetDirectories();
        if (!directoryInfo1.Exists)
          throw new DirectoryNotFoundException("Source directory does not exist or could not be found: " + sourceDirName);
        if (!Directory.Exists(destDirName))
          Directory.CreateDirectory(destDirName);
        foreach (FileInfo file in directoryInfo1.GetFiles())
        {
          string destFileName = Path.Combine(destDirName, file.Name);
          file.CopyTo(destFileName, false);
        }
        if (!copySubDirs)
          return;
        foreach (DirectoryInfo directoryInfo2 in directories)
        {
          string destDirName1 = Path.Combine(destDirName, directoryInfo2.Name);
          Form1.DirectoryCopy(directoryInfo2.FullName, destDirName1, copySubDirs);
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public string Left(string value, int maxLength)
    {
      if (string.IsNullOrEmpty(value))
        return value;
      maxLength = Math.Abs(maxLength);
      return value.Length <= maxLength ? value : value.Substring(0, maxLength);
    }

    public void CopyToSQLUniversal(DataGridView dtGrid, string[] arr, string TableName)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection))
          {
            for (int sourceColumnIndex = 0; sourceColumnIndex < ((IEnumerable<string>) arr).Count<string>(); ++sourceColumnIndex)
              sqlBulkCopy.ColumnMappings.Add(sourceColumnIndex, arr[sourceColumnIndex]);
            sqlBulkCopy.BatchSize = 800000;
            sqlBulkCopy.DestinationTableName = TableName;
            sqlBulkCopy.BulkCopyTimeout = 600;
            sqlBulkCopy.WriteToServer((DbDataReader) Form1.dt.CreateDataReader());
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public DataTable READExcel(string path)
    {
      try
      {
        using (Stream stream = (Stream) File.OpenRead(path))
        {
          using (ExcelEngine excelEngine = new ExcelEngine())
          {
            IWorksheet worksheet = excelEngine.Excel.Workbooks.Open(stream).Worksheets[0];
            DataTable dataTable = worksheet.ExportDataTable(worksheet.UsedRange["A1:AA300000"], ExcelExportDataTableOptions.ColumnNames);
            if (dataTable.Columns[0].ColumnName == "PRODUKTCODE")
            {
              dataTable.Columns.Add("SORT_DATE1", typeof (System.DateTime)).SetOrdinal(1);
              dataTable.Columns.Add("TS_ABS1", typeof (System.DateTime)).SetOrdinal(2);
              Form1.blnFlag = true;
            }
            else
            {
              dataTable.Columns.Add("SORT_DATE1", typeof (System.DateTime)).SetOrdinal(2);
              dataTable.Columns.Add("TS_ABS1", typeof (System.DateTime)).SetOrdinal(3);
              Form1.blnFlag = false;
            }
            foreach (DataRow row in (InternalDataCollectionBase) dataTable.Rows)
            {
              string s1 = row["SORT_DATE"].ToString();
              if (s1 == "01.01.0001 00:00:00")
                s1 = row["TS_ABS"].ToString();
              string s2 = row["TS_ABS"].ToString();
              if (s2 == "01.01.0001 00:00:00")
                s2 = row["TS_ABS"].ToString();
              row["TS_ABS1"] = (object) System.DateTime.Parse(s2);
              row["SORT_DATE1"] = (object) System.DateTime.Parse(s1);
            }
            dataTable.Columns.Remove("TS_ABS");
            dataTable.Columns.Remove("SORT_DATE");
            dataTable.Columns["TS_ABS1"].ColumnName = "TS_ABS";
            dataTable.Columns["SORT_DATE1"].ColumnName = "SORT_DATE";
            return dataTable;
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
        return (DataTable) null;
      }
    }

    public void CopyToSQL(DataGridView dtGrid)
    {
      try
      {
        DataTable dataSource = (DataTable) dtGrid.DataSource;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection))
          {
            sqlBulkCopy.ColumnMappings.Add(0, "ID");
            sqlBulkCopy.ColumnMappings.Add(1, "PRODUKTCODE");
            sqlBulkCopy.ColumnMappings.Add(2, "SORT_DATE");
            sqlBulkCopy.ColumnMappings.Add(3, "TS_ABS");
            sqlBulkCopy.ColumnMappings.Add(4, "LAUFNR");
            sqlBulkCopy.ColumnMappings.Add(5, "CHNR_ENDPRODUKT");
            sqlBulkCopy.ColumnMappings.Add(6, "PROCESS_CODE");
            sqlBulkCopy.ColumnMappings.Add(7, "PROCESS_CODE_NAME");
            sqlBulkCopy.ColumnMappings.Add(8, "PARAMETER_NAME");
            sqlBulkCopy.ColumnMappings.Add(9, "ASSAY");
            sqlBulkCopy.ColumnMappings.Add(10, "VIRT_OZID");
            sqlBulkCopy.ColumnMappings.Add(11, "TREND_WERT");
            sqlBulkCopy.ColumnMappings.Add(12, "TREND_WERT_2");
            sqlBulkCopy.ColumnMappings.Add(13, "ISTWERT_LIMS");
            sqlBulkCopy.ColumnMappings.Add(14, "LCL");
            sqlBulkCopy.ColumnMappings.Add(15, "UCL");
            sqlBulkCopy.ColumnMappings.Add(16, "CL");
            sqlBulkCopy.ColumnMappings.Add(17, "UAL");
            sqlBulkCopy.ColumnMappings.Add(18, "LAL");
            sqlBulkCopy.ColumnMappings.Add(19, "DECIMAL_PLACES_XCL_SUBSTITUTED");
            sqlBulkCopy.ColumnMappings.Add(20, "DECIMAL_PLACES_AL");
            sqlBulkCopy.ColumnMappings.Add(21, "DATA_TYPE");
            sqlBulkCopy.ColumnMappings.Add(22, "SOURCE_SYSTEM");
            sqlBulkCopy.ColumnMappings.Add(23, "EXCURSION");
            sqlBulkCopy.ColumnMappings.Add(24, "REFERENCED_CPV");
            sqlBulkCopy.ColumnMappings.Add(25, "IS_IN_RUN_NUMBER_RANGE");
            sqlBulkCopy.ColumnMappings.Add(26, "LOCATION");
            sqlBulkCopy.BatchSize = 800000;
            sqlBulkCopy.DestinationTableName = "Products";
            sqlBulkCopy.BulkCopyTimeout = 600;
            sqlBulkCopy.WriteToServer((DbDataReader) dataSource.CreateDataReader());
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public void CopyToSQL2(DataGridView dtGrid)
    {
      try
      {
        DataTable dataSource = (DataTable) dtGrid.DataSource;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection))
          {
            sqlBulkCopy.ColumnMappings.Add(0, "PRODUKTCODE");
            sqlBulkCopy.ColumnMappings.Add(1, "SORT_DATE");
            sqlBulkCopy.ColumnMappings.Add(2, "TS_ABS");
            sqlBulkCopy.ColumnMappings.Add(3, "LAUFNR");
            sqlBulkCopy.ColumnMappings.Add(4, "CHNR_ENDPRODUKT");
            sqlBulkCopy.ColumnMappings.Add(5, "PROCESS_CODE");
            sqlBulkCopy.ColumnMappings.Add(6, "PROCESS_CODE_NAME");
            sqlBulkCopy.ColumnMappings.Add(7, "PARAMETER_NAME");
            sqlBulkCopy.ColumnMappings.Add(8, "ASSAY");
            sqlBulkCopy.ColumnMappings.Add(9, "VIRT_OZID");
            sqlBulkCopy.ColumnMappings.Add(10, "TREND_WERT");
            sqlBulkCopy.ColumnMappings.Add(11, "TREND_WERT_2");
            sqlBulkCopy.ColumnMappings.Add(12, "ISTWERT_LIMS");
            sqlBulkCopy.ColumnMappings.Add(13, "LCL");
            sqlBulkCopy.ColumnMappings.Add(14, "UCL");
            sqlBulkCopy.ColumnMappings.Add(15, "CL");
            sqlBulkCopy.ColumnMappings.Add(16, "UAL");
            sqlBulkCopy.ColumnMappings.Add(17, "LAL");
            sqlBulkCopy.ColumnMappings.Add(18, "DECIMAL_PLACES_XCL_SUBSTITUTED");
            sqlBulkCopy.ColumnMappings.Add(19, "DECIMAL_PLACES_AL");
            sqlBulkCopy.ColumnMappings.Add(20, "DATA_TYPE");
            sqlBulkCopy.ColumnMappings.Add(21, "SOURCE_SYSTEM");
            sqlBulkCopy.ColumnMappings.Add(22, "EXCURSION");
            sqlBulkCopy.ColumnMappings.Add(23, "REFERENCED_CPV");
            sqlBulkCopy.ColumnMappings.Add(24, "IS_IN_RUN_NUMBER_RANGE");
            sqlBulkCopy.ColumnMappings.Add(25, "LOCATION");
            sqlBulkCopy.BatchSize = 800000;
            sqlBulkCopy.DestinationTableName = "Products";
            sqlBulkCopy.BulkCopyTimeout = 600;
            DataColumnCollection columns = dataSource.Columns;
            if (columns.Contains("Column1"))
              dataSource.Columns.Remove("Column1");
            if (columns.Contains("Column2"))
              dataSource.Columns.Remove("Column2");
            if (columns.Contains("ID"))
              dataSource.Columns.Remove("ID");
            sqlBulkCopy.WriteToServer((DbDataReader) dataSource.CreateDataReader());
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void timer1_Tick(object sender, EventArgs e)
    {
      try
      {
        if (!Form1.stopCalc)
          return;
        if (this.txProgressBar.Text.Length == 100)
        {
          this.txProgressBar.Visible = false;
          this.txProgressBar.Text = "";
        }
        else
        {
          this.txProgressBar.Visible = true;
          if (Form1.last == 1)
          {
            this.txProgressBar.Text += "█";
            Form1.last = 2;
          }
          else
          {
            this.txProgressBar.Text += "█";
            Form1.last = 1;
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void strH()
    {
      try
      {
        string str1 = Form1.strOutputDir.Replace("/", "\\");
        string str2 = str1 + Form1.strCalcID + "\\";
        DirectoryInfo directoryInfo = new DirectoryInfo(str1 + Form1.strCalcID + "\\");
        List<string> stringList = new List<string>();
        List<string> list = ((IEnumerable<FileInfo>) directoryInfo.GetFiles("*.xlsx")).Where<FileInfo>((System.Func<FileInfo, bool>) (file => file.Name.EndsWith(".xlsx"))).Select<FileInfo, string>((System.Func<FileInfo, string>) (file => file.Name)).ToList<string>();
        int num = list.Count<string>();
        for (int index = 0; index < num; ++index)
        {
          if (list[index].ToString().Length < 35)
            this.listBox2.Items.Add((object) list[index].ToString());
        }
        Directory.GetCurrentDirectory();
        Form1.dt = this.READExcel(str1 + Form1.strCalcID + "\\" + list[0].ToString());
        string[] arr = new string[27]
        {
          "PRODUKTCODE",
          "TREND_WERT",
          "TREND_WERT_2",
          "CL",
          "LCL",
          "UCL",
          "LAL",
          "UAL",
          "TS_ABS",
          "SORT_DATE",
          "LAUFNR",
          "EXCURSION",
          "VIRT_OZID",
          "VALUE",
          "BatchID",
          "lowSD",
          "uppSD",
          "mu",
          "sigma",
          "upp",
          "delta",
          "rSigma",
          "valid",
          "signal",
          "lag",
          "Column1",
          "Column2"
        };
        string TableName = "CalculationResult";
        this.dataGridViewTemp.Columns.Clear();
        this.dataGridViewTemp.Columns.Add("PRODUKTCODE", "");
        this.dataGridViewTemp.Columns.Add("TREND_WERT", "");
        this.dataGridViewTemp.Columns.Add("TREND_WERT_2", "");
        this.dataGridViewTemp.Columns.Add("CL", "");
        this.dataGridViewTemp.Columns.Add("LCL", "");
        this.dataGridViewTemp.Columns.Add("UCL", "");
        this.dataGridViewTemp.Columns.Add("LAL", "");
        this.dataGridViewTemp.Columns.Add("UAL", "");
        this.dataGridViewTemp.Columns.Add("TS_ABS", "");
        this.dataGridViewTemp.Columns.Add("SORT_DATE", "");
        this.dataGridViewTemp.Columns.Add("LAUFNR", "");
        this.dataGridViewTemp.Columns.Add("EXCURSION", "");
        this.dataGridViewTemp.Columns.Add("VIRT_OZID", "");
        this.dataGridViewTemp.Columns.Add("VALUE", "");
        this.dataGridViewTemp.Columns.Add("BatchID", "");
        this.dataGridViewTemp.Columns.Add("lowSD", "");
        this.dataGridViewTemp.Columns.Add("uppSD", "");
        this.dataGridViewTemp.Columns.Add("mu", "");
        this.dataGridViewTemp.Columns.Add("sigma", "");
        this.dataGridViewTemp.Columns.Add("upp", "");
        this.dataGridViewTemp.Columns.Add("delta", "");
        this.dataGridViewTemp.Columns.Add("rSigma", "");
        this.dataGridViewTemp.Columns.Add("valid", "");
        this.dataGridViewTemp.Columns.Add("signal", "");
        this.dataGridViewTemp.Columns.Add("lag", "");
        this.dataGridViewTemp.Columns.Add("Column1", "");
        this.dataGridViewTemp.Columns.Add("Column2", "");
        this.CopyToSQLUniversal(this.dataGridViewTemp, arr, TableName);
        SqlConnection connection = new SqlConnection(Form1.connectionString);
        connection.Open();
        System.DateTime utcNow = System.DateTime.UtcNow;
        string cmdText = "insert into CalculationRaw select '" + Form1.strCalcID + "','" + utcNow.ToString() + "', [PRODUKTCODE],[TREND_WERT],[TREND_WERT_2],[CL],[LCL],[UCL],[LAL],[UAL],[TS_ABS],[SORT_DATE],[LAUFNR],[EXCURSION],[VIRT_OZID],[VALUE],[BatchID],[lowSD],[uppSD],[mu],[sigma],[upp],[delta],[rSigma],[valid],[signal],[lag] from CalculationResult";
        SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
        new SqlCommand(cmdText, connection).ExecuteNonQuery();
        new SqlCommand("insert into[dbo].[CalcRow] SELECT distinct  calcid,VIRT_OZID, count(signal) as totaln,dbo.PercentStatisticallyFit0_Ozid(calcid, VIRT_OZID) as procent0, dbo.PercentStatisticallyFit1_Ozid(calcid, VIRT_OZID) as procent1,dbo.PercentStatisticallyFit2_Ozid(calcid, VIRT_OZID) as procent2, dbo.PercentStatisticallyFit3_Ozid(calcid, VIRT_OZID) as procent3, 'False',  'False',''  FROM [dbo].[CalculationRaw]  where calcid = '" + Form1.strCalcID + "' group by VIRT_OZID, calcid", connection).ExecuteNonQuery();
        new SqlCommand("insert into[dbo].[CalcRowSearch] SELECT distinct  calcid,VIRT_OZID, dbo.KPIcount0(calcid, VIRT_OZID) as totaln0,dbo.KPIcount1(calcid, VIRT_OZID) as totaln1,dbo.KPIcount2(calcid, VIRT_OZID) as totaln2,dbo.KPIcount3(calcid, VIRT_OZID) as totaln3,dbo.PercentStatisticallyFit0_Ozid(calcid, VIRT_OZID) as procent0, dbo.PercentStatisticallyFit1_Ozid(calcid, VIRT_OZID) as procent1,dbo.PercentStatisticallyFit2_Ozid(calcid, VIRT_OZID) as procent2, dbo.PercentStatisticallyFit3_Ozid(calcid, VIRT_OZID) as procent3, 'False',  'False','False',''   FROM[dbo].[CalculationRaw]  where calcid = '" + Form1.strCalcID + "' group by VIRT_OZID, calcid", connection).ExecuteNonQuery();
        connection.Close();
        CultureInfo cultureInfo = new CultureInfo("de-DE");
        Form1.stopCalc = false;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public static byte[] GetPhoto(string filePath)
    {
      try
      {
        FileStream input = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        BinaryReader binaryReader = new BinaryReader((Stream) input);
        byte[] photo = binaryReader.ReadBytes((int) input.Length);
        binaryReader.Close();
        input.Close();
        return photo;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
        return (byte[]) null;
      }
    }

    private void btnCalculation_Click(object sender, EventArgs e)
    {
      try
      {
        if (MessageBox.Show("Are you sure to start calculation Yes/No", "", MessageBoxButtons.YesNo) != DialogResult.Yes)
          return;
        try
        {
          this.txProgressBar.Visible = true;
          this.txProgressBar.Text = "";
          this.timer1.Tick += new EventHandler(this.timer1_Tick);
          this.timer1.Start();
          this.backgroundWorker2.RunWorkerAsync();
          this.txtResult.Visible = false;
          this.txtResult.Text = "";
          SqlConnection connection1 = new SqlConnection(Form1.connectionString);
          connection1.Open();
          System.DateTime utcNow = System.DateTime.UtcNow;
          string str1 = "";
          new SqlCommand("delete from CalculationResult", connection1).ExecuteNonQuery();
          connection1.Close();
          string strRscript1 = Form1.strRscript;
          string strRscript2 = Form1.strRscript;
          string strRpath1 = Form1.strRpath;
          string empty = string.Empty;
          try
          {
            if (Form1.strExcelFileName != null)
            {
              ProcessStartInfo processStartInfo = new ProcessStartInfo();
              processStartInfo.FileName = strRscript1;
              processStartInfo.WorkingDirectory = Path.GetDirectoryName(strRpath1);
              Form1.strCalcID = System.DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss", (IFormatProvider) CultureInfo.InvariantCulture).ToString().Replace("_", "").Replace(" ", "").Replace("-", "").Replace(".", "").Replace(":", "");
              string path1 = Form1.strOutputDir + Form1.strCalcID;
              Directory.CreateDirectory(path1);
              string[] strArray1 = new string[3]
              {
                Form1.strExcelFileName.Replace("//", "/"),
                this.cmbProductCode.Text,
                path1 + "/"
              };
              strArray1[0] = Form1.strExcelFileName.Replace("//", "/");
              strRscript2.Replace("\\\\", "\\");
              processStartInfo.Arguments = strRpath1 + " " + strArray1[0] + " " + strArray1[1] + " " + strArray1[2];
              using (new StreamWriter(strRpath1, true))
              {
                Form1.strFileName = this.openFileDialog1.FileName;
                this.btnOpenFile.BackColor = Color.LightBlue;
                DataTable dataTable1 = new DataTable();
                DataTable dataTable2 = new DataTable();
                Form1.dt = (DataTable) this.dataGridViewRaw.DataSource;
                Form1.dt = this.READExcel(Form1.strExcelFileName.Replace("//", "/"));
                Form1.dt.Columns.Add("UserID");
                Form1.dt.Columns.Add("ModifiedDate");
                Form1.dt.Columns.Add("GraphID");
                Form1.dt.Columns.Add("CalcID");
                Form1.dt.Columns.Add("FilterID");
              }
              string[] arr = new string[32]
              {
                "ID",
                "PRODUKTCODE",
                "SORT_DATE",
                "TS_ABS",
                "LAUFNR",
                "CHNR_ENDPRODUKT",
                "PROCESS_CODE",
                "PROCESS_CODE_NAME",
                "PARAMETER_NAME",
                "ASSAY",
                "VIRT_OZID",
                "TREND_WERT",
                "TREND_WERT_2",
                "ISTWERT_LIMS",
                "LCL",
                "UCL",
                "CL",
                "UAL",
                "LAL",
                "DECIMAL_PLACES_XCL_SUBSTITUTED",
                "DECIMAL_PLACES_AL",
                "DATA_TYPE",
                "SOURCE_SYSTEM",
                "EXCURSION",
                "REFERENCED_CPV",
                "IS_IN_RUN_NUMBER_RANGE",
                "LOCATION",
                "UserID",
                "ModifiedDate",
                "GraphID",
                "CalcID",
                "FilterID"
              };
              string TableName = "ProductsFilteredTemp";
              if (Form1.strHasID == "1")
                this.dataGridViewTemp.Columns.Add("ID", "");
              this.dataGridViewTemp.Columns.Add("PRODUKTCODE", "");
              this.dataGridViewTemp.Columns.Add("SORT_DATE", "");
              this.dataGridViewTemp.Columns.Add("TS_ABS", "");
              this.dataGridViewTemp.Columns.Add("LAUFNR", "");
              this.dataGridViewTemp.Columns.Add("CHNR_ENDPRODUKT", "");
              this.dataGridViewTemp.Columns.Add("PROCESS_CODE", "");
              this.dataGridViewTemp.Columns.Add("PROCESS_CODE_NAME", "");
              this.dataGridViewTemp.Columns.Add("PARAMETER_NAME", "");
              this.dataGridViewTemp.Columns.Add("ASSAY", "");
              this.dataGridViewTemp.Columns.Add("VIRT_OZID", "");
              this.dataGridViewTemp.Columns.Add("TREND_WERT", "");
              this.dataGridViewTemp.Columns.Add("TREND_WERT_2", "");
              this.dataGridViewTemp.Columns.Add("ISTWERT_LIMS", "");
              this.dataGridViewTemp.Columns.Add("LCL", "");
              this.dataGridViewTemp.Columns.Add("UCL", "");
              this.dataGridViewTemp.Columns.Add("CL", "");
              this.dataGridViewTemp.Columns.Add("UAL", "");
              this.dataGridViewTemp.Columns.Add("LAL", "");
              this.dataGridViewTemp.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", "");
              this.dataGridViewTemp.Columns.Add("DECIMAL_PLACES_AL", "");
              this.dataGridViewTemp.Columns.Add("DATA_TYPE", "");
              this.dataGridViewTemp.Columns.Add("SOURCE_SYSTEM", "");
              this.dataGridViewTemp.Columns.Add("EXCURSION", "");
              this.dataGridViewTemp.Columns.Add("REFERENCED_CPV", "");
              this.dataGridViewTemp.Columns.Add("IS_IN_RUN_NUMBER_RANGE", "");
              this.dataGridViewTemp.Columns.Add("LOCATION", "");
              this.dataGridViewTemp.Columns.Add("UserID", "");
              this.dataGridViewTemp.Columns.Add("ModifiedDate", "");
              this.dataGridViewTemp.Columns.Add("GraphID", "");
              this.dataGridViewTemp.Columns.Add("CalcID", Form1.strCalcID);
              this.dataGridViewTemp.Columns.Add("FilterID", "");
              this.CopyToSQLUniversal(this.dataGridViewTemp, arr, TableName);
              SqlConnection connection2 = new SqlConnection(Form1.connectionString);
              connection1.Open();
              connection2.Open();
              int num1 = new SqlCommand("update ProductsFilteredTemp set CalcID = '" + Form1.strCalcID + "'", connection2).ExecuteNonQuery();
              string connectionString = Form1.connectionString;
              SqlConnection connection3 = new SqlConnection(Form1.connectionString);
              connection3.Open();
              if (Form1.strHasID == "1")
              {
                string[] strArray2 = new string[32]
                {
                  "ID",
                  "PRODUKTCODE",
                  "SORT_DATE",
                  "TS_ABS",
                  "LAUFNR",
                  "CHNR_ENDPRODUKT",
                  "PROCESS_CODE",
                  "PROCESS_CODE_NAME",
                  "PARAMETER_NAME",
                  "ASSAY",
                  "VIRT_OZID",
                  "TREND_WERT",
                  "TREND_WERT_2",
                  "ISTWERT_LIMS",
                  "LCL",
                  "UCL",
                  "CL",
                  "UAL",
                  "LAL",
                  "DECIMAL_PLACES_XCL_SUBSTITUTED",
                  "DECIMAL_PLACES_AL",
                  "DATA_TYPE",
                  "SOURCE_SYSTEM",
                  "EXCURSION",
                  "REFERENCED_CPV",
                  "IS_IN_RUN_NUMBER_RANGE",
                  "LOCATION",
                  "UserID",
                  "ModifiedDate",
                  "GraphID",
                  "CalcID",
                  "FilterID"
                };
              }
              if (Form1.strHasID == "0")
              {
                string[] strArray3 = new string[31]
                {
                  "PRODUKTCODE",
                  "SORT_DATE",
                  "TS_ABS",
                  "LAUFNR",
                  "CHNR_ENDPRODUKT",
                  "PROCESS_CODE",
                  "PROCESS_CODE_NAME",
                  "PARAMETER_NAME",
                  "ASSAY",
                  "VIRT_OZID",
                  "TREND_WERT",
                  "TREND_WERT_2",
                  "ISTWERT_LIMS",
                  "LCL",
                  "UCL",
                  "CL",
                  "UAL",
                  "LAL",
                  "DECIMAL_PLACES_XCL_SUBSTITUTED",
                  "DECIMAL_PLACES_AL",
                  "DATA_TYPE",
                  "SOURCE_SYSTEM",
                  "EXCURSION",
                  "REFERENCED_CPV",
                  "IS_IN_RUN_NUMBER_RANGE",
                  "LOCATION",
                  "UserID",
                  "ModifiedDate",
                  "GraphID",
                  "CalcID",
                  "FilterID"
                };
              }
              SqlCommand sqlCommand1 = new SqlCommand("dataIn", connection3);
              sqlCommand1.CommandType = CommandType.StoredProcedure;
              sqlCommand1.ExecuteNonQuery();
              connection3.Close();
              num1 = new SqlCommand("delete from ProductsFilteredTemp ", connection2).ExecuteNonQuery();
              processStartInfo.RedirectStandardInput = false;
              processStartInfo.RedirectStandardOutput = true;
              processStartInfo.UseShellExecute = false;
              processStartInfo.CreateNoWindow = true;
              SqlConnection connection4 = new SqlConnection(Form1.connectionString);
              connection4.Open();
              str1 = "";
              new SqlCommand("delete from TempParams", connection4).ExecuteNonQuery();
              new SqlCommand("insert into TempParams select '" + strArray1[0] + "','" + strArray1[1] + "','" + strArray1[2] + "'", connection4).ExecuteNonQuery();
              connection4.Close();
              try
              {
                string[] strArray4 = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini");
                Form1.connectionString = strArray4[0];
                Form1.strRscript = strArray4[1];
                Form1.strRpath = strArray4[2];
                string strRpath2 = Form1.strRpath;
                Form1.strDataDir = strArray4[3];
                Form1.strOutputDir = strArray4[4];
                string path2 = strArray4[2];
                string[] contents = File.ReadAllLines(path2);
                using (SqlConnection connection5 = new SqlConnection(Form1.connectionString))
                {
                  connection5.Open();
                  str1 = "";
                  string cmdText = "select distinct * from TempParams";
                  SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection5);
                  sqlCommand2.ExecuteScalar();
                  SqlDataReader sqlDataReader = sqlCommand2.ExecuteReader();
                  SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection5);
                  while (sqlDataReader.Read())
                  {
                    Form1.strArg[0] = sqlDataReader[0].ToString();
                    Form1.strArg[1] = sqlDataReader[1].ToString();
                    Form1.strArg[2] = sqlDataReader[2].ToString();
                  }
                  sqlDataReader.Close();
                  connection5.Close();
                  try
                  {
                    Console.WriteLine("started");
                  }
                  catch (Exception ex)
                  {
                    Console.WriteLine(ex.Message);
                  }
                }
                contents[15] = "filename <-paste0('" + Form1.strArg[0].ToString() + "')";
                contents[16] = "prodcode <-paste0('" + Form1.strArg[1].ToString() + "')";
                contents[17] = "OutputDir <-paste0('" + Form1.strArg[2].ToString() + "')";
                File.WriteAllLines(path2, contents);
                REngine.SetEnvironmentVariables(Directory.GetCurrentDirectory() + "\\bin\\x64", Directory.GetCurrentDirectory());
                this.engine = REngine.GetInstance();
                (strRpath2 + "\\R_LIBS_USER").Replace("\\", "/");
                this.engine.Evaluate("source('" + path2.Replace("\\", "/") + "')");
                this.txtResult.Text = "Calculation is finished, the number of calculation is " + Form1.strCalcID;
                this.txProgressBar.Visible = false;
                this.btnCalculationView.Enabled = true;
                this.btnCalculationSearch.Enabled = true;
              }
              catch (Exception ex)
              {
                int num2 = (int) MessageBox.Show(ex.Message);
              }
              try
              {
                this.strH();
              }
              catch (Exception ex)
              {
                int num3 = (int) MessageBox.Show(ex.Message);
              }
              this.txtResult.Visible = true;
              empty.Replace("# A tibble: 0 Ã—", "");
              string[] strArray5 = new string[1]{ "" };
              SqlConnection connection6 = new SqlConnection(Form1.connectionString);
              connection6.Open();
              int length = Directory.GetFiles(strArray1[2], "*.jpeg", SearchOption.AllDirectories).Length;
              if (length > 0)
              {
                string[] files = Directory.GetFiles(strArray1[2], "*.jpeg");
                for (int index = 0; index < length; ++index)
                {
                  string text = this.cmbProductCode.Text;
                  string str2 = files[index];
                  int num4 = str2.IndexOf(text);
                  string str3 = str2.Substring(num4 + text.Length + 1);
                  string str4 = str3.Substring(0, str3.Length - 23);
                  string str5 = str4;
                  System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
                  Form1.GetPhoto(files[index]);
                  try
                  {
                    byte[] numArray = File.ReadAllBytes(files[index]);
                    SqlCommand sqlCommand4 = new SqlCommand("insert  into Graphs([GraphName],[VIRT_OZID],[CalcID],[ImageValue], [ID]) values (@GraphName,@VIRT_OZID,@CalcID,@ImageValue, @ID)", connection6);
                    sqlCommand4.Parameters.AddWithValue("@GraphName", (object) str4);
                    sqlCommand4.Parameters.AddWithValue("@VIRT_OZID", (object) str5);
                    sqlCommand4.Parameters.AddWithValue("@CalcID", (object) Form1.strCalcID.ToString());
                    sqlCommand4.Parameters.AddWithValue("@ImageValue", (object) numArray);
                    sqlCommand4.Parameters.AddWithValue("@ID", (object) (Form1.strCalcID.ToString() + str5));
                    sqlCommand4.ExecuteNonQuery();
                  }
                  catch (Exception ex)
                  {
                    int num5 = (int) MessageBox.Show(ex.Message);
                    connection6.Close();
                  }
                }
                this.timer1.Enabled = true;
                this.timer1_Tick(sender, e);
                DirectoryInfo directoryInfo = new DirectoryInfo(Form1.strOutputDir.Replace("/", "\\"));
                foreach (FileSystemInfo file in directoryInfo.GetFiles())
                  file.Delete();
                foreach (DirectoryInfo directory in directoryInfo.GetDirectories())
                  directory.Delete(true);
                connection6.Close();
              }
              if (length != 0)
                this.txtResult.Text = "Calculation is finished, the number of calculation is " + Form1.strCalcID;
              else
                this.txtResult.Text = "Please repeat calculations with other parameters, over insufficient data, there is no files in the output folder " + Form1.strOutputDir + Form1.strCalcID;
              this.txProgressBar.Visible = false;
            }
            else
            {
              int num = (int) MessageBox.Show("Please export the filtered viewgrid data to the excel file first! ");
            }
          }
          catch (Exception ex)
          {
            int num = (int) MessageBox.Show(ex.Message);
          }
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.Message);
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnReset_Click(object sender, EventArgs e)
    {
      this.chLastN.Enabled = true;
      this.chLastN.Checked = false;
      this.chkAllRefCpv.Checked = true;
      this.chkAllVirtOzid.Checked = true;
      this.dtSortDateFrom.Text = Form1.strInitialDate;
      this.dtSortDateTo.Text = Form1.strEndDate;
      this.chkSortDate.Enabled = true;
      this.chkSortDate.Checked = false;
      this.chkExclVirtOzid.Checked = false;
      this.dtSortDateFrom.Enabled = true;
      this.dtSortDateTo.Enabled = true;
      this.chkSortDate.Enabled = true;
      this.cmbLaufnrMin.Enabled = true;
      this.cmbLaufnrMax.Enabled = true;
      this.chkLaufNr.Enabled = true;
      this.txtLastNdataPoints.Enabled = true;
      this.chLastN.Enabled = true;
      this.txtLastMdataPoints.Enabled = true;
      this.chLastM.Enabled = true;
      this.rbDays.Enabled = true;
      this.rbWeeks.Enabled = true;
      this.rbMonths.Enabled = true;
      this.rbYears.Enabled = true;
      this.lstCheckExclVirtOzid.Items.Clear();
      this.listBox1.Items.Clear();
      this.txtLastNdataPoints.Text = "5000";
      this.txtLastMdataPoints.Text = "3";
      this.rbYears.Enabled = true;
      this.chkSortDate.Enabled = true;
      this.chkSortDate.Checked = false;
      this.chLastM.Enabled = true;
      this.chLastM.Checked = false;
      for (int index = 0; index < this.clbVirtOzid.Items.Count; ++index)
        this.clbVirtOzid.SetItemChecked(index, false);
      this.chkAllRefCpv.Checked = true;
      this.chkAllVirtOzid.Checked = true;
      this.cmbProductCode.Focus();
      Form1.filterOn = false;
      this.lblDataGridTitle.Text = "Table of raw data entity (Original: Yes  -  Filtered: NO ) ";
      this.btnCalculationView.Visible = true;
      this.chkSortDate.Enabled = false;
      this.clbRefCpv.Enabled = false;
      this.clbVirtOzid.Enabled = false;
      this.dataGridViewRaw.Visible = true;
      try
      {
        SqlConnection connection = new SqlConnection(Form1.connectionString);
        SqlCommand selectCommand = new SqlCommand("Select * From Products where ProduktCode='" + this.cmbProductCode.Text + "'", connection);
        connection.Open();
        new SqlDataAdapter(selectCommand).Fill(Form1.dt);
        this.dataGridViewRaw.DataSource = (object) Form1.dt;
        CultureInfo cultureInfo = new CultureInfo("de-DE");
        connection.Close();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public void DataTableToExcel(DataTable dt)
    {
      try
      {
        string sheetName = "Records";
        string strDataDir = Form1.strDataDir;
        if (!Directory.Exists(strDataDir))
          Directory.CreateDirectory(strDataDir);
        using (XLWorkbook xlWorkbook = new XLWorkbook())
        {
          xlWorkbook.Worksheets.Add(dt, sheetName);
          string str = System.DateTime.Now.ToString().Replace(".", "-");
          Form1.strExcelFileName = strDataDir + "\\" + this.cmbProductCode.Text + str.ToString().Replace(":", "-").Replace(" ", "") + ".xlsx";
          xlWorkbook.SaveAs(Form1.strExcelFileName);
          Form1.strExcelFileName = Form1.strExcelFileName.Replace("\\", "/");
          Form1.strExcelFileName = Form1.strExcelFileName.Replace("//", "/");
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void button1_Click_1(object sender, EventArgs e)
    {
      try
      {
        if (!Form1.filterIsOn)
        {
          int num1 = (int) MessageBox.Show("Please filter the data first!");
        }
        else
        {
          string str1 = "SELECT ID,PRODUKTCODE,SORT_DATE,TS_ABS,LAUFNR,CHNR_ENDPRODUKT,PROCESS_CODE,PROCESS_CODE_NAME,PARAMETER_NAME,ASSAY,VIRT_OZID,TREND_WERT,TREND_WERT_2,ISTWERT_LIMS,LCL,UCL,CL,UAL,LAL,DECIMAL_PLACES_XCL_SUBSTITUTED,DECIMAL_PLACES_AL,DATA_TYPE,SOURCE_SYSTEM,EXCURSION,REFERENCED_CPV,IS_IN_RUN_NUMBER_RANGE,LOCATION FROM [dbo].[Products]";
          string str2 = this.cmbProductCode.Text + System.DateTime.Now.ToString().Replace(":", "-") + ".csv";
          string connectionString = Form1.connectionString;
          DataTable dataTable = new DataTable("dataGridViewRaw");
          SqlConnection connection = new SqlConnection(connectionString.ToString().Replace("//", "/"));
          Form1.strSQLfiltered = Form1.strSQLfiltered.Replace("and Virt_Ozid in )", "");
          SqlDataAdapter adapter = new SqlDataAdapter();
          if (Form1.strSQLfiltered == "")
            Form1.strSQLfiltered = str1;
          adapter.SelectCommand = new SqlCommand(Form1.strSQLfiltered, connection);
          SqlCommandBuilder sqlCommandBuilder = new SqlCommandBuilder(adapter);
          adapter.Fill(dataTable);
          this.DataTableToExcel(dataTable);
        }
        try
        {
          string path1 = WindowsIdentity.GetCurrent().Name.Replace("\\", "=");
          foreach (string str in ((IEnumerable<string>) path1.Split(new char[1]
          {
            '='
          }, StringSplitOptions.RemoveEmptyEntries)).Select<string, string>((System.Func<string, string>) (x => x.Trim())))
            path1 = str;
          string path2 = "C:\\Dokumente und Einstellungen\\" + path1 + "\\AppData\\Local\\R\\win-library\\4.2";
          if (!Directory.Exists(path2))
          {
            Directory.CreateDirectory(path1);
            Form1.DirectoryCopy(".\\rlib", "C:\\Dokumente und Einstellungen\\" + path1 + "\\AppData\\Local\\R\\win-library\\4.2", true);
          }
          else
          {
            DirectoryInfo directoryInfo = new DirectoryInfo(path2);
            Directory.GetFiles(path2, "*");
            foreach (string directory in Directory.GetDirectories(path2, "*.*"))
              Directory.Delete(directory, true);
            int num2 = (int) MessageBox.Show("It is started to copy libraries to C:\\Dokumente und Einstellungen\\" + path1 + "\\AppData\\Local\\R\\win-library\\4.2 folder.");
            Form1.DirectoryCopy(".\\rlib", "C:\\Dokumente und Einstellungen\\" + path1 + "\\AppData\\Local\\R\\win-library\\4.2", true);
          }
        }
        catch (Exception ex)
        {
          int num3 = (int) MessageBox.Show(ex.Message);
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnFilterData_Click_1(object sender, EventArgs e)
    {
      try
      {
        SqlConnection connection1 = new SqlConnection(Form1.connectionString);
        SqlCommand selectCommand = new SqlCommand("Select * From Products where ProduktCode='" + this.cmbProductCode.Text + "'", connection1);
        connection1.Open();
        new SqlDataAdapter(selectCommand).Fill(Form1.dt);
        this.dataGridViewRaw.DataSource = (object) Form1.dt;
        CultureInfo cultureInfo = new CultureInfo("de-DE");
        connection1.Close();
        Form1.filterIsOn = true;
        Form1.intMaxPoints = 200000f;
        Form1.strFilterExclude = "";
        string str1 = "";
        using (SqlConnection connection2 = new SqlConnection(Form1.connectionString))
        {
          connection2.Open();
          string cmdText1 = "select min(CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 )) from Products where laufnr = '" + this.cmbLaufnrMin.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText1, connection2);
          object obj = sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader1 = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText1, connection2);
          while (sqlDataReader1.Read())
            Form1.strLaufNRDateMin = sqlDataReader1[0].ToString();
          sqlDataReader1.Close();
          string cmdText2 = "select max(CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 )) from Products where laufnr = '" + this.cmbLaufnrMax.Text + "'";
          SqlCommand sqlCommand3 = new SqlCommand(cmdText2, connection2);
          obj = sqlCommand3.ExecuteScalar();
          SqlDataReader sqlDataReader2 = sqlCommand3.ExecuteReader();
          sqlCommand2 = new SqlCommand(cmdText2, connection2);
          while (sqlDataReader2.Read())
            Form1.strLaufNRDateMax = sqlDataReader2[0].ToString();
          sqlDataReader2.Close();
          connection2.Close();
        }
        str1 = "";
        string str2 = " (1 = 1)  and 0=0";
        if (!this.chkAllVirtOzid.Checked)
        {
          string str3 = " and Virt_Ozid in ('";
          for (int index = 0; index < this.clbVirtOzid.Items.Count; ++index)
          {
            if (this.clbVirtOzid.GetItemCheckState(index) == CheckState.Checked)
              str3 = str3 + (string) this.clbVirtOzid.Items[index] + "','";
          }
          string str4 = this.Left(str3, str3.Length - 2) + ") ";
          if (str4.IndexOf("()") <= 0)
            str2 += str4;
        }
        if (!this.chkAllRefCpv.Checked)
        {
          string str5 = " and REFERENCED_CPV in ('";
          for (int index = 0; index < this.clbRefCpv.Items.Count; ++index)
          {
            if (this.clbRefCpv.GetItemCheckState(index) == CheckState.Checked)
              str5 = str5 + (string) this.clbRefCpv.Items[index] + "','";
          }
          string str6 = this.Left(str5, str5.Length - 2) + ") ";
          if (str6.IndexOf("()") <= 0)
            str2 += str6;
        }
        Form1.strFilterLaufnr = "";
        string str7 = "";
        string str8 = "";
        using (SqlConnection connection3 = new SqlConnection(Form1.connectionString))
        {
          string cmdText = "select distinct min(TS_ABS) from Products where laufnr = '" + this.cmbLaufnrMin.Text + "'";
          connection3.Open();
          SqlCommand sqlCommand = new SqlCommand(cmdText, connection3);
          sqlCommand.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          while (sqlDataReader.Read())
            str7 = sqlDataReader[0].ToString();
          sqlDataReader.Close();
          connection3.Close();
        }
        using (SqlConnection connection4 = new SqlConnection(Form1.connectionString))
        {
          string cmdText = "select distinct max(TS_ABS) from Products where laufnr = '" + this.cmbLaufnrMax.Text + "'";
          connection4.Open();
          SqlCommand sqlCommand = new SqlCommand(cmdText, connection4);
          sqlCommand.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          while (sqlDataReader.Read())
            str8 = sqlDataReader[0].ToString();
          sqlDataReader.Close();
          connection4.Close();
        }
        cultureInfo = CultureInfo.CreateSpecificCulture("fr-FR");
        if (this.cmbLaufnrMin.Text != "" && this.cmbLaufnrMax.Text != "" && this.chkLaufNr.Checked)
          Form1.strFilterLaufnr = Form1.strFilterLaufnr + " and TS_ABS between CONVERT(Datetime,'" + str7 + "', 103) and CONVERT(Datetime,'" + str8 + "', 103) ";
        string str9 = "";
        string str10 = "";
        using (SqlConnection connection5 = new SqlConnection(Form1.connectionString))
        {
          connection5.Open();
          SqlCommand sqlCommand = new SqlCommand("select max(TS_ABS) FROM [dbo].[Products]", connection5);
          sqlCommand.ExecuteScalar();
          sqlCommand.CommandTimeout = 600;
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          while (sqlDataReader.Read())
            str10 = sqlDataReader[0].ToString();
          str10 = str10.Replace(".", "");
          string str11 = this.Left(str10, 2);
          string str12 = str10.Substring(2, 2);
          str10 = str10.Substring(4, 4) + str12 + str11;
          sqlDataReader.Close();
          connection5.Close();
        }
        if (this.txtLastMdataPoints.Text != "" && this.IsInteger(this.txtLastMdataPoints.Text) && this.chLastM.Checked)
        {
          if (this.rbDays.Checked && this.chLastM.Checked)
          {
            str9 = str9 + " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(day, -" + this.txtLastMdataPoints.Text + ", '" + str10 + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";
            str2 += str9;
          }
          if (this.rbWeeks.Checked && this.chLastM.Checked)
          {
            str9 = str9 + " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(week, -" + this.txtLastMdataPoints.Text + ",'" + str10 + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";
            str2 += str9;
          }
          if (this.rbMonths.Checked && this.chLastM.Checked)
          {
            str9 = str9 + " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(month, -" + this.txtLastMdataPoints.Text + ", '" + str10 + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";
            str2 += str9;
          }
          if (this.rbYears.Checked && this.chLastM.Checked)
          {
            str9 = str9 + " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(year, -" + this.txtLastMdataPoints.Text + ", '" + str10 + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";
            str2 += str9;
          }
        }
        if (this.dtSortDateFrom.Text != "" && this.dtSortDateTo.Text != "" && this.chkSortDate.Checked)
        {
          string str13 = str9 + " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT( DATETIME, ISNULL(  '" + this.dtSortDateFrom.Text + "' , '1900-01-01'), 103 ) and CONVERT( DATETIME, ISNULL( '" + this.dtSortDateTo.Text + "' , '1900-01-01'), 103 )  ";
          str2 += str13;
        }
        string str14;
        if (this.txtLastNdataPoints.Text != "" && this.IsInteger(this.txtLastNdataPoints.Text) && this.chLastN.Checked)
        {
          Form1.intMaxPoints = (float) Convert.ToInt32(this.txtLastNdataPoints.Text);
          string str15 = str2 + " and 0=0 ";
          str14 = "select top " + Form1.intMaxPoints.ToString() + " * from Products where " + str15 + " order by  [TS_ABS] desc";
        }
        else
          str14 = "select * from Products where " + (str2 + Form1.strFilterLaufnr + " and 0=0") + " and ProduktCode = '" + this.cmbProductCode.Text + "' order by CONVERT(DATETIME, ISNULL( [TS_ABS], '1900-01-01'), 103 ) desc";
        this.label1.Text = str14;
        string cmdText3 = str14.Replace("and Virt_Ozid in )", "").Replace("and REFERENCED_CPV in )", "");
        Form1.strSQLfiltered = cmdText3;
        this.txtTestSQL.Text = cmdText3;
        string connectionString = Form1.connectionString;
        using (SqlConnection connection6 = new SqlConnection(Form1.connectionString))
        {
          try
          {
            connection6.Open();
            SqlCommand sqlCommand = new SqlCommand(cmdText3, connection6);
            sqlCommand.CommandTimeout = 600;
            sqlCommand.ExecuteScalar();
            DataTable dataTable = new DataTable();
            if (Form1.strHasID == "0")
            {
              dataTable.Columns.Add("PRODUKTCODE", typeof (string));
              dataTable.Columns.Add("SORT_DATE", typeof (System.DateTime));
              dataTable.Columns.Add("TS_ABS", typeof (System.DateTime));
              dataTable.Columns.Add("LAUFNR", typeof (string));
              dataTable.Columns.Add("CHNR_ENDPRODUKT", typeof (string));
              dataTable.Columns.Add("PROCESS_CODE", typeof (string));
              dataTable.Columns.Add("PROCESS_CODE_NAME", typeof (string));
              dataTable.Columns.Add("PARAMETER_NAME", typeof (string));
              dataTable.Columns.Add("ASSAY", typeof (string));
              dataTable.Columns.Add("VIRT_OZID", typeof (string));
              dataTable.Columns.Add("TREND_WERT", typeof (string));
              dataTable.Columns.Add("TREND_WERT_2", typeof (string));
              dataTable.Columns.Add("ISTWERT_LIMS", typeof (string));
              dataTable.Columns.Add("LCL", typeof (string));
              dataTable.Columns.Add("UCL", typeof (string));
              dataTable.Columns.Add("CL", typeof (string));
              dataTable.Columns.Add("UAL", typeof (string));
              dataTable.Columns.Add("LAL", typeof (string));
              dataTable.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", typeof (string));
              dataTable.Columns.Add("DECIMAL_PLACES_AL", typeof (string));
              dataTable.Columns.Add("DATA_TYPE", typeof (string));
              dataTable.Columns.Add("SOURCE_SYSTEM", typeof (string));
              dataTable.Columns.Add("EXCURSION", typeof (string));
              dataTable.Columns.Add("REFERENCED_CPV", typeof (string));
              dataTable.Columns.Add("IS_IN_RUN_NUMBER_RANGE", typeof (string));
              dataTable.Columns.Add("LOCATION", typeof (string));
            }
            if (Form1.strHasID == "1")
            {
              dataTable.Columns.Add("ID", typeof (string));
              dataTable.Columns.Add("PRODUKTCODE", typeof (string));
              dataTable.Columns.Add("SORT_DATE", typeof (System.DateTime));
              dataTable.Columns.Add("TS_ABS", typeof (System.DateTime));
              dataTable.Columns.Add("LAUFNR", typeof (string));
              dataTable.Columns.Add("CHNR_ENDPRODUKT", typeof (string));
              dataTable.Columns.Add("PROCESS_CODE", typeof (string));
              dataTable.Columns.Add("PROCESS_CODE_NAME", typeof (string));
              dataTable.Columns.Add("PARAMETER_NAME", typeof (string));
              dataTable.Columns.Add("ASSAY", typeof (string));
              dataTable.Columns.Add("VIRT_OZID", typeof (string));
              dataTable.Columns.Add("TREND_WERT", typeof (string));
              dataTable.Columns.Add("TREND_WERT_2", typeof (string));
              dataTable.Columns.Add("ISTWERT_LIMS", typeof (string));
              dataTable.Columns.Add("LCL", typeof (string));
              dataTable.Columns.Add("UCL", typeof (string));
              dataTable.Columns.Add("CL", typeof (string));
              dataTable.Columns.Add("UAL", typeof (string));
              dataTable.Columns.Add("LAL", typeof (string));
              dataTable.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", typeof (string));
              dataTable.Columns.Add("DECIMAL_PLACES_AL", typeof (string));
              dataTable.Columns.Add("DATA_TYPE", typeof (string));
              dataTable.Columns.Add("SOURCE_SYSTEM", typeof (string));
              dataTable.Columns.Add("EXCURSION", typeof (string));
              dataTable.Columns.Add("REFERENCED_CPV", typeof (string));
              dataTable.Columns.Add("IS_IN_RUN_NUMBER_RANGE", typeof (string));
              dataTable.Columns.Add("LOCATION", typeof (string));
            }
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            while (sqlDataReader.Read())
            {
              if (Form1.strHasID == "1")
                dataTable.Rows.Add((object) sqlDataReader["ID"].ToString(), (object) sqlDataReader["PRODUKTCODE"].ToString(), sqlDataReader["SORT_DATE"], sqlDataReader["TS_ABS"], (object) sqlDataReader["LAUFNR"].ToString(), (object) sqlDataReader["CHNR_ENDPRODUKT"].ToString(), (object) sqlDataReader["PROCESS_CODE"].ToString(), (object) sqlDataReader["PROCESS_CODE_NAME"].ToString(), (object) sqlDataReader["PARAMETER_NAME"].ToString(), (object) sqlDataReader["ASSAY"].ToString(), (object) sqlDataReader["VIRT_OZID"].ToString(), (object) sqlDataReader["TREND_WERT"].ToString(), (object) sqlDataReader["TREND_WERT_2"].ToString(), (object) sqlDataReader["ISTWERT_LIMS"].ToString(), (object) sqlDataReader["LCL"].ToString(), (object) sqlDataReader["UCL"].ToString(), (object) sqlDataReader["CL"].ToString(), (object) sqlDataReader["UAL"].ToString(), (object) sqlDataReader["LAL"].ToString(), (object) sqlDataReader["DECIMAL_PLACES_XCL_SUBSTITUTED"].ToString(), (object) sqlDataReader["DECIMAL_PLACES_AL"].ToString(), (object) sqlDataReader["DATA_TYPE"].ToString(), (object) sqlDataReader["SOURCE_SYSTEM"].ToString(), (object) sqlDataReader["EXCURSION"].ToString(), (object) sqlDataReader["REFERENCED_CPV"].ToString(), (object) sqlDataReader["IS_IN_RUN_NUMBER_RANGE"].ToString(), (object) sqlDataReader["LOCATION"].ToString());
              if (Form1.strHasID == "0")
                dataTable.Rows.Add((object) sqlDataReader["PRODUKTCODE"].ToString(), sqlDataReader["SORT_DATE"], sqlDataReader["TS_ABS"], (object) sqlDataReader["LAUFNR"].ToString(), (object) sqlDataReader["CHNR_ENDPRODUKT"].ToString(), (object) sqlDataReader["PROCESS_CODE"].ToString(), (object) sqlDataReader["PROCESS_CODE_NAME"].ToString(), (object) sqlDataReader["PARAMETER_NAME"].ToString(), (object) sqlDataReader["ASSAY"].ToString(), (object) sqlDataReader["VIRT_OZID"].ToString(), (object) sqlDataReader["TREND_WERT"].ToString(), (object) sqlDataReader["TREND_WERT_2"].ToString(), (object) sqlDataReader["ISTWERT_LIMS"].ToString(), (object) sqlDataReader["LCL"].ToString(), (object) sqlDataReader["UCL"].ToString(), (object) sqlDataReader["CL"].ToString(), (object) sqlDataReader["UAL"].ToString(), (object) sqlDataReader["LAL"].ToString(), (object) sqlDataReader["DECIMAL_PLACES_XCL_SUBSTITUTED"].ToString(), (object) sqlDataReader["DECIMAL_PLACES_AL"].ToString(), (object) sqlDataReader["DATA_TYPE"].ToString(), (object) sqlDataReader["SOURCE_SYSTEM"].ToString(), (object) sqlDataReader["EXCURSION"].ToString(), (object) sqlDataReader["REFERENCED_CPV"].ToString(), (object) sqlDataReader["IS_IN_RUN_NUMBER_RANGE"].ToString(), (object) sqlDataReader["LOCATION"].ToString());
            }
            this.dataGridViewRaw.Columns["TS_ABS"].DefaultCellStyle.Format = "dd.MM.yyyy";
            this.dataGridViewRaw.Columns["SORT_DATE"].DefaultCellStyle.Format = "dd.MM.yyyy";
            this.dataGridViewRaw.DataSource = (object) dataTable;
            sqlDataReader.Close();
            connection6.Close();
          }
          catch (Exception ex)
          {
            int num = (int) MessageBox.Show(ex.Message);
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chkSortDate_CheckedChanged(object sender, EventArgs e)
    {
      if (this.chkSortDate.Checked)
      {
        this.chLastN.Checked = false;
        this.chLastN.Enabled = false;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = false;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = false;
        this.rbDays.Enabled = false;
        this.rbWeeks.Enabled = false;
        this.rbMonths.Enabled = false;
        this.rbYears.Enabled = false;
        this.dtSortDateFrom.Enabled = true;
        this.dtSortDateTo.Enabled = true;
        this.cmbLaufnrMin.Enabled = false;
        this.cmbLaufnrMax.Enabled = false;
        this.txtLastNdataPoints.Enabled = false;
        this.txtLastMdataPoints.Enabled = false;
      }
      else
      {
        this.chLastN.Checked = false;
        this.chLastN.Enabled = true;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = true;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = true;
        this.rbDays.Enabled = true;
        this.rbWeeks.Enabled = true;
        this.rbMonths.Enabled = true;
        this.rbYears.Enabled = true;
        this.dtSortDateFrom.Enabled = true;
        this.dtSortDateTo.Enabled = true;
        this.cmbLaufnrMin.Enabled = true;
        this.cmbLaufnrMax.Enabled = true;
        this.txtLastNdataPoints.Enabled = true;
        this.txtLastMdataPoints.Enabled = true;
      }
    }

    private void chkLaufNr_CheckedChanged(object sender, EventArgs e)
    {
      if (this.chkLaufNr.Checked)
      {
        this.chLastN.Checked = false;
        this.chLastN.Enabled = false;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = false;
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = false;
        this.rbDays.Enabled = false;
        this.rbWeeks.Enabled = false;
        this.rbMonths.Enabled = false;
        this.rbYears.Enabled = false;
        this.dtSortDateFrom.Enabled = false;
        this.dtSortDateTo.Enabled = false;
        this.cmbLaufnrMin.Enabled = true;
        this.cmbLaufnrMax.Enabled = true;
        this.txtLastNdataPoints.Enabled = false;
        this.txtLastMdataPoints.Enabled = false;
      }
      else
      {
        this.chLastN.Checked = false;
        this.chLastN.Enabled = true;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = true;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = true;
        this.rbDays.Enabled = true;
        this.rbWeeks.Enabled = true;
        this.rbMonths.Enabled = true;
        this.rbYears.Enabled = true;
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = true;
        this.dtSortDateFrom.Enabled = true;
        this.dtSortDateTo.Enabled = true;
        this.cmbLaufnrMin.Enabled = true;
        this.cmbLaufnrMax.Enabled = true;
        this.txtLastNdataPoints.Enabled = true;
        this.txtLastMdataPoints.Enabled = true;
      }
    }

    private void chLastN_CheckedChanged(object sender, EventArgs e)
    {
      if (this.chLastN.Checked)
      {
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = false;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = false;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = false;
        this.rbDays.Enabled = false;
        this.rbWeeks.Enabled = false;
        this.rbMonths.Enabled = false;
        this.rbYears.Enabled = false;
        this.dtSortDateFrom.Enabled = false;
        this.dtSortDateTo.Enabled = false;
        this.txtLastMdataPoints.Enabled = false;
        this.dtSortDateFrom.Enabled = false;
        this.dtSortDateTo.Enabled = false;
        this.cmbLaufnrMin.Enabled = false;
        this.cmbLaufnrMax.Enabled = false;
        this.txtLastNdataPoints.Enabled = true;
        this.txtLastMdataPoints.Enabled = false;
      }
      else
      {
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = true;
        this.chLastM.Checked = false;
        this.chLastM.Enabled = true;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = true;
        this.rbDays.Enabled = true;
        this.rbWeeks.Enabled = true;
        this.rbMonths.Enabled = true;
        this.rbYears.Enabled = true;
        this.dtSortDateFrom.Enabled = true;
        this.dtSortDateTo.Enabled = true;
        this.cmbLaufnrMin.Enabled = true;
        this.cmbLaufnrMax.Enabled = true;
        this.txtLastNdataPoints.Enabled = true;
        this.txtLastMdataPoints.Enabled = true;
      }
    }

    private void chLastM_CheckedChanged(object sender, EventArgs e)
    {
      if (this.chLastM.Checked)
      {
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = false;
        this.chLastN.Checked = false;
        this.chLastN.Enabled = false;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = false;
        this.rbDays.Enabled = true;
        this.rbWeeks.Enabled = true;
        this.rbMonths.Enabled = true;
        this.rbYears.Enabled = true;
        this.dtSortDateFrom.Enabled = false;
        this.dtSortDateTo.Enabled = false;
        this.cmbLaufnrMin.Enabled = false;
        this.cmbLaufnrMax.Enabled = false;
        this.txtLastNdataPoints.Enabled = false;
        this.txtLastMdataPoints.Enabled = true;
      }
      else
      {
        this.chkSortDate.Checked = false;
        this.chkSortDate.Enabled = true;
        this.chLastN.Checked = false;
        this.chLastN.Enabled = true;
        this.chkLaufNr.Checked = false;
        this.chkLaufNr.Enabled = true;
        this.rbDays.Enabled = false;
        this.rbWeeks.Enabled = false;
        this.rbMonths.Enabled = false;
        this.rbYears.Enabled = false;
        this.dtSortDateFrom.Enabled = true;
        this.dtSortDateTo.Enabled = true;
        this.cmbLaufnrMin.Enabled = true;
        this.cmbLaufnrMax.Enabled = true;
        this.txtLastNdataPoints.Enabled = true;
        this.txtLastMdataPoints.Enabled = true;
      }
    }

    private void chkAllRefCpv_CheckedChanged(object sender, EventArgs e)
    {
      try
      {
        string cmdText = "select distinct REFERENCED_CPV from Products order by REFERENCED_CPV";
        if (this.chkAllRefCpv.Checked)
        {
          this.clbRefCpv.Enabled = false;
          this.clbRefCpv.Items.Clear();
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
            sqlCommand.ExecuteScalar();
            new DataTable().Columns.Add("REFERENCED_CPV", typeof (string));
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            this.chkAllRefCpv.Checked = true;
            while (sqlDataReader.Read())
              this.clbRefCpv.Items.Add((object) sqlDataReader["REFERENCED_CPV"].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
          this.clbRefCpv.Enabled = true;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void chkAllVirtOzid_CheckedChanged(object sender, EventArgs e)
    {
      try
      {
        string cmdText = "select distinct VIRT_OZID from Products order by VIRT_OZID";
        if (this.chkAllVirtOzid.Checked)
        {
          this.clbVirtOzid.Enabled = false;
          this.clbVirtOzid.Items.Clear();
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
            sqlCommand.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            while (sqlDataReader.Read())
            {
              this.clbVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
              this.lstCheckExclVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
            }
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
          this.clbVirtOzid.Enabled = true;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnFilterData_Click(object sender, EventArgs e)
    {
      if (this.dtSortDateFrom.Value > this.dtSortDateTo.Value && this.chkSortDate.Checked)
      {
        int num1 = (int) MessageBox.Show("The first Date should be less or equal to the second Date!");
      }
      else
      {
        try
        {
          Form1.filterOn = true;
          if (this.clbRefCpv.Items.Count == 0)
            this.chkAllRefCpv.Checked = true;
          if (this.clbVirtOzid.Items.Count == 0)
            this.chkAllVirtOzid.Checked = true;
          this.lblDataGridTitle.Text = "Table of raw data entity (Original: NO  -  Filtered: Yes ) ";
          this.btnFilterData_Click_1(sender, e);
          int num2 = (int) MessageBox.Show("Action : <Filter Data> is finished!");
        }
        catch (Exception ex)
        {
          int num3 = (int) MessageBox.Show(ex.Message);
        }
      }
      Form1.LastWindowState = FormWindowState.Minimized;
    }

    private void Form1_Resize(object sender, EventArgs e)
    {
      if (this.WindowState == Form1.LastWindowState)
        return;
      Form1.LastWindowState = this.WindowState;
      if (this.WindowState == FormWindowState.Maximized)
      {
        int num1 = (int) MessageBox.Show("Maximized");
      }
      if (this.WindowState == FormWindowState.Normal)
      {
        int num2 = (int) MessageBox.Show("Restored");
      }
    }

    private void Form1_Load(object sender, EventArgs e)
    {
      CultureInfo.CreateSpecificCulture("en-GB");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chkAllRefCpv, "Enable/Disable RefCpv items, if it is enabled then we use all of them");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chkAllVirtOzid, "Enable/Disable Virt_Ozid items, if it is enabled then we use all of them");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chkSortDate, "Enable/Disable filter Sort Date");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chLastN, "Enable/Disable filter Last N points");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chLastM, "Enable/Disable filter Last M points");
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.chkLaufNr, "Enable/Disable filter LaufNr");
      string[] strArray = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini");
      Form1.connectionString = strArray[0];
      Form1.strRscript = strArray[1];
      Form1.strRpath = strArray[2];
      Form1.strDataDir = strArray[3];
      Form1.strOutputDir = strArray[4];
      this.delProdukts();
      using (SqlConnection connection = new SqlConnection(Form1.connectionString))
      {
        connection.Open();
        new SqlCommand("delete from DataGrid ", connection).ExecuteScalar();
        connection.Close();
      }
      this.btnReset.Enabled = false;
    }

    public void delProdukts()
    {
      using (SqlConnection sqlConnection = new SqlConnection(Form1.connectionString))
      {
        SqlConnection connection = new SqlConnection(Form1.connectionString);
        SqlCommand sqlCommand1 = new SqlCommand("delete From Products", connection);
        connection.Open();
        SqlCommand sqlCommand2 = new SqlCommand("delete From Products", connection);
        sqlCommand2.CommandTimeout = 600;
        sqlCommand2.ExecuteNonQuery();
        sqlConnection.Close();
      }
    }

    private void tabMain_Click(object sender, EventArgs e)
    {
      try
      {
        if (this.tabMain.SelectedTab.Text == "Calculation View")
        {
          this.lblProductCode.Visible = false;
          this.cmbProductCode.Visible = false;
          this.lblRferencedCPV.Visible = false;
          this.clbRefCpv.Visible = false;
          this.lblVirtOzid.Visible = false;
          this.clbVirtOzid.Visible = false;
          this.chkAllVirtOzid.Visible = false;
          this.btnFilterData.Visible = false;
          this.btnReset.Visible = false;
          this.lblSortDate.Visible = false;
          this.lblLaufNRfrom.Visible = false;
          this.lblLastNdataPoints.Visible = false;
          this.lblLastMdataPoints.Visible = false;
          this.dtSortDateFrom.Visible = false;
          this.dtSortDateTo.Visible = false;
          this.chkSortDate.Visible = false;
          this.label2.Visible = false;
          this.cmbLaufnrMin.Visible = false;
          this.cmbLaufnrMax.Visible = false;
          this.chkLaufNr.Visible = false;
          this.lblActiveLN.Visible = false;
          this.txtLastNdataPoints.Visible = false;
          this.chLastN.Visible = false;
          this.rbDays.Visible = false;
          this.rbWeeks.Visible = false;
          this.rbMonths.Visible = false;
          this.rbYears.Visible = false;
          this.btnOpenFile.Visible = false;
          this.button1.Visible = false;
          this.btnCalculationSearch.Visible = false;
          this.btnCalculationView.Visible = false;
          this.btnCalculation.Visible = false;
          this.panel3.Visible = false;
          this.groupFilterSelection.Visible = false;
          this.panelSelection.Visible = false;
          this.panelButtons.Visible = false;
          this.tabMain.SelectTab("ViewCalculation");
          this.frmSave_Load(sender, e);
        }
        if (this.tabMain.SelectedTab.Text == "Compare Calculations")
        {
          this.tabMain.SelectTab("CompareCalculation");
          string sqlQuery1 = "select distinct VIRT_OZID from CalculationRaw order by VIRT_OZID";
          string sqlQuery2 = "select distinct CalcID from CalculationRaw order by CalcID";
          string sqlQuery3 = "select distinct PRODUCTCODE from CalculationRaw order by PRODUCTCODE";
          this.chkVirtOzid3.Items.Clear();
          this.chkVirtOzid4.Items.Clear();
          this.cmbProdID3.Items.Clear();
          this.cmbProdID4.Items.Clear();
          this.cmbCalcID3.Items.Clear();
          this.cmbCalcID4.Items.Clear();
          this.SQLRunFillChechedListBox(sqlQuery1, this.chkVirtOzid3);
          this.SQLRunFillChechedListBox(sqlQuery1, this.chkVirtOzid4);
          this.SQLRunFill(sqlQuery3, this.cmbProdID3);
          this.SQLRunFill(sqlQuery3, this.cmbProdID4);
          this.SQLRunFill(sqlQuery2, this.cmbCalcID3);
          this.SQLRunFill(sqlQuery2, this.cmbCalcID4);
        }
        if (this.tabMain.SelectedTab.Text == "Calculation Search")
        {
          this.panel3.Visible = true;
          this.groupFilterSelection.Visible = true;
          this.panelSelection.Visible = true;
          this.panelButtons.Visible = true;
          this.tabMain.SelectTab("SearchCalculation");
          this.frmHistResults_Load(sender, e);
        }
        if (!(this.tabMain.SelectedTab.Text == "Main"))
          return;
        this.lblProductCode.Visible = true;
        this.cmbProductCode.Visible = true;
        this.lblRferencedCPV.Visible = true;
        this.clbRefCpv.Visible = true;
        this.lblVirtOzid.Visible = true;
        this.clbVirtOzid.Visible = true;
        this.chkAllVirtOzid.Visible = true;
        this.btnFilterData.Visible = true;
        this.btnReset.Visible = true;
        this.lblSortDate.Visible = true;
        this.lblLaufNRfrom.Visible = true;
        this.lblLastNdataPoints.Visible = true;
        this.lblLastMdataPoints.Visible = true;
        this.dtSortDateFrom.Visible = true;
        this.dtSortDateTo.Visible = true;
        this.chkSortDate.Visible = true;
        this.label2.Visible = true;
        this.cmbLaufnrMin.Visible = true;
        this.cmbLaufnrMax.Visible = true;
        this.chkLaufNr.Visible = true;
        this.lblActiveLN.Visible = true;
        this.txtLastNdataPoints.Visible = true;
        this.chLastN.Visible = true;
        this.rbDays.Visible = true;
        this.rbWeeks.Visible = true;
        this.rbMonths.Visible = true;
        this.rbYears.Visible = true;
        this.btnOpenFile.Visible = true;
        this.button1.Visible = true;
        this.btnCalculationSearch.Visible = true;
        this.btnCalculationView.Visible = true;
        this.btnCalculation.Visible = true;
        this.tabMain.SelectTab("Main");
        if (this.dataGridViewRaw.Rows.Count < 2)
          this.Form1_Load(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void btnReset_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.rbYears.Checked = true;
        this.lblDataGridTitle.Text = "Table of raw data entity(Original: Yes - Filtered: NO)";
        this.btnReset_Click(sender, e);
        this.chkSortDate.Enabled = true;
        this.chkLaufNr.Enabled = true;
        this.chLastN.Enabled = true;
        this.chLastM.Enabled = true;
        this.chkSortDate.Checked = false;
        this.chkLaufNr.Checked = false;
        this.chLastN.Checked = false;
        this.chLastM.Checked = false;
        int num = (int) MessageBox.Show("Action : <Reset Filter> is finished!");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void button1_Click(object sender, EventArgs e)
    {
      try
      {
        this.button1_Click_1(sender, e);
        int num = (int) MessageBox.Show("Action : <Export Data to Excel For Calculation> is finished!");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void btnCalculation_Click_1(object sender, EventArgs e)
    {
      try
      {
        Form1.stopCalc = true;
        this.btnCalculation_Click(sender, e);
        int num = (int) MessageBox.Show("The calculation is finished!");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void btnOpenFile_Click(object sender, EventArgs e)
    {
      Form1.filterOn = false;
      this.lblDataGridTitle.Text = "Table of raw data entity (Original: Yes  -  Filtered: NO ) ";
      this.btnCalculationView.Visible = true;
      this.chkSortDate.Enabled = false;
      this.clbRefCpv.Enabled = false;
      this.clbVirtOzid.Enabled = false;
      this.dataGridViewRaw.Visible = true;
      using (SqlConnection connection = new SqlConnection(Form1.connectionString))
      {
        connection.Open();
        SqlCommand sqlCommand = new SqlCommand("ProdUpdate", connection);
        sqlCommand.CommandType = CommandType.StoredProcedure;
        sqlCommand.ExecuteNonQuery();
        connection.Close();
      }
      Form1.arrDataGridView[Form1.intCounterDataGridViews] = new DataGridView();
      Form1.arrDataGridView[Form1.intCounterDataGridViews].Left = this.dataGridViewRaw.Left;
      Form1.arrDataGridView[Form1.intCounterDataGridViews].Top = this.dataGridViewRaw.Top;
      Form1.arrDataGridView[Form1.intCounterDataGridViews].Visible = true;
      Form1.arrDataGridView[Form1.intCounterDataGridViews].Enabled = true;
      try
      {
        string path = Directory.GetCurrentDirectory() + "\\app.ini";
        string[] strArray1 = File.ReadAllLines(path);
        Form1.connectionString = strArray1[0];
        Form1.strRscript = strArray1[1];
        Form1.strRpath = strArray1[2];
        Form1.strDataDir = strArray1[3];
        Form1.strOutputDir = strArray1[4];
        SqlConnection connection1 = new SqlConnection(Form1.connectionString);
        SqlCommand sqlCommand1 = new SqlCommand("delete From Products", connection1);
        connection1.Open();
        new SqlCommand("delete From Products", connection1).CommandTimeout = 600;
        sqlCommand1.CommandTimeout = 600;
        this.rbYears.Checked = true;
        this.dtSortDateFrom.Text = "01.01.2000";
        this.chkExclVirtOzid.Checked = true;
        this.chkSortDate.Enabled = true;
        this.chkSortDate.Checked = false;
        this.clbRefCpv.Items.Clear();
        this.clbVirtOzid.Items.Clear();
        using (new StreamWriter(path, true))
        {
          this.openFileDialog1.Filter = "Excel Files | *.xlsx";
          this.openFileDialog1.ShowDialog();
          Form1.strFileName = this.openFileDialog1.FileName;
          if (this.openFileDialog1.FileName.ToString() != "openFileDialog1")
          {
            this.clbRefCpv.Items.Clear();
            this.clbVirtOzid.Items.Clear();
            this.btnOpenFile.BackColor = Color.LightBlue;
            DataTable dataTable = new DataTable();
            Form1.dt = this.READExcel(this.openFileDialog1.FileName.ToString());
            SqlConnection connection2 = new SqlConnection(Form1.connectionString);
            SqlCommand selectCommand = new SqlCommand("Select * From Products", connection2);
            selectCommand.CommandTimeout = 180;
            connection2.Open();
            new SqlDataAdapter(selectCommand).Fill(Form1.dt);
            Form1.strHasID = !(Form1.dt.Columns[0].ColumnName == "PRODUKTCODE") ? (!(Form1.dt.Columns[0].ColumnName == "ID") && !(Form1.dt.Columns[0].ColumnName == "Column1") ? "2" : "1") : "0";
            this.dataGridViewRaw.DataSource = (object) Form1.dt;
            CultureInfo cultureInfo = new CultureInfo("de-DE");
            this.dataGridViewRaw.Columns["TS_ABS"].DefaultCellStyle.Format = "dd.MM.yyyy";
            this.dataGridViewRaw.Columns["SORT_DATE"].DefaultCellStyle.Format = "dd.MM.yyyy";
            connection2.Close();
            if (Form1.strHasID == "0")
              this.CopyToSQL2(this.dataGridViewRaw);
            if (Form1.strHasID == "1")
              this.CopyToSQL(this.dataGridViewRaw);
          }
          else
            this.btnOpenFile.Focus();
        }
        if (Form1.strHasID == "1")
        {
          string[] strArray2 = new string[32]
          {
            "ID",
            "PRODUKTCODE",
            "SORT_DATE",
            "TS_ABS",
            "LAUFNR",
            "CHNR_ENDPRODUKT",
            "PROCESS_CODE",
            "PROCESS_CODE_NAME",
            "PARAMETER_NAME",
            "ASSAY",
            "VIRT_OZID",
            "TREND_WERT",
            "TREND_WERT_2",
            "ISTWERT_LIMS",
            "LCL",
            "UCL",
            "CL",
            "UAL",
            "LAL",
            "DECIMAL_PLACES_XCL_SUBSTITUTED",
            "DECIMAL_PLACES_AL",
            "DATA_TYPE",
            "SOURCE_SYSTEM",
            "EXCURSION",
            "REFERENCED_CPV",
            "IS_IN_RUN_NUMBER_RANGE",
            "LOCATION",
            "UserID",
            "ModifiedDate",
            "GraphID",
            "CalcID",
            "FilterID"
          };
        }
        if (Form1.strHasID == "0")
        {
          string[] strArray3 = new string[31]
          {
            "PRODUKTCODE",
            "SORT_DATE",
            "TS_ABS",
            "LAUFNR",
            "CHNR_ENDPRODUKT",
            "PROCESS_CODE",
            "PROCESS_CODE_NAME",
            "PARAMETER_NAME",
            "ASSAY",
            "VIRT_OZID",
            "TREND_WERT",
            "TREND_WERT_2",
            "ISTWERT_LIMS",
            "LCL",
            "UCL",
            "CL",
            "UAL",
            "LAL",
            "DECIMAL_PLACES_XCL_SUBSTITUTED",
            "DECIMAL_PLACES_AL",
            "DATA_TYPE",
            "SOURCE_SYSTEM",
            "EXCURSION",
            "REFERENCED_CPV",
            "IS_IN_RUN_NUMBER_RANGE",
            "LOCATION",
            "UserID",
            "ModifiedDate",
            "GraphID",
            "CalcID",
            "FilterID"
          };
        }
        string connectionString = Form1.connectionString;
        string cmdText1 = "select distinct VIRT_OZID from Products order by VIRT_OZID";
        string cmdText2 = "select distinct REFERENCED_CPV from Products order by REFERENCED_CPV";
        string cmdText3 = !(this.cmbProductCode.Text == "") ? "select distinct PRODUKTCODE from Products where PRODUKTCODE!='" + this.cmbProductCode.Text + "'" : "select distinct PRODUKTCODE from Products ";
        using (SqlConnection connection3 = new SqlConnection(connectionString))
        {
          connection3.Open();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText1, connection3);
          sqlCommand2.ExecuteScalar();
          sqlCommand2.CommandTimeout = 600;
          new DataTable().Columns.Add("VIRT_OZID", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand2.ExecuteReader();
          while (sqlDataReader.Read())
          {
            this.clbVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
            this.lstCheckExclVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
          }
          sqlDataReader.Close();
          connection3.Close();
        }
        using (SqlConnection connection4 = new SqlConnection(connectionString))
        {
          connection4.Open();
          SqlCommand sqlCommand3 = new SqlCommand(cmdText2, connection4);
          sqlCommand3.ExecuteScalar();
          sqlCommand3.CommandTimeout = 600;
          new DataTable().Columns.Add("REFERENCED_CPV", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
          this.chkAllRefCpv.Checked = true;
          while (sqlDataReader.Read())
            this.clbRefCpv.Items.Add((object) sqlDataReader["REFERENCED_CPV"].ToString());
          sqlDataReader.Close();
          connection4.Close();
        }
        using (SqlConnection connection5 = new SqlConnection(connectionString))
        {
          connection5.Open();
          SqlCommand sqlCommand4 = new SqlCommand(cmdText3, connection5);
          sqlCommand4.ExecuteScalar();
          sqlCommand4.CommandTimeout = 600;
          new DataTable().Columns.Add("PRODUKTCODE", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand4.ExecuteReader();
          string str = "";
          while (sqlDataReader.Read())
          {
            this.cmbProductCode.Items.Add((object) sqlDataReader["PRODUKTCODE"].ToString());
            str = sqlDataReader["PRODUKTCODE"].ToString();
          }
          this.cmbProductCode.Text = str;
          sqlDataReader.Close();
          connection5.Close();
        }
        using (SqlConnection connection6 = new SqlConnection(connectionString))
        {
          connection6.Open();
          SqlCommand sqlCommand5 = new SqlCommand("SELECT  distinct  top 1000 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select distinct top 5000 TS_ABS FROM [dbo].[Products] order by TS_ABS )", connection6);
          sqlCommand5.CommandTimeout = 600;
          SqlDataReader sqlDataReader1 = sqlCommand5.ExecuteReader();
          while (sqlDataReader1.Read())
            this.cmbLaufnrMin.Items.Add((object) sqlDataReader1["LAUFNR"].ToString());
          SqlCommand sqlCommand6 = new SqlCommand("SELECT  distinct  top 1 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select min(TS_ABS) FROM [dbo].[Products] ) ", connection6);
          sqlDataReader1.Close();
          sqlCommand6.CommandTimeout = 600;
          SqlDataReader sqlDataReader2 = sqlCommand6.ExecuteReader();
          while (sqlDataReader2.Read())
            this.cmbLaufnrMin.Text = sqlDataReader2["LAUFNR"].ToString();
          sqlDataReader2.Close();
          sqlCommand6.CommandTimeout = 600;
          SqlDataReader sqlDataReader3 = new SqlCommand("SELECT  distinct  top 1000 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select distinct top 5000 TS_ABS FROM [dbo].[Products] order by TS_ABS desc)", connection6).ExecuteReader();
          while (sqlDataReader3.Read())
            this.cmbLaufnrMax.Items.Add((object) sqlDataReader3["LAUFNR"].ToString());
          SqlCommand sqlCommand7 = new SqlCommand("SELECT  distinct  top 1 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select max(TS_ABS) FROM [dbo].[Products] )", connection6);
          sqlDataReader3.Close();
          sqlCommand7.CommandTimeout = 600;
          SqlDataReader sqlDataReader4 = sqlCommand7.ExecuteReader();
          while (sqlDataReader4.Read())
            this.cmbLaufnrMax.Text = sqlDataReader4["LAUFNR"].ToString();
          sqlDataReader4.Close();
          connection6.Close();
        }
        using (SqlConnection connection7 = new SqlConnection(connectionString))
        {
          connection7.Open();
          SqlCommand sqlCommand8 = new SqlCommand("SELECT  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS", connection7);
          sqlCommand8.CommandTimeout = 600;
          SqlDataReader sqlDataReader5 = sqlCommand8.ExecuteReader();
          SqlCommand sqlCommand9 = new SqlCommand("SELECT top 1  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS", connection7);
          sqlDataReader5.Close();
          SqlDataReader sqlDataReader6 = sqlCommand9.ExecuteReader();
          while (sqlDataReader6.Read())
          {
            this.dtSortDateFrom.Text = sqlDataReader6["TS_ABS"].ToString();
            Form1.strInitialDate = sqlDataReader6["TS_ABS"].ToString();
          }
          sqlDataReader6.Close();
          SqlDataReader sqlDataReader7 = new SqlCommand("SELECT  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS desc", connection7).ExecuteReader();
          while (sqlDataReader7.Read())
            this.dtSortDateTo.Text = sqlDataReader7["TS_ABS"].ToString();
          SqlCommand sqlCommand10 = new SqlCommand("SELECT top 1  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS desc", connection7);
          sqlDataReader7.Close();
          SqlDataReader sqlDataReader8 = sqlCommand10.ExecuteReader();
          while (sqlDataReader8.Read())
          {
            this.dtSortDateTo.Text = sqlDataReader8["TS_ABS"].ToString();
            Form1.strEndDate = sqlDataReader8["TS_ABS"].ToString();
          }
          sqlDataReader8.Close();
          connection7.Close();
        }
        using (SqlConnection connection8 = new SqlConnection(connectionString))
        {
          connection8.Open();
          SqlCommand sqlCommand11 = new SqlCommand("insert into DataGrid select " + Form1.intCounterDataGridViews.ToString() + ",'" + this.cmbProductCode.Text + "'", connection8);
          sqlCommand11.ExecuteScalar();
          sqlCommand11.CommandTimeout = 600;
          SqlCommand sqlCommand12 = new SqlCommand("ProdUpdate", connection8);
          sqlCommand12.CommandType = CommandType.StoredProcedure;
          sqlCommand12.ExecuteNonQuery();
          connection8.Close();
        }
        if (this.cmbProductCode.Text == "" && this.cmbProductCode.Items.Count < 2)
          this.cmbProductCode.Text = this.cmbProductCode.Items[this.cmbProductCode.Items.Count - 1].ToString();
        int num = (int) MessageBox.Show("Action : <Select Data Raw file> is finished!");
      }
      catch (Exception ex)
      {
        if (ex.ToString().IndexOf("The process cannot access the file") > 0)
        {
          int num1 = (int) MessageBox.Show("Please close the Excel file " + Form1.strFileName + " before reading the data into the application");
        }
        else
        {
          int num2 = (int) MessageBox.Show(ex.ToString());
        }
      }
      if (this.clbVirtOzid.Items.Count > 0)
        this.btnReset.Enabled = true;
      if (this.cmbProductCode.Items.Count > 1)
      {
        for (int index1 = 0; index1 < this.cmbProductCode.Items.Count; ++index1)
        {
          for (int index2 = this.cmbProductCode.Items.Count - 1; index2 >= 0; --index2)
          {
            if ((string) this.cmbProductCode.Items[index2] == (string) this.cmbProductCode.Items[index1])
              this.cmbProductCode.Items.Remove((object) index2);
          }
        }
      }
      this.chkAllVirtOzid.Checked = true;
    }

    private void btnCalculationView_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.tabMain.SelectTab("ViewCalculation");
        this.frmSave_Load(sender, e);
        int num = (int) MessageBox.Show("Action : <Calculation View> is finished!");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnCalculationSearch_Click(object sender, EventArgs e)
    {
      try
      {
        this.tabMain.SelectTab("SearchCalculation");
        this.frmHistResults_Load(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnSave_Click(object sender, EventArgs e)
    {
      try
      {
        Form1.connectionString = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini")[0];
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
        if (!Form1.IsCalcIDAvailable)
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            Form1.NParameterTotal = this.txtNParameterTotal.Text;
            Form1.NStatistically = this.txtNStatistically.Text;
            Form1.PercentStatistically = this.txtPercentStatistically.Text;
            Form1.DoNotFitStatistically = this.txtDoNotFitStatistically.Text;
            Form1.CalcID = this.cmbCalcID.Text;
            Form1.User = this.txtUser.Text;
            Form1.TimePointData = this.txtTimePointData.Text;
            Form1.TimePointCalc = this.txtTimePointCalc.Text;
            Form1.Note = this.txtNote.Text;
            Form1.Active = !this.chkActive.Checked ? 0 : 1;
            SqlCommand sqlCommand3 = new SqlCommand("insert into CalcResultView (ID, NParameterTotal, " + "NStatistically, PercentStatistically, DoNotFitStatistically, CalcID, [User],TimePointData,TimePointCalc, Note,Active) values(" + " '" + Form1.ID.ToString() + "','" + Form1.NParameterTotal + "','" + Form1.NStatistically + "','" + Form1.PercentStatistically + "','" + Form1.DoNotFitStatistically + "','" + Form1.CalcID + "','" + Form1.User + "','" + Form1.TimePointData + "','" + Form1.TimePointCalc + "','" + Form1.Note + "','" + Form1.Active.ToString() + "')", connection);
            sqlCommand3.Parameters.Add("@ID", SqlDbType.Int).Value = (object) Form1.ID;
            sqlCommand3.Parameters.Add("@NParameterTotal", SqlDbType.NVarChar, 50).Value = (object) Form1.NParameterTotal;
            sqlCommand3.Parameters.Add("@NStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.NStatistically;
            sqlCommand3.Parameters.Add("@PercentStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.PercentStatistically;
            sqlCommand3.Parameters.Add("@DoNotFitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.DoNotFitStatistically;
            sqlCommand3.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
            sqlCommand3.Parameters.Add("@User", SqlDbType.NVarChar, 50).Value = (object) Form1.User;
            sqlCommand3.Parameters.Add("@TimePointData", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointData;
            sqlCommand3.Parameters.Add("@TimePointCalc", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointCalc;
            sqlCommand3.Parameters.Add("@Note", SqlDbType.NVarChar, 250).Value = (object) Form1.Note;
            sqlCommand3.Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
            sqlCommand3.ExecuteNonQuery();
            if (short.Parse(this.txtNParameterTotal.Text) <= (short) 0)
              return;
            for (int index = 0; index < (int) short.Parse(this.txtNParameterTotal.Text); ++index)
            {
              Form1.OZID = this.strMatrix[index, 1];
              Form1.CalcID = this.cmbCalcID.Text;
              Form1.TotalN = this.strMatrix[index, 2];
              Form1.KPI0 = this.strMatrix[index, 3];
              Form1.KPI1 = this.strMatrix[index, 4];
              Form1.KPI2 = this.strMatrix[index, 5];
              Form1.KPI3 = this.strMatrix[index, 6];
              Form1.FitStatistically = this.strMatrix[index, 7];
              Form1.RelevantForDiscussion = this.strMatrix[index, 8];
              Form1.Additional_note = this.strMatrix[index, 9];
              string cmdText = "select ID from Graphs where VIRT_OZID ='" + Form1.OZID + "' and CalcID = '" + Form1.CalcID + "'";
              SqlDataReader sqlDataReader = new SqlCommand(cmdText, connection).ExecuteReader();
              SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
              while (sqlDataReader.Read())
                Form1.GraphID = sqlDataReader["ID"].ToString();
              sqlDataReader.Close();
              SqlCommand sqlCommand5 = new SqlCommand("insert into VIRT_OZID_per_calculation (VIRT_OZID,CalcID,TotalN,KPI0,KPI1,KPI2,KPI3,Additional_note,FitStatistically,RelevantForDiscussion,GraphID) " + " values(" + " '" + Form1.OZID + "','" + Form1.CalcID + "','" + Form1.TotalN + "','" + Form1.KPI0 + "','" + Form1.KPI1 + "','" + Form1.KPI2 + "','" + Form1.KPI3 + "','" + Form1.Additional_note + "','" + Form1.FitStatistically + "','" + Form1.RelevantForDiscussion + "','" + Form1.GraphID + "')", connection);
              sqlCommand5.Parameters.Add("@VIRT_OZID", SqlDbType.NVarChar, 250).Value = (object) Form1.OZID;
              sqlCommand5.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
              sqlCommand5.Parameters.Add("@TotalN", SqlDbType.NVarChar, 50).Value = (object) Form1.TotalN;
              sqlCommand5.Parameters.Add("@KPI0", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI0;
              sqlCommand5.Parameters.Add("@KPI1", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI1;
              sqlCommand5.Parameters.Add("@KPI2", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI2;
              sqlCommand5.Parameters.Add("@KPI3", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI3;
              sqlCommand5.Parameters.Add("@Additional_note", SqlDbType.NVarChar, 50).Value = (object) Form1.Additional_note;
              sqlCommand5.Parameters.Add("@FitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.FitStatistically;
              sqlCommand5.Parameters.Add("@RelevantForDiscussion", SqlDbType.NVarChar, 250).Value = (object) Form1.RelevantForDiscussion;
              sqlCommand5.Parameters.Add("@GraphID", SqlDbType.NVarChar, 50).Value = (object) Form1.GraphID;
              sqlCommand5.ExecuteNonQuery();
            }
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            Form1.NParameterTotal = this.txtNParameterTotal.Text;
            Form1.NStatistically = this.txtNStatistically.Text;
            Form1.PercentStatistically = this.txtPercentStatistically.Text;
            Form1.DoNotFitStatistically = this.txtDoNotFitStatistically.Text;
            Form1.CalcID = this.cmbCalcID.Text;
            Form1.User = this.txtUser.Text;
            Form1.TimePointData = this.txtTimePointData.Text;
            Form1.TimePointCalc = this.txtTimePointCalc.Text;
            Form1.Note = this.txtNote.Text;
            Form1.Active = !this.chkActive.Checked ? 0 : 1;
            SqlCommand sqlCommand6 = new SqlCommand("update CalcResultView " + " set ID='" + Form1.ID.ToString() + "'," + " NParameterTotal='" + Form1.NParameterTotal + "'," + " NStatistically='" + Form1.NStatistically + "'," + " PercentStatistically='" + Form1.PercentStatistically + "'," + " DoNotFitStatistically='" + Form1.DoNotFitStatistically + "'," + " CalcID='" + Form1.CalcID + "'," + " [User]='" + Form1.User + "'," + " TimePointData='" + Form1.TimePointData + "'," + " TimePointCalc='" + Form1.TimePointCalc + "'," + " Note='" + Form1.Note + "'," + " Active  ='" + Form1.Active.ToString() + "' where  CalcID='" + Form1.CalcID + "'", connection);
            sqlCommand6.Parameters.Add("@ID", SqlDbType.Int).Value = (object) Form1.ID;
            sqlCommand6.Parameters.Add("@NParameterTotal", SqlDbType.NVarChar, 50).Value = (object) Form1.NParameterTotal;
            sqlCommand6.Parameters.Add("@NStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.NStatistically;
            sqlCommand6.Parameters.Add("@PercentStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.PercentStatistically;
            sqlCommand6.Parameters.Add("@DoNotFitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.DoNotFitStatistically;
            sqlCommand6.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
            sqlCommand6.Parameters.Add("@User", SqlDbType.NVarChar, 50).Value = (object) Form1.User;
            sqlCommand6.Parameters.Add("@TimePointData", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointData;
            sqlCommand6.Parameters.Add("@TimePointCalc", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointCalc;
            sqlCommand6.Parameters.Add("@Note", SqlDbType.NVarChar, 250).Value = (object) Form1.Note;
            sqlCommand6.Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
            sqlCommand6.ExecuteNonQuery();
            if (short.Parse(this.txtNParameterTotal.Text) > (short) 0)
            {
              for (int index = 0; index < (int) short.Parse(this.txtNParameterTotal.Text); ++index)
              {
                if (!this.rbAll.Checked)
                  ;
                Form1.OZID = this.strMatrix[index, 1];
                Form1.CalcID = this.cmbCalcID.Text;
                Form1.TotalN = this.strMatrix[index, 2];
                Form1.KPI0 = this.strMatrix[index, 3];
                Form1.KPI1 = this.strMatrix[index, 4];
                Form1.KPI2 = this.strMatrix[index, 5];
                Form1.KPI3 = this.strMatrix[index, 6];
                Form1.FitStatistically = this.strMatrix[index, 7];
                Form1.RelevantForDiscussion = this.strMatrix[index, 8];
                Form1.Additional_note = this.strMatrix[index, 9];
                string cmdText = "update VIRT_OZID_per_calculation " + " set VIRT_OZID='" + Form1.OZID + "'," + " CalcID='" + this.cmbCalcID.Text + "'," + " TotalN='" + Form1.TotalN + "'," + " KPI0='" + Form1.KPI0 + "'," + " KPI1='" + Form1.KPI1 + "'," + " KPI2='" + Form1.KPI2 + "'," + " KPI3='" + Form1.KPI3 + "'," + " FitStatistically='" + Form1.FitStatistically + "'," + " RelevantForDiscussion='" + Form1.RelevantForDiscussion + "'," + " Additional_note='" + Form1.Additional_note + "' where VIRT_OZID='" + Form1.OZID + "' and CalcID='" + this.cmbCalcID.Text + "'";
                if (this.strMatrix[index, 1] != null)
                {
                  SqlCommand sqlCommand7 = new SqlCommand(cmdText, connection);
                  sqlCommand7.Parameters.Add("@VIRT_OZID", SqlDbType.NVarChar, 250).Value = (object) Form1.OZID;
                  sqlCommand7.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
                  sqlCommand7.Parameters.Add("@TotalN", SqlDbType.NVarChar, 50).Value = (object) Form1.TotalN;
                  sqlCommand7.Parameters.Add("@KPI0", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI0;
                  sqlCommand7.Parameters.Add("@KPI1", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI1;
                  sqlCommand7.Parameters.Add("@KPI2", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI2;
                  sqlCommand7.Parameters.Add("@KPI3", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI3;
                  sqlCommand7.Parameters.Add("@Additional_note", SqlDbType.NVarChar, 50).Value = (object) Form1.Additional_note;
                  sqlCommand7.Parameters.Add("@FitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.FitStatistically;
                  sqlCommand7.Parameters.Add("@RelevantForDiscussion", SqlDbType.NVarChar, 250).Value = (object) Form1.RelevantForDiscussion;
                  sqlCommand7.ExecuteNonQuery();
                }
              }
            }
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void frmSave_Load(object sender, EventArgs e)
    {
      this.dataGridViewRaw.Visible = true;
      string[] strArray = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini");
      Form1.connectionString = strArray[0];
      string str1 = strArray[1];
      string str2 = strArray[2];
      string str3 = strArray[3];
      string str4 = strArray[4];
      Form1.strOutPutPath = strArray[4];
      try
      {
        new ToolTip()
        {
          AutoPopDelay = 5000,
          InitialDelay = 1000,
          ReshowDelay = 500,
          ShowAlways = true
        }.SetToolTip((Control) this.chkActive, "Set the calculation to be in active status");
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct CalcID, CalcDate from CalculationRaw where ProductCode = '" + this.cmbProductCode.Text + "' order by CalcDate";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          new DataTable().Columns.Add("CalcID", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          this.cmbCalcID.Items.Clear();
          while (sqlDataReader.Read())
            this.cmbCalcID.Items.Add((object) sqlDataReader["CalcID"].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbCalcID_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.txtNParameterTotal.Text = "";
        this.txtNStatistically.Text = "";
        this.txtPercentStatistically.Text = "";
        this.txtDoNotFitStatistically.Text = "";
        this.txtUser.Text = "";
        this.txtTimePointData.Text = "";
        this.txtTimePointCalc.Text = "";
        this.txtNote.Text = "";
        this.chkActive.Checked = false;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          this.panel1.Controls.Clear();
          string cmdText1 = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText1, connection);
          sqlCommand1.ExecuteScalar();
          new DataTable().Columns.Add("VIRT_OZID", typeof (string));
          SqlDataReader sqlDataReader1 = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText1, connection);
          int index1 = 0;
          while (sqlDataReader1.Read())
          {
            TextBox textBox1 = new TextBox();
            textBox1.Name = "tb1" + index1.ToString();
            textBox1.Enabled = false;
            textBox1.Text = sqlDataReader1["VIRT_OZID"].ToString();
            textBox1.Top = index1 * 20;
            textBox1.Left = 3;
            TextBox textBox2 = textBox1;
            this.strMatrix[index1, 1] = sqlDataReader1["VIRT_OZID"].ToString();
            Form1.strOZID = textBox2.Text;
            Form1.strArr2[index1] = sqlDataReader1["VIRT_OZID"].ToString();
            TextBox textBox3 = new TextBox();
            textBox3.Enabled = false;
            textBox3.Width = 50;
            textBox3.Text = sqlDataReader1["TotalN"].ToString();
            textBox3.Top = index1 * 20;
            textBox3.Left = 110;
            TextBox textBox4 = textBox3;
            this.strMatrix[index1, 2] = sqlDataReader1["TotalN"].ToString();
            TextBox textBox5 = new TextBox();
            textBox5.Enabled = false;
            textBox5.Width = 50;
            textBox5.Text = sqlDataReader1["KPI0"].ToString();
            textBox5.Top = index1 * 20;
            textBox5.Left = 190;
            TextBox textBox6 = textBox5;
            this.strMatrix[index1, 3] = sqlDataReader1["KPI0"].ToString();
            TextBox textBox7 = new TextBox();
            textBox7.Enabled = false;
            textBox7.Width = 50;
            textBox7.Text = sqlDataReader1["KPI1"].ToString();
            textBox7.Top = index1 * 20;
            textBox7.Left = 270;
            TextBox textBox8 = textBox7;
            this.strMatrix[index1, 4] = sqlDataReader1["KPI1"].ToString();
            TextBox textBox9 = new TextBox();
            textBox9.Enabled = false;
            textBox9.Width = 50;
            textBox9.Text = sqlDataReader1["KPI2"].ToString();
            textBox9.Top = index1 * 20;
            textBox9.Left = 360;
            TextBox textBox10 = textBox9;
            this.strMatrix[index1, 5] = sqlDataReader1["KPI2"].ToString();
            TextBox textBox11 = new TextBox();
            textBox11.Enabled = false;
            textBox11.Width = 50;
            textBox11.Text = sqlDataReader1["KPI3"].ToString();
            textBox11.Top = index1 * 20;
            textBox11.Left = 440;
            TextBox textBox12 = textBox11;
            this.strMatrix[index1, 6] = sqlDataReader1["KPI3"].ToString();
            CheckBox checkBox1 = new CheckBox();
            checkBox1.Name = "ch1" + index1.ToString();
            checkBox1.Text = string.Format("{0}", (object) "yes");
            checkBox1.Top = index1 * 20;
            checkBox1.Left = 580;
            CheckBox checkBox2 = checkBox1;
            checkBox2.Click += new EventHandler(this.chk_Click);
            int num1 = sqlDataReader1["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader1["FitStatistically"].ToString() == "True" ? 1 : 0);
            checkBox2.Checked = num1 != 0;
            string[,] strMatrix1 = this.strMatrix;
            int index2 = index1;
            bool flag = checkBox2.Checked;
            string str1 = flag.ToString();
            strMatrix1[index2, 7] = str1;
            Button button1 = new Button();
            button1.Name = "b1" + index1.ToString();
            button1.Text = string.Format("{0}", (object) "Chart");
            button1.Top = index1 * 20;
            button1.Left = 680;
            Button button2 = button1;
            this.panel1.Controls.Add((Control) button2);
            button2.Click += new EventHandler(this.ba_Click);
            CheckBox checkBox3 = new CheckBox();
            checkBox3.Name = "ch2" + index1.ToString();
            checkBox3.Text = string.Format("{0}", (object) "yes");
            checkBox3.Top = index1 * 20;
            checkBox3.Left = 770;
            CheckBox checkBox4 = checkBox3;
            checkBox4.Click += new EventHandler(this.chk_Click);
            int num2 = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
            checkBox4.Checked = num2 != 0;
            string[,] strMatrix2 = this.strMatrix;
            int index3 = index1;
            flag = checkBox4.Checked;
            string str2 = flag.ToString();
            strMatrix2[index3, 8] = str2;
            TextBox textBox13 = new TextBox();
            textBox13.Name = "tb7" + index1.ToString();
            textBox13.Text = sqlDataReader1["Additional_note"].ToString();
            textBox13.Width = 150;
            textBox13.Top = index1 * 20;
            textBox13.Left = 900;
            TextBox textBox14 = textBox13;
            this.strMatrix[index1, 9] = textBox14.Text;
            textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
            this.panel1.Controls.Add((Control) textBox2);
            this.panel1.Controls.Add((Control) textBox4);
            this.panel1.Controls.Add((Control) textBox6);
            this.panel1.Controls.Add((Control) textBox8);
            this.panel1.Controls.Add((Control) textBox10);
            this.panel1.Controls.Add((Control) textBox12);
            this.panel1.Controls.Add((Control) checkBox2);
            this.panel1.Controls.Add((Control) checkBox4);
            this.panel1.Controls.Add((Control) textBox14);
            this.panel1.Controls.Add(this.ba);
            ++index1;
          }
          sqlDataReader1.Close();
          string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlDataReader sqlDataReader2 = sqlCommand2.ExecuteReader();
          SqlCommand sqlCommand3 = new SqlCommand(cmdText2, connection);
          int num = 0;
          while (sqlDataReader2.Read())
            ++num;
          this.txtNParameterTotal.Text = num.ToString();
          string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
          sqlDataReader2.Close();
          SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
          while (sqlDataReader3.Read())
            this.txtTimePointCalc.Text = (string) sqlDataReader3[0];
          string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal = '0'";
          sqlDataReader3.Close();
          SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
          while (sqlDataReader4.Read())
            this.txtNStatistically.Text = sqlDataReader4[0].ToString();
          string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID2.Text + "')";
          sqlDataReader4.Close();
          SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
          while (sqlDataReader5.Read())
            this.txtPercentStatistically.Text = sqlDataReader5[0].ToString();
          string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal != '0'  ";
          sqlDataReader5.Close();
          SqlDataReader sqlDataReader6 = new SqlCommand(cmdText6, connection).ExecuteReader();
          while (sqlDataReader6.Read())
            this.txtDoNotFitStatistically.Text = sqlDataReader6[0].ToString();
          sqlDataReader6.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        if (ex.Source != null)
        {
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
          sqlCommand4.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand4.ExecuteReader();
          SqlCommand sqlCommand5 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
          {
            this.txtNParameterTotal.Text = sqlDataReader["NParameterTotal"].ToString();
            this.txtNStatistically.Text = sqlDataReader["NStatistically"].ToString();
            this.txtPercentStatistically.Text = sqlDataReader["PercentStatistically"].ToString();
            this.txtDoNotFitStatistically.Text = sqlDataReader["DoNotFitStatistically"].ToString();
            this.cmbCalcID.Text = sqlDataReader["CalcID"].ToString();
            this.txtUser.Text = sqlDataReader["User"].ToString();
            this.txtTimePointData.Text = sqlDataReader["TimePointData"].ToString();
            this.txtTimePointCalc.Text = sqlDataReader["TimePointCalc"].ToString();
            this.txtNote.Text = sqlDataReader["Note"].ToString();
            this.chkActive.Checked = !(sqlDataReader["Active"].ToString() == "0");
          }
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand6 = new SqlCommand(cmdText, connection);
          sqlCommand6.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand6.ExecuteReader();
          SqlCommand sqlCommand7 = new SqlCommand(cmdText, connection);
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chk_MouseMove(object sender, EventArgs e)
    {
      try
      {
        ToolTip toolTip = new ToolTip();
        CheckBox checkBox = (CheckBox) sender;
        if (checkBox.Name.Substring(0, 3) == "ch1")
        {
          toolTip.AutoPopDelay = 5000;
          toolTip.InitialDelay = 1000;
          toolTip.ReshowDelay = 500;
          toolTip.ShowAlways = true;
          toolTip.SetToolTip((Control) checkBox, "This checkbox displays Fitstatistically status");
        }
        if (checkBox.Name.Substring(0, 3) == "ch2")
        {
          toolTip.AutoPopDelay = 5000;
          toolTip.InitialDelay = 1000;
          toolTip.ReshowDelay = 500;
          toolTip.ShowAlways = true;
          toolTip.SetToolTip((Control) checkBox, "This checkbox displays Relevant For Discussion status");
        }
        if (!(checkBox.Name.Substring(0, 3) == "ch3"))
          return;
        toolTip.AutoPopDelay = 5000;
        toolTip.InitialDelay = 1000;
        toolTip.ReshowDelay = 500;
        toolTip.ShowAlways = true;
        toolTip.SetToolTip((Control) checkBox, "This checkbox displays Active status of the calculation");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void chk_Click(object sender, EventArgs e)
    {
      try
      {
        CheckBox checkBox = (CheckBox) sender;
        if (checkBox.Name.Substring(0, 3) == "ch1")
          this.strMatrix[(int) short.Parse(checkBox.Name.Substring(3, checkBox.Name.Length - 3)), 7] = !checkBox.Checked ? "False" : "True";
        if (checkBox.Name.Substring(0, 3) == "ch2")
          this.strMatrix[(int) short.Parse(checkBox.Name.Substring(3, checkBox.Name.Length - 3)), 8] = !checkBox.Checked ? "False" : "True";
        if (!(checkBox.Name.Substring(0, 3) == "ch3"))
          return;
        this.strMatrix[(int) short.Parse(checkBox.Name.Substring(3, checkBox.Name.Length - 3)), 9] = !checkBox.Checked ? "False" : "True";
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void tb_LostFocus(object sender, EventArgs e)
    {
      TextBox textBox = (TextBox) sender;
      this.strMatrix[(int) short.Parse(textBox.Name.Substring(3, textBox.Name.Length - 3)), 9] = textBox.Text;
    }

    public void ba34_Click2(object sender, EventArgs e)
    {
      this.pictureBox3.Image = (Image) null;
      this.pictureBox4.Image = (Image) null;
      this.pictureBox3.Visible = false;
      this.pictureBox4.Visible = false;
      Form1.blnPressed = false;
      try
      {
        new SqlConnection(Form1.connectionString).Open();
        Button button = (Button) sender;
        int num1 = (int) short.Parse(button.Name.Substring(1, button.Name.Length - 1));
        if (num1 < 20)
          Form1.strOZID = Form1.strArr2[num1 - 10];
        if (num1 < 200 && num1 > 19)
          Form1.strOZID = Form1.strArr2[num1 - 100];
        try
        {
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID3.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count > 0)
          {
            byte[] buffer = (byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"];
            MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
            this.pictureBox3.Visible = true;
            this.pictureBox3.Image = Image.FromStream((Stream) memoryStream, true);
          }
        }
        catch (Exception ex)
        {
          int num2 = (int) MessageBox.Show(ex.ToString());
        }
        try
        {
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID4.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count <= 0)
            return;
          byte[] buffer = (byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"];
          MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
          this.pictureBox4.Visible = true;
          this.pictureBox4.Image = Image.FromStream((Stream) memoryStream, true);
        }
        catch (Exception ex)
        {
          int num3 = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public void ba34_Click(object sender, EventArgs e)
    {
      this.pictureBox3.Image = (Image) null;
      this.pictureBox4.Image = (Image) null;
      this.pictureBox3.Visible = false;
      this.pictureBox4.Visible = false;
      Form1.blnPressed = false;
      try
      {
        new SqlConnection(Form1.connectionString).Open();
        Button button = (Button) sender;
        int num1 = (int) short.Parse(button.Name.Substring(1, button.Name.Length - 1));
        if (num1 < 20)
          Form1.strOZID = Form1.strArr2[num1 - 10];
        if (num1 < 200 && num1 > 19)
          Form1.strOZID = Form1.strArr2[num1 - 100];
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection1 = new SqlConnection(Form1.connectionString);
          connection1.Open();
          SqlDataAdapter sqlDataAdapter1 = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID3.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection1));
          DataSet dataSet1 = new DataSet();
          sqlDataAdapter1.Fill(dataSet1, "Graphs");
          int count1 = dataSet1.Tables["Graphs"].Rows.Count;
          if (count1 > 0)
          {
            byte[] buffer = (byte[]) dataSet1.Tables["Graphs"].Rows[count1 - 1]["ImageValue"];
            MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
            this.btnZoomOut3.Visible = true;
            this.btnPrint3.Visible = true;
            this.pictureBox3.Visible = true;
            this.pictureBox3.Image = Image.FromStream((Stream) memoryStream, true);
          }
          this.picGraph.Visible = true;
          SqlConnection connection2 = new SqlConnection(Form1.connectionString);
          connection2.Open();
          SqlDataAdapter sqlDataAdapter2 = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID4.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection2));
          DataSet dataSet2 = new DataSet();
          sqlDataAdapter2.Fill(dataSet2, "Graphs");
          int count2 = dataSet2.Tables["Graphs"].Rows.Count;
          if (count2 > 0)
          {
            byte[] buffer = (byte[]) dataSet2.Tables["Graphs"].Rows[count2 - 1]["ImageValue"];
            MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
            this.btnZoomOut4.Visible = true;
            this.btnPrint4.Visible = true;
            this.pictureBox4.Visible = true;
            this.pictureBox4.Image = Image.FromStream((Stream) memoryStream, true);
          }
          this.ba34_Click2(sender, e);
        }
        catch (Exception ex)
        {
          int num2 = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    public void ba1_Click(object sender, EventArgs e)
    {
      try
      {
        this.btnZoomIn1.Visible = true;
        this.pictureBox2.Visible = true;
        this.btnPrint.Visible = true;
        this.label82.Visible = true;
        this.label83.Visible = true;
        new SqlConnection(Form1.connectionString).Open();
        Form1.blnPressed = false;
        Button button = (Button) sender;
        int num1 = (int) short.Parse(button.Name.Substring(1, button.Name.Length - 1));
        if (num1 < 20)
          Form1.strOZID = Form1.strArr2[num1 - 10];
        if (num1 < 200 && num1 > 19)
          Form1.strOZID = Form1.strArr2[num1 - 100];
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID2.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count <= 0)
            return;
          this.pictureBox2.Image = Image.FromStream((Stream) new MemoryStream((byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"]));
          this.btnPrint.Visible = true;
        }
        catch (Exception ex)
        {
          int num2 = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void ba_Zoom(object sender, EventArgs e)
    {
      try
      {
        SqlConnection connection = new SqlConnection(Form1.connectionString);
        connection.Open();
        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
        DataSet dataSet = new DataSet();
        sqlDataAdapter.Fill(dataSet, "Graphs");
        int count = dataSet.Tables["Graphs"].Rows.Count;
        if (count <= 0)
          return;
        byte[] buffer = (byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"];
        MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
        PictureBox pictureBox = new PictureBox();
        pictureBox.Image = Image.FromStream((Stream) memoryStream, true);
        pictureBox.Location = new Point(3, 3);
        pictureBox.Size = new Size(1100, 900);
        pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
        pictureBox.Refresh();
        Form form = new Form();
        form.Size = new Size(1200, 1000);
        form.Controls.Add((Control) pictureBox);
        int num = (int) form.ShowDialog();
        connection.Close();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void ba_Click(object sender, EventArgs e)
    {
      this.btnZoom1.Visible = true;
      this.picGraph.Image = (Image) null;
      this.picGraph.Visible = false;
      this.picGraph.Image = (Image) null;
      try
      {
        new SqlConnection(Form1.connectionString).Open();
        Button button = (Button) sender;
        int num1 = (int) short.Parse(button.Name.Substring(1, button.Name.Length - 1));
        if (num1 < 20)
          Form1.strOZID = Form1.strArr2[num1 - 10];
        if (num1 < 200 && num1 > 19)
          Form1.strOZID = Form1.strArr2[num1 - 100];
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection1 = new SqlConnection(Form1.connectionString);
          connection1.Open();
          SqlDataAdapter sqlDataAdapter1 = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection1));
          DataSet dataSet1 = new DataSet();
          sqlDataAdapter1.Fill(dataSet1, "Graphs");
          int count1 = dataSet1.Tables["Graphs"].Rows.Count;
          this.picGraph.Visible = true;
          SqlConnection connection2 = new SqlConnection(Form1.connectionString);
          connection2.Open();
          SqlDataAdapter sqlDataAdapter2 = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection2));
          DataSet dataSet2 = new DataSet();
          sqlDataAdapter2.Fill(dataSet2, "Graphs");
          int count2 = dataSet2.Tables["Graphs"].Rows.Count;
          if (count2 <= 0)
            return;
          byte[] buffer = (byte[]) dataSet2.Tables["Graphs"].Rows[count2 - 1]["ImageValue"];
          this.picGraph.Image = Image.FromStream((Stream) new MemoryStream(buffer, 0, buffer.Length), true);
          this.btnPrint2.Visible = true;
        }
        catch (Exception ex)
        {
          int num2 = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void rbFitStat_CheckedChanged(object sender, EventArgs e)
    {
      try
      {
        this.strMatrix = new string[142, 11];
        if (!this.rbNotFitStat.Checked && !this.rbAll.Checked)
          this.rbFitStat.Checked = true;
        this.picGraph.Visible = false;
        Form1.IsCalcIDAvailable = false;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (!Form1.IsCalcIDAvailable)
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') and CalcID = '" + this.cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
            sqlCommand3.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader1 = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
            int index = 0;
            while (sqlDataReader1.Read())
            {
              ++index;
              TextBox textBox1 = new TextBox();
              textBox1.Name = "tb1" + index.ToString();
              textBox1.Enabled = false;
              textBox1.Text = sqlDataReader1[1].ToString();
              textBox1.Top = index * 20;
              textBox1.Left = 3;
              TextBox textBox2 = textBox1;
              this.strMatrix[index, 1] = sqlDataReader1[1].ToString();
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index] = sqlDataReader1[1].ToString();
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 50;
              textBox3.Text = sqlDataReader1[2].ToString();
              textBox3.Top = index * 20;
              textBox3.Left = 110;
              TextBox textBox4 = textBox3;
              this.strMatrix[index, 2] = sqlDataReader1[2].ToString();
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 50;
              textBox5.Text = sqlDataReader1[3].ToString();
              textBox5.Top = index * 20;
              textBox5.Left = 190;
              TextBox textBox6 = textBox5;
              this.strMatrix[index, 3] = sqlDataReader1[3].ToString();
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = sqlDataReader1[4].ToString();
              textBox7.Top = index * 20;
              textBox7.Left = 270;
              TextBox textBox8 = textBox7;
              this.strMatrix[index, 4] = sqlDataReader1[4].ToString();
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = sqlDataReader1[5].ToString();
              textBox9.Top = index * 20;
              textBox9.Left = 360;
              TextBox textBox10 = textBox9;
              this.strMatrix[index, 5] = sqlDataReader1[5].ToString();
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = sqlDataReader1[6].ToString();
              textBox11.Top = index * 20;
              textBox11.Left = 440;
              TextBox textBox12 = textBox11;
              this.strMatrix[index, 6] = sqlDataReader1[6].ToString();
              CheckBox checkBox1 = new CheckBox();
              checkBox1.Name = "ch1" + index.ToString();
              checkBox1.Text = string.Format("{0}", (object) "yes");
              checkBox1.Top = index * 20;
              checkBox1.Left = 580;
              CheckBox checkBox2 = checkBox1;
              checkBox2.Click += new EventHandler(this.chk_Click);
              if (sqlDataReader1[7] == (object) "true" || sqlDataReader1[7] == (object) "True")
                checkBox2.Checked = true;
              this.strMatrix[index, 7] = checkBox2.Checked.ToString();
              Button button1 = new Button();
              button1.Name = "b1" + index.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index * 20;
              button1.Left = 680;
              Button button2 = button1;
              this.panel1.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox3 = new CheckBox();
              checkBox3.Name = "ch2" + index.ToString();
              checkBox3.Text = string.Format("{0}", (object) "yes");
              checkBox3.Top = index * 20;
              checkBox3.Left = 770;
              CheckBox checkBox4 = checkBox3;
              checkBox4.Click += new EventHandler(this.chk_Click);
              int num = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox4.Checked = num != 0;
              this.strMatrix[index, 8] = checkBox4.Checked.ToString();
              TextBox textBox13 = new TextBox();
              textBox13.Name = "tb7" + index.ToString();
              textBox13.Width = 150;
              textBox13.Top = index * 20;
              textBox13.Left = 900;
              TextBox textBox14 = textBox13;
              this.strMatrix[index, 9] = textBox14.Text;
              textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox2);
              this.panel1.Controls.Add((Control) textBox4);
              this.panel1.Controls.Add((Control) textBox6);
              this.panel1.Controls.Add((Control) textBox8);
              this.panel1.Controls.Add((Control) textBox10);
              this.panel1.Controls.Add((Control) textBox12);
              this.panel1.Controls.Add((Control) checkBox2);
              this.panel1.Controls.Add((Control) checkBox4);
              this.panel1.Controls.Add((Control) textBox14);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader1.Close();
            string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader2 = sqlCommand4.ExecuteReader();
            SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
            int num1 = 0;
            while (sqlDataReader2.Read())
              ++num1;
            this.txtNParameterTotal.Text = num1.ToString();
            string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader2.Close();
            SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
            while (sqlDataReader3.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader3[0];
            string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader3.Close();
            SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
            while (sqlDataReader4.Read())
              this.txtNStatistically.Text = sqlDataReader4[0].ToString();
            string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader4.Close();
            SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
            while (sqlDataReader5.Read())
              this.txtPercentStatistically.Text = sqlDataReader5[0].ToString();
            string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader5.Close();
            SqlDataReader sqlDataReader6 = new SqlCommand(cmdText6, connection).ExecuteReader();
            while (sqlDataReader6.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader6[0].ToString();
            sqlDataReader6.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source == null)
            return;
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      else
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText7 = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "' and FitStatistically = 'true'";
            SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
            sqlCommand6.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader7 = sqlCommand6.ExecuteReader();
            SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
            int index = 0;
            while (sqlDataReader7.Read())
            {
              ++index;
              TextBox textBox15 = new TextBox();
              textBox15.Name = "tb1" + index.ToString();
              textBox15.Enabled = false;
              textBox15.Text = sqlDataReader7["VIRT_OZID"].ToString();
              textBox15.Top = index * 20;
              textBox15.Left = 3;
              TextBox textBox16 = textBox15;
              this.strMatrix[index, 1] = sqlDataReader7["VIRT_OZID"].ToString();
              Form1.strOZID = textBox16.Text;
              Form1.strArr2[index] = sqlDataReader7["VIRT_OZID"].ToString();
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = false;
              textBox17.Width = 50;
              textBox17.Text = sqlDataReader7["TotalN"].ToString();
              textBox17.Top = index * 20;
              textBox17.Left = 110;
              TextBox textBox18 = textBox17;
              this.strMatrix[index, 2] = sqlDataReader7["TotalN"].ToString();
              TextBox textBox19 = new TextBox();
              textBox19.Enabled = false;
              textBox19.Width = 50;
              textBox19.Text = sqlDataReader7["KPI0"].ToString();
              textBox19.Top = index * 20;
              textBox19.Left = 190;
              TextBox textBox20 = textBox19;
              this.strMatrix[index, 3] = sqlDataReader7["KPI0"].ToString();
              TextBox textBox21 = new TextBox();
              textBox21.Enabled = false;
              textBox21.Width = 50;
              textBox21.Text = sqlDataReader7["KPI1"].ToString();
              textBox21.Top = index * 20;
              textBox21.Left = 270;
              TextBox textBox22 = textBox21;
              this.strMatrix[index, 4] = sqlDataReader7["KPI1"].ToString();
              TextBox textBox23 = new TextBox();
              textBox23.Enabled = false;
              textBox23.Width = 50;
              textBox23.Text = sqlDataReader7["KPI2"].ToString();
              textBox23.Top = index * 20;
              textBox23.Left = 360;
              TextBox textBox24 = textBox23;
              this.strMatrix[index, 5] = sqlDataReader7["KPI2"].ToString();
              TextBox textBox25 = new TextBox();
              textBox25.Enabled = false;
              textBox25.Width = 50;
              textBox25.Text = sqlDataReader7["KPI3"].ToString();
              textBox25.Top = index * 20;
              textBox25.Left = 440;
              TextBox textBox26 = textBox25;
              this.strMatrix[index, 6] = sqlDataReader7["KPI3"].ToString();
              CheckBox checkBox5 = new CheckBox();
              checkBox5.Name = "ch1" + index.ToString();
              checkBox5.Text = string.Format("{0}", (object) "yes");
              checkBox5.Top = index * 20;
              checkBox5.Left = 580;
              CheckBox checkBox6 = checkBox5;
              checkBox6.Click += new EventHandler(this.chk_Click);
              int num2 = sqlDataReader7["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader7["FitStatistically"].ToString() == "True" ? 1 : 0);
              checkBox6.Checked = num2 != 0;
              this.strMatrix[index, 7] = checkBox6.Checked.ToString();
              Button button3 = new Button();
              button3.Name = "b1" + index.ToString();
              button3.Text = string.Format("{0}", (object) "Chart");
              button3.Top = index * 20;
              button3.Left = 680;
              Button button4 = button3;
              this.panel1.Controls.Add((Control) button4);
              button4.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox7 = new CheckBox();
              checkBox7.Name = "ch2" + index.ToString();
              checkBox7.Text = string.Format("{0}", (object) "yes");
              checkBox7.Top = index * 20;
              checkBox7.Left = 770;
              CheckBox checkBox8 = checkBox7;
              checkBox8.Click += new EventHandler(this.chk_Click);
              int num3 = sqlDataReader7["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader7["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox8.Checked = num3 != 0;
              this.strMatrix[index, 8] = checkBox8.Checked.ToString();
              TextBox textBox27 = new TextBox();
              textBox27.Name = "tb7" + index.ToString();
              textBox27.Text = sqlDataReader7["Additional_note"].ToString();
              textBox27.Width = 150;
              textBox27.Top = index * 20;
              textBox27.Left = 900;
              TextBox textBox28 = textBox27;
              this.strMatrix[index, 9] = textBox28.Text;
              textBox28.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox16);
              this.panel1.Controls.Add((Control) textBox18);
              this.panel1.Controls.Add((Control) textBox20);
              this.panel1.Controls.Add((Control) textBox22);
              this.panel1.Controls.Add((Control) textBox24);
              this.panel1.Controls.Add((Control) textBox26);
              this.panel1.Controls.Add((Control) checkBox6);
              this.panel1.Controls.Add((Control) checkBox8);
              this.panel1.Controls.Add((Control) textBox28);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader7.Close();
            string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader8 = sqlCommand7.ExecuteReader();
            SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
            int num = 0;
            while (sqlDataReader8.Read())
              ++num;
            this.txtNParameterTotal.Text = num.ToString();
            string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader8.Close();
            SqlDataReader sqlDataReader9 = new SqlCommand(cmdText9, connection).ExecuteReader();
            while (sqlDataReader9.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader9[0];
            string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader9.Close();
            SqlDataReader sqlDataReader10 = new SqlCommand(cmdText10, connection).ExecuteReader();
            while (sqlDataReader10.Read())
              this.txtNStatistically.Text = sqlDataReader10[0].ToString();
            string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader10.Close();
            SqlDataReader sqlDataReader11 = new SqlCommand(cmdText11, connection).ExecuteReader();
            while (sqlDataReader11.Read())
              this.txtPercentStatistically.Text = sqlDataReader11[0].ToString();
            string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader11.Close();
            SqlDataReader sqlDataReader12 = new SqlCommand(cmdText12, connection).ExecuteReader();
            while (sqlDataReader12.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader12[0].ToString();
            sqlDataReader12.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source != null)
          {
            int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
          }
        }
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand9 = new SqlCommand(cmdText, connection);
            sqlCommand9.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand9.ExecuteReader();
            SqlCommand sqlCommand10 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
            {
              this.txtNParameterTotal.Text = sqlDataReader["NParameterTotal"].ToString();
              this.txtNStatistically.Text = sqlDataReader["NStatistically"].ToString();
              this.txtPercentStatistically.Text = sqlDataReader["PercentStatistically"].ToString();
              this.txtDoNotFitStatistically.Text = sqlDataReader["DoNotFitStatistically"].ToString();
              this.cmbCalcID.Text = sqlDataReader["CalcID"].ToString();
              this.txtUser.Text = sqlDataReader["User"].ToString();
              this.txtTimePointData.Text = sqlDataReader["TimePointData"].ToString();
              this.txtTimePointCalc.Text = sqlDataReader["TimePointCalc"].ToString();
              this.txtNote.Text = sqlDataReader["Note"].ToString();
              this.chkActive.Checked = !(sqlDataReader["Active"].ToString() == "0");
            }
            sqlDataReader.Close();
            connection.Close();
          }
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand11 = new SqlCommand(cmdText, connection);
            sqlCommand11.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand11.ExecuteReader();
            SqlCommand sqlCommand12 = new SqlCommand(cmdText, connection);
            sqlDataReader.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.Message);
        }
      }
    }

    private void rbNotFitStat_CheckedChanged(object sender, EventArgs e)
    {
      try
      {
        this.strMatrix = new string[142, 11];
        if (!this.rbAll.Checked && !this.rbFitStat.Checked)
          this.rbNotFitStat.Checked = true;
        this.picGraph.Visible = false;
        Form1.IsCalcIDAvailable = false;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (!Form1.IsCalcIDAvailable)
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') and CalcID = '" + this.cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
            sqlCommand3.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader1 = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
            int index = 0;
            while (sqlDataReader1.Read())
            {
              ++index;
              TextBox textBox1 = new TextBox();
              textBox1.Name = "tb1" + index.ToString();
              textBox1.Enabled = false;
              textBox1.Text = sqlDataReader1[1].ToString();
              textBox1.Top = index * 20;
              textBox1.Left = 3;
              TextBox textBox2 = textBox1;
              this.strMatrix[index, 1] = sqlDataReader1[1].ToString();
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index] = sqlDataReader1[1].ToString();
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 50;
              textBox3.Text = sqlDataReader1[2].ToString();
              textBox3.Top = index * 20;
              textBox3.Left = 110;
              TextBox textBox4 = textBox3;
              this.strMatrix[index, 2] = sqlDataReader1[2].ToString();
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 50;
              textBox5.Text = sqlDataReader1[3].ToString();
              textBox5.Top = index * 20;
              textBox5.Left = 190;
              TextBox textBox6 = textBox5;
              this.strMatrix[index, 3] = sqlDataReader1[3].ToString();
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = sqlDataReader1[4].ToString();
              textBox7.Top = index * 20;
              textBox7.Left = 270;
              TextBox textBox8 = textBox7;
              this.strMatrix[index, 4] = sqlDataReader1[4].ToString();
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = sqlDataReader1[5].ToString();
              textBox9.Top = index * 20;
              textBox9.Left = 360;
              TextBox textBox10 = textBox9;
              this.strMatrix[index, 5] = sqlDataReader1[5].ToString();
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = sqlDataReader1[6].ToString();
              textBox11.Top = index * 20;
              textBox11.Left = 440;
              TextBox textBox12 = textBox11;
              this.strMatrix[index, 6] = sqlDataReader1[6].ToString();
              CheckBox checkBox1 = new CheckBox();
              checkBox1.Name = "ch1" + index.ToString();
              checkBox1.Text = string.Format("{0}", (object) "yes");
              checkBox1.Top = index * 20;
              checkBox1.Left = 580;
              CheckBox checkBox2 = checkBox1;
              checkBox2.Click += new EventHandler(this.chk_Click);
              this.strMatrix[index, 7] = checkBox2.Checked.ToString();
              Button button1 = new Button();
              button1.Name = "b1" + index.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index * 20;
              button1.Left = 680;
              Button button2 = button1;
              this.panel1.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox3 = new CheckBox();
              checkBox3.Name = "ch2" + index.ToString();
              checkBox3.Text = string.Format("{0}", (object) "yes");
              checkBox3.Top = index * 20;
              checkBox3.Left = 770;
              CheckBox checkBox4 = checkBox3;
              checkBox4.Click += new EventHandler(this.chk_Click);
              this.strMatrix[index, 8] = checkBox4.Checked.ToString();
              TextBox textBox13 = new TextBox();
              textBox13.Name = "tb7" + index.ToString();
              textBox13.Width = 150;
              textBox13.Top = index * 20;
              textBox13.Left = 900;
              TextBox textBox14 = textBox13;
              int num = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox4.Checked = num != 0;
              this.strMatrix[index, 9] = textBox14.Text;
              textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox2);
              this.panel1.Controls.Add((Control) textBox4);
              this.panel1.Controls.Add((Control) textBox6);
              this.panel1.Controls.Add((Control) textBox8);
              this.panel1.Controls.Add((Control) textBox10);
              this.panel1.Controls.Add((Control) textBox12);
              this.panel1.Controls.Add((Control) checkBox2);
              this.panel1.Controls.Add((Control) checkBox4);
              this.panel1.Controls.Add((Control) textBox14);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader1.Close();
            string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader2 = sqlCommand4.ExecuteReader();
            SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
            int num1 = 0;
            while (sqlDataReader2.Read())
              ++num1;
            this.txtNParameterTotal.Text = num1.ToString();
            string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader2.Close();
            SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
            while (sqlDataReader3.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader3[0];
            string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader3.Close();
            SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
            while (sqlDataReader4.Read())
              this.txtNStatistically.Text = sqlDataReader4[0].ToString();
            string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader4.Close();
            SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
            while (sqlDataReader5.Read())
              this.txtPercentStatistically.Text = sqlDataReader5[0].ToString();
            string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader5.Close();
            SqlDataReader sqlDataReader6 = new SqlCommand(cmdText6, connection).ExecuteReader();
            while (sqlDataReader6.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader6[0].ToString();
            sqlDataReader6.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source == null)
            return;
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      else
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText7 = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "' and FitStatistically = 'false'";
            SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
            sqlCommand6.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader7 = sqlCommand6.ExecuteReader();
            SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
            int index = 0;
            while (sqlDataReader7.Read())
            {
              ++index;
              TextBox textBox15 = new TextBox();
              textBox15.Name = "tb1" + index.ToString();
              textBox15.Enabled = false;
              textBox15.Text = sqlDataReader7["VIRT_OZID"].ToString();
              textBox15.Top = index * 20;
              textBox15.Left = 3;
              TextBox textBox16 = textBox15;
              this.strMatrix[index, 1] = sqlDataReader7["VIRT_OZID"].ToString();
              Form1.strOZID = textBox16.Text;
              Form1.strArr2[index] = sqlDataReader7["VIRT_OZID"].ToString();
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = false;
              textBox17.Width = 50;
              textBox17.Text = sqlDataReader7["TotalN"].ToString();
              textBox17.Top = index * 20;
              textBox17.Left = 110;
              TextBox textBox18 = textBox17;
              this.strMatrix[index, 2] = sqlDataReader7["TotalN"].ToString();
              TextBox textBox19 = new TextBox();
              textBox19.Enabled = false;
              textBox19.Width = 50;
              textBox19.Text = sqlDataReader7["KPI0"].ToString();
              textBox19.Top = index * 20;
              textBox19.Left = 190;
              TextBox textBox20 = textBox19;
              this.strMatrix[index, 3] = sqlDataReader7["KPI0"].ToString();
              TextBox textBox21 = new TextBox();
              textBox21.Enabled = false;
              textBox21.Width = 50;
              textBox21.Text = sqlDataReader7["KPI1"].ToString();
              textBox21.Top = index * 20;
              textBox21.Left = 270;
              TextBox textBox22 = textBox21;
              this.strMatrix[index, 4] = sqlDataReader7["KPI1"].ToString();
              TextBox textBox23 = new TextBox();
              textBox23.Enabled = false;
              textBox23.Width = 50;
              textBox23.Text = sqlDataReader7["KPI2"].ToString();
              textBox23.Top = index * 20;
              textBox23.Left = 360;
              TextBox textBox24 = textBox23;
              this.strMatrix[index, 5] = sqlDataReader7["KPI2"].ToString();
              TextBox textBox25 = new TextBox();
              textBox25.Enabled = false;
              textBox25.Width = 50;
              textBox25.Text = sqlDataReader7["KPI3"].ToString();
              textBox25.Top = index * 20;
              textBox25.Left = 440;
              TextBox textBox26 = textBox25;
              this.strMatrix[index, 6] = sqlDataReader7["KPI3"].ToString();
              CheckBox checkBox5 = new CheckBox();
              checkBox5.Name = "ch1" + index.ToString();
              checkBox5.Text = string.Format("{0}", (object) "yes");
              checkBox5.Top = index * 20;
              checkBox5.Left = 580;
              CheckBox checkBox6 = checkBox5;
              checkBox6.Click += new EventHandler(this.chk_Click);
              checkBox6.Checked = sqlDataReader7["FitStatistically"].ToString() == "true";
              this.strMatrix[index, 7] = checkBox6.Checked.ToString();
              Button button3 = new Button();
              button3.Name = "b1" + index.ToString();
              button3.Text = string.Format("{0}", (object) "Chart");
              button3.Top = index * 20;
              button3.Left = 680;
              Button button4 = button3;
              this.panel1.Controls.Add((Control) button4);
              button4.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox7 = new CheckBox();
              checkBox7.Name = "ch2" + index.ToString();
              checkBox7.Text = string.Format("{0}", (object) "yes");
              checkBox7.Top = index * 20;
              checkBox7.Left = 770;
              CheckBox checkBox8 = checkBox7;
              checkBox8.Click += new EventHandler(this.chk_Click);
              int num = sqlDataReader7["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader7["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox8.Checked = num != 0;
              this.strMatrix[index, 8] = checkBox8.Checked.ToString();
              TextBox textBox27 = new TextBox();
              textBox27.Name = "tb7" + index.ToString();
              textBox27.Text = sqlDataReader7["Additional_note"].ToString();
              textBox27.Width = 150;
              textBox27.Top = index * 20;
              textBox27.Left = 900;
              TextBox textBox28 = textBox27;
              this.strMatrix[index, 9] = textBox28.Text;
              textBox28.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox16);
              this.panel1.Controls.Add((Control) textBox18);
              this.panel1.Controls.Add((Control) textBox20);
              this.panel1.Controls.Add((Control) textBox22);
              this.panel1.Controls.Add((Control) textBox24);
              this.panel1.Controls.Add((Control) textBox26);
              this.panel1.Controls.Add((Control) checkBox6);
              this.panel1.Controls.Add((Control) checkBox8);
              this.panel1.Controls.Add((Control) textBox28);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader7.Close();
            string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader8 = sqlCommand7.ExecuteReader();
            SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
            int num2 = 0;
            while (sqlDataReader8.Read())
              ++num2;
            this.txtNParameterTotal.Text = num2.ToString();
            string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader8.Close();
            SqlDataReader sqlDataReader9 = new SqlCommand(cmdText9, connection).ExecuteReader();
            while (sqlDataReader9.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader9[0];
            string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader9.Close();
            SqlDataReader sqlDataReader10 = new SqlCommand(cmdText10, connection).ExecuteReader();
            while (sqlDataReader10.Read())
              this.txtNStatistically.Text = sqlDataReader10[0].ToString();
            string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader10.Close();
            SqlDataReader sqlDataReader11 = new SqlCommand(cmdText11, connection).ExecuteReader();
            while (sqlDataReader11.Read())
              this.txtPercentStatistically.Text = sqlDataReader11[0].ToString();
            string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader11.Close();
            SqlDataReader sqlDataReader12 = new SqlCommand(cmdText12, connection).ExecuteReader();
            while (sqlDataReader12.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader12[0].ToString();
            sqlDataReader12.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source != null)
          {
            int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
          }
        }
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand9 = new SqlCommand(cmdText, connection);
            sqlCommand9.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand9.ExecuteReader();
            SqlCommand sqlCommand10 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
            {
              this.txtNParameterTotal.Text = sqlDataReader["NParameterTotal"].ToString();
              this.txtNStatistically.Text = sqlDataReader["NStatistically"].ToString();
              this.txtPercentStatistically.Text = sqlDataReader["PercentStatistically"].ToString();
              this.txtDoNotFitStatistically.Text = sqlDataReader["DoNotFitStatistically"].ToString();
              this.cmbCalcID.Text = sqlDataReader["CalcID"].ToString();
              this.txtUser.Text = sqlDataReader["User"].ToString();
              this.txtTimePointData.Text = sqlDataReader["TimePointData"].ToString();
              this.txtTimePointCalc.Text = sqlDataReader["TimePointCalc"].ToString();
              this.txtNote.Text = sqlDataReader["Note"].ToString();
              this.chkActive.Checked = !(sqlDataReader["Active"].ToString() == "0");
            }
            sqlDataReader.Close();
            connection.Close();
          }
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand11 = new SqlCommand(cmdText, connection);
            sqlCommand11.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand11.ExecuteReader();
            SqlCommand sqlCommand12 = new SqlCommand(cmdText, connection);
            sqlDataReader.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.Message);
        }
      }
    }

    private void rbAll_CheckedChanged(object sender, EventArgs e)
    {
      try
      {
        this.strMatrix = new string[142, 11];
        if (!this.rbNotFitStat.Checked && !this.rbFitStat.Checked)
          this.rbAll.Checked = true;
        this.picGraph.Visible = false;
        Form1.IsCalcIDAvailable = false;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (!Form1.IsCalcIDAvailable)
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') and CalcID = '" + this.cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
            sqlCommand3.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader1 = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
            int index = 0;
            while (sqlDataReader1.Read())
            {
              ++index;
              TextBox textBox1 = new TextBox();
              textBox1.Name = "tb1" + index.ToString();
              textBox1.Enabled = false;
              textBox1.Text = sqlDataReader1[1].ToString();
              textBox1.Top = index * 20;
              textBox1.Left = 3;
              TextBox textBox2 = textBox1;
              this.strMatrix[index, 1] = sqlDataReader1[1].ToString();
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index] = sqlDataReader1[1].ToString();
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 50;
              textBox3.Text = sqlDataReader1[2].ToString();
              textBox3.Top = index * 20;
              textBox3.Left = 110;
              TextBox textBox4 = textBox3;
              this.strMatrix[index, 2] = sqlDataReader1[2].ToString();
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 50;
              textBox5.Text = sqlDataReader1[3].ToString();
              textBox5.Top = index * 20;
              textBox5.Left = 190;
              TextBox textBox6 = textBox5;
              this.strMatrix[index, 3] = sqlDataReader1[3].ToString();
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = sqlDataReader1[4].ToString();
              textBox7.Top = index * 20;
              textBox7.Left = 270;
              TextBox textBox8 = textBox7;
              this.strMatrix[index, 4] = sqlDataReader1[4].ToString();
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = sqlDataReader1[5].ToString();
              textBox9.Top = index * 20;
              textBox9.Left = 360;
              TextBox textBox10 = textBox9;
              this.strMatrix[index, 5] = sqlDataReader1[5].ToString();
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = sqlDataReader1[6].ToString();
              textBox11.Top = index * 20;
              textBox11.Left = 440;
              TextBox textBox12 = textBox11;
              this.strMatrix[index, 6] = sqlDataReader1[6].ToString();
              CheckBox checkBox1 = new CheckBox();
              checkBox1.Name = "ch1" + index.ToString();
              checkBox1.Text = string.Format("{0}", (object) "yes");
              checkBox1.Top = index * 20;
              checkBox1.Left = 580;
              CheckBox checkBox2 = checkBox1;
              checkBox2.Click += new EventHandler(this.chk_Click);
              this.strMatrix[index, 7] = checkBox2.Checked.ToString();
              Button button1 = new Button();
              button1.Name = "b1" + index.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index * 20;
              button1.Left = 680;
              Button button2 = button1;
              this.panel1.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox3 = new CheckBox();
              checkBox3.Name = "ch2" + index.ToString();
              checkBox3.Text = string.Format("{0}", (object) "yes");
              checkBox3.Top = index * 20;
              checkBox3.Left = 770;
              CheckBox checkBox4 = checkBox3;
              checkBox4.Click += new EventHandler(this.chk_Click);
              this.strMatrix[index, 8] = checkBox4.Checked.ToString();
              TextBox textBox13 = new TextBox();
              textBox13.Name = "tb7" + index.ToString();
              textBox13.Width = 150;
              textBox13.Top = index * 20;
              textBox13.Left = 900;
              TextBox textBox14 = textBox13;
              int num = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox4.Checked = num != 0;
              this.strMatrix[index, 9] = textBox14.Text;
              textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox2);
              this.panel1.Controls.Add((Control) textBox4);
              this.panel1.Controls.Add((Control) textBox6);
              this.panel1.Controls.Add((Control) textBox8);
              this.panel1.Controls.Add((Control) textBox10);
              this.panel1.Controls.Add((Control) textBox12);
              this.panel1.Controls.Add((Control) checkBox2);
              this.panel1.Controls.Add((Control) checkBox4);
              this.panel1.Controls.Add((Control) textBox14);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader1.Close();
            string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader2 = sqlCommand4.ExecuteReader();
            SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
            int num1 = 0;
            while (sqlDataReader2.Read())
              ++num1;
            this.txtNParameterTotal.Text = num1.ToString();
            string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader2.Close();
            SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
            while (sqlDataReader3.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader3[0];
            string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader3.Close();
            SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
            while (sqlDataReader4.Read())
              this.txtNStatistically.Text = sqlDataReader4[0].ToString();
            string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader4.Close();
            SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
            while (sqlDataReader5.Read())
              this.txtPercentStatistically.Text = sqlDataReader5[0].ToString();
            string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader5.Close();
            SqlDataReader sqlDataReader6 = new SqlCommand(cmdText6, connection).ExecuteReader();
            while (sqlDataReader6.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader6[0].ToString();
            sqlDataReader6.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source == null)
            return;
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      else
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText7 = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "' ";
            SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
            sqlCommand6.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader7 = sqlCommand6.ExecuteReader();
            SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
            int index = 0;
            while (sqlDataReader7.Read())
            {
              ++index;
              TextBox textBox15 = new TextBox();
              textBox15.Name = "tb1" + index.ToString();
              textBox15.Enabled = false;
              textBox15.Text = sqlDataReader7["VIRT_OZID"].ToString();
              textBox15.Top = index * 20;
              textBox15.Left = 3;
              TextBox textBox16 = textBox15;
              this.strMatrix[index, 1] = sqlDataReader7["VIRT_OZID"].ToString();
              Form1.strOZID = textBox16.Text;
              Form1.strArr2[index] = sqlDataReader7["VIRT_OZID"].ToString();
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = false;
              textBox17.Width = 50;
              textBox17.Text = sqlDataReader7["TotalN"].ToString();
              textBox17.Top = index * 20;
              textBox17.Left = 110;
              TextBox textBox18 = textBox17;
              this.strMatrix[index, 2] = sqlDataReader7["TotalN"].ToString();
              TextBox textBox19 = new TextBox();
              textBox19.Enabled = false;
              textBox19.Width = 50;
              textBox19.Text = sqlDataReader7["KPI0"].ToString();
              textBox19.Top = index * 20;
              textBox19.Left = 190;
              TextBox textBox20 = textBox19;
              this.strMatrix[index, 3] = sqlDataReader7["KPI0"].ToString();
              TextBox textBox21 = new TextBox();
              textBox21.Enabled = false;
              textBox21.Width = 50;
              textBox21.Text = sqlDataReader7["KPI1"].ToString();
              textBox21.Top = index * 20;
              textBox21.Left = 270;
              TextBox textBox22 = textBox21;
              this.strMatrix[index, 4] = sqlDataReader7["KPI1"].ToString();
              TextBox textBox23 = new TextBox();
              textBox23.Enabled = false;
              textBox23.Width = 50;
              textBox23.Text = sqlDataReader7["KPI2"].ToString();
              textBox23.Top = index * 20;
              textBox23.Left = 360;
              TextBox textBox24 = textBox23;
              this.strMatrix[index, 5] = sqlDataReader7["KPI2"].ToString();
              TextBox textBox25 = new TextBox();
              textBox25.Enabled = false;
              textBox25.Width = 50;
              textBox25.Text = sqlDataReader7["KPI3"].ToString();
              textBox25.Top = index * 20;
              textBox25.Left = 440;
              TextBox textBox26 = textBox25;
              this.strMatrix[index, 6] = sqlDataReader7["KPI3"].ToString();
              CheckBox checkBox5 = new CheckBox();
              checkBox5.Name = "ch1" + index.ToString();
              checkBox5.Text = string.Format("{0}", (object) "yes");
              checkBox5.Top = index * 20;
              checkBox5.Left = 580;
              CheckBox checkBox6 = checkBox5;
              checkBox6.Click += new EventHandler(this.chk_Click);
              int num2 = sqlDataReader7["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader7["FitStatistically"].ToString() == "True" ? 1 : 0);
              checkBox6.Checked = num2 != 0;
              this.strMatrix[index, 7] = checkBox6.Checked.ToString();
              Button button3 = new Button();
              button3.Name = "b1" + index.ToString();
              button3.Text = string.Format("{0}", (object) "Chart");
              button3.Top = index * 20;
              button3.Left = 680;
              Button button4 = button3;
              this.panel1.Controls.Add((Control) button4);
              button4.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox7 = new CheckBox();
              checkBox7.Name = "ch2" + index.ToString();
              checkBox7.Text = string.Format("{0}", (object) "yes");
              checkBox7.Top = index * 20;
              checkBox7.Left = 770;
              CheckBox checkBox8 = checkBox7;
              checkBox8.Click += new EventHandler(this.chk_Click);
              int num3 = sqlDataReader7["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader7["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox8.Checked = num3 != 0;
              this.strMatrix[index, 8] = checkBox8.Checked.ToString();
              TextBox textBox27 = new TextBox();
              textBox27.Name = "tb7" + index.ToString();
              textBox27.Text = sqlDataReader7["Additional_note"].ToString();
              textBox27.Width = 150;
              textBox27.Top = index * 20;
              textBox27.Left = 900;
              TextBox textBox28 = textBox27;
              this.strMatrix[index, 9] = textBox28.Text;
              textBox28.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox16);
              this.panel1.Controls.Add((Control) textBox18);
              this.panel1.Controls.Add((Control) textBox20);
              this.panel1.Controls.Add((Control) textBox22);
              this.panel1.Controls.Add((Control) textBox24);
              this.panel1.Controls.Add((Control) textBox26);
              this.panel1.Controls.Add((Control) checkBox6);
              this.panel1.Controls.Add((Control) checkBox8);
              this.panel1.Controls.Add((Control) textBox28);
              this.panel1.Controls.Add(this.ba);
            }
            sqlDataReader7.Close();
            string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader8 = sqlCommand7.ExecuteReader();
            SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
            int num = 0;
            while (sqlDataReader8.Read())
              ++num;
            this.txtNParameterTotal.Text = num.ToString();
            string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader8.Close();
            SqlDataReader sqlDataReader9 = new SqlCommand(cmdText9, connection).ExecuteReader();
            while (sqlDataReader9.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader9[0];
            string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader9.Close();
            SqlDataReader sqlDataReader10 = new SqlCommand(cmdText10, connection).ExecuteReader();
            while (sqlDataReader10.Read())
              this.txtNStatistically.Text = sqlDataReader10[0].ToString();
            string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader10.Close();
            SqlDataReader sqlDataReader11 = new SqlCommand(cmdText11, connection).ExecuteReader();
            while (sqlDataReader11.Read())
              this.txtPercentStatistically.Text = sqlDataReader11[0].ToString();
            string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader11.Close();
            SqlDataReader sqlDataReader12 = new SqlCommand(cmdText12, connection).ExecuteReader();
            while (sqlDataReader12.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader12[0].ToString();
            sqlDataReader12.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source != null)
          {
            int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
          }
        }
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand9 = new SqlCommand(cmdText, connection);
            sqlCommand9.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand9.ExecuteReader();
            SqlCommand sqlCommand10 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
            {
              this.txtNParameterTotal.Text = sqlDataReader["NParameterTotal"].ToString();
              this.txtNStatistically.Text = sqlDataReader["NStatistically"].ToString();
              this.txtPercentStatistically.Text = sqlDataReader["PercentStatistically"].ToString();
              this.txtDoNotFitStatistically.Text = sqlDataReader["DoNotFitStatistically"].ToString();
              this.cmbCalcID.Text = sqlDataReader["CalcID"].ToString();
              this.txtUser.Text = sqlDataReader["User"].ToString();
              this.txtTimePointData.Text = sqlDataReader["TimePointData"].ToString();
              this.txtTimePointCalc.Text = sqlDataReader["TimePointCalc"].ToString();
              this.txtNote.Text = sqlDataReader["Note"].ToString();
              this.chkActive.Checked = !(sqlDataReader["Active"].ToString() == "0");
            }
            sqlDataReader.Close();
            connection.Close();
          }
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand11 = new SqlCommand(cmdText, connection);
            sqlCommand11.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand11.ExecuteReader();
            SqlCommand sqlCommand12 = new SqlCommand(cmdText, connection);
            sqlDataReader.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.Message);
        }
      }
    }

    private void cmbCalcID_SelectedIndexChanged_1(object sender, EventArgs e)
    {
      new ToolTip()
      {
        AutoPopDelay = 5000,
        InitialDelay = 1000,
        ReshowDelay = 500,
        ShowAlways = true
      }.SetToolTip((Control) this.label8, "The calculation of KPI is initial and does not change");
      this.btnZoom1.Visible = false;
      this.btnPrint2.Visible = false;
      try
      {
        this.rbAll.Checked = true;
        this.picGraph.Visible = false;
        Form1.IsCalcIDAvailable = false;
        this.rbAll.Enabled = true;
        this.rbFitStat.Enabled = true;
        this.rbNotFitStat.Enabled = true;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (!Form1.IsCalcIDAvailable)
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            SqlDataReader sqlDataReader1 = new SqlCommand("select * FROM[dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "'", connection).ExecuteReader();
            int num1 = 0;
            while (sqlDataReader1.Read())
              ++num1;
            sqlDataReader1.Close();
            string cmdText1;
            if (num1 == 0)
              cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') and CalcID = '" + this.cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
            else
              cmdText1 = "select * FROM[dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader2 = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
            int index1 = 0;
            while (sqlDataReader2.Read())
            {
              TextBox textBox1 = new TextBox();
              textBox1.Name = "tb1" + index1.ToString();
              textBox1.Enabled = false;
              textBox1.Text = sqlDataReader2[1].ToString();
              textBox1.Top = index1 * 20;
              textBox1.Left = 3;
              textBox1.BackColor = Color.White;
              TextBox textBox2 = textBox1;
              this.strMatrix[index1, 1] = sqlDataReader2[1].ToString();
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index1] = sqlDataReader2[1].ToString();
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 50;
              textBox3.Text = sqlDataReader2[2].ToString();
              textBox3.Top = index1 * 20;
              textBox3.Left = 110;
              textBox3.BackColor = Color.White;
              TextBox textBox4 = textBox3;
              this.strMatrix[index1, 2] = sqlDataReader2[2].ToString();
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 50;
              textBox5.Text = sqlDataReader2[3].ToString();
              textBox5.Top = index1 * 20;
              textBox5.Left = 190;
              textBox5.BackColor = Color.White;
              TextBox textBox6 = textBox5;
              this.strMatrix[index1, 3] = sqlDataReader2[3].ToString();
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = sqlDataReader2[4].ToString();
              textBox7.Top = index1 * 20;
              textBox7.Left = 270;
              textBox7.BackColor = Color.White;
              TextBox textBox8 = textBox7;
              this.strMatrix[index1, 4] = sqlDataReader2[4].ToString();
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = sqlDataReader2[5].ToString();
              textBox9.Top = index1 * 20;
              textBox9.Left = 360;
              textBox9.BackColor = Color.White;
              TextBox textBox10 = textBox9;
              this.strMatrix[index1, 5] = sqlDataReader2[5].ToString();
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = sqlDataReader2[6].ToString();
              textBox11.Top = index1 * 20;
              textBox11.Left = 440;
              textBox11.BackColor = Color.White;
              TextBox textBox12 = textBox11;
              this.strMatrix[index1, 6] = sqlDataReader2[6].ToString();
              CheckBox checkBox1 = new CheckBox();
              checkBox1.Name = "ch1" + index1.ToString();
              checkBox1.Text = string.Format("{0}", (object) "yes");
              checkBox1.Top = index1 * 20;
              checkBox1.Left = 580;
              CheckBox checkBox2 = checkBox1;
              checkBox2.Click += new EventHandler(this.chk_Click);
              string[,] strMatrix1 = this.strMatrix;
              int index2 = index1;
              bool flag = checkBox2.Checked;
              string str1 = flag.ToString();
              strMatrix1[index2, 7] = str1;
              Button button1 = new Button();
              button1.Name = "b1" + index1.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index1 * 20;
              button1.Left = 680;
              Button button2 = button1;
              int num2 = sqlDataReader2["FitStatistically"].ToString() == "true" || sqlDataReader2["FitStatistically"].ToString() == "True" ? 1 : (this.strMatrix[index1, 3].ToString() == "100" ? 1 : 0);
              checkBox2.Checked = num2 != 0;
              this.panel1.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba_Click);
              checkBox2.MouseLeave += new EventHandler(this.chk_MouseMove);
              CheckBox checkBox3 = new CheckBox();
              checkBox3.Name = "ch2" + index1.ToString();
              checkBox3.Text = string.Format("{0}", (object) "yes");
              checkBox3.Top = index1 * 20;
              checkBox3.Left = 770;
              CheckBox checkBox4 = checkBox3;
              checkBox4.Click += new EventHandler(this.chk_Click);
              checkBox4.MouseLeave += new EventHandler(this.chk_MouseMove);
              string[,] strMatrix2 = this.strMatrix;
              int index3 = index1;
              flag = checkBox4.Checked;
              string str2 = flag.ToString();
              strMatrix2[index3, 8] = str2;
              TextBox textBox13 = new TextBox();
              textBox13.Name = "tb7" + index1.ToString();
              textBox13.Width = 150;
              textBox13.Top = index1 * 20;
              textBox13.Left = 900;
              TextBox textBox14 = textBox13;
              int num3 = sqlDataReader2["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader2["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox4.Checked = num3 != 0;
              this.strMatrix[index1, 9] = textBox14.Text;
              textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox2);
              this.panel1.Controls.Add((Control) textBox4);
              this.panel1.Controls.Add((Control) textBox6);
              this.panel1.Controls.Add((Control) textBox8);
              this.panel1.Controls.Add((Control) textBox10);
              this.panel1.Controls.Add((Control) textBox12);
              this.panel1.Controls.Add((Control) checkBox2);
              this.panel1.Controls.Add((Control) checkBox4);
              this.panel1.Controls.Add((Control) textBox14);
              this.panel1.Controls.Add(this.ba);
              ++index1;
            }
            sqlDataReader2.Close();
            string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader3 = sqlCommand4.ExecuteReader();
            SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
            int num4 = 0;
            while (sqlDataReader3.Read())
              ++num4;
            this.txtNParameterTotal.Text = num4.ToString();
            string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader3.Close();
            SqlDataReader sqlDataReader4 = new SqlCommand(cmdText3, connection).ExecuteReader();
            while (sqlDataReader4.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader4[0];
            string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader4.Close();
            SqlDataReader sqlDataReader5 = new SqlCommand(cmdText4, connection).ExecuteReader();
            while (sqlDataReader5.Read())
              this.txtNStatistically.Text = sqlDataReader5[0].ToString();
            string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader5.Close();
            SqlDataReader sqlDataReader6 = new SqlCommand(cmdText5, connection).ExecuteReader();
            while (sqlDataReader6.Read())
              this.txtPercentStatistically.Text = sqlDataReader6[0].ToString();
            string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader6.Close();
            SqlDataReader sqlDataReader7 = new SqlCommand(cmdText6, connection).ExecuteReader();
            while (sqlDataReader7.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader7[0].ToString();
            sqlDataReader7.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source == null)
            return;
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      else
      {
        try
        {
          this.txtNParameterTotal.Text = "";
          this.txtNStatistically.Text = "";
          this.txtPercentStatistically.Text = "";
          this.txtDoNotFitStatistically.Text = "";
          this.txtUser.Text = "";
          this.txtTimePointData.Text = "";
          this.txtTimePointCalc.Text = "";
          this.txtNote.Text = "";
          this.chkActive.Checked = false;
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel1.Controls.Clear();
            string cmdText7 = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
            sqlCommand6.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader8 = sqlCommand6.ExecuteReader();
            SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
            int index = 0;
            while (sqlDataReader8.Read())
            {
              TextBox textBox15 = new TextBox();
              textBox15.Name = "tb1" + index.ToString();
              textBox15.Enabled = false;
              textBox15.Text = sqlDataReader8["VIRT_OZID"].ToString();
              textBox15.Top = index * 20;
              textBox15.Left = 3;
              textBox15.BackColor = Color.White;
              TextBox textBox16 = textBox15;
              this.strMatrix[index, 1] = sqlDataReader8["VIRT_OZID"].ToString();
              Form1.strOZID = textBox16.Text;
              Form1.strArr2[index] = sqlDataReader8["VIRT_OZID"].ToString();
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = false;
              textBox17.Width = 50;
              textBox17.Text = sqlDataReader8["TotalN"].ToString();
              textBox17.Top = index * 20;
              textBox17.Left = 110;
              textBox17.BackColor = Color.White;
              TextBox textBox18 = textBox17;
              this.strMatrix[index, 2] = sqlDataReader8["TotalN"].ToString();
              TextBox textBox19 = new TextBox();
              textBox19.Enabled = false;
              textBox19.Width = 50;
              textBox19.Text = sqlDataReader8["KPI0"].ToString();
              textBox19.Top = index * 20;
              textBox19.Left = 190;
              textBox19.BackColor = Color.White;
              TextBox textBox20 = textBox19;
              this.strMatrix[index, 3] = sqlDataReader8["KPI0"].ToString();
              TextBox textBox21 = new TextBox();
              textBox21.Enabled = false;
              textBox21.Width = 50;
              textBox21.Text = sqlDataReader8["KPI1"].ToString();
              textBox21.Top = index * 20;
              textBox21.Left = 270;
              textBox21.BackColor = Color.White;
              TextBox textBox22 = textBox21;
              this.strMatrix[index, 4] = sqlDataReader8["KPI1"].ToString();
              TextBox textBox23 = new TextBox();
              textBox23.Enabled = false;
              textBox23.Width = 50;
              textBox23.Text = sqlDataReader8["KPI2"].ToString();
              textBox23.Top = index * 20;
              textBox23.Left = 360;
              textBox23.BackColor = Color.White;
              TextBox textBox24 = textBox23;
              this.strMatrix[index, 5] = sqlDataReader8["KPI2"].ToString();
              TextBox textBox25 = new TextBox();
              textBox25.Enabled = false;
              textBox25.Width = 50;
              textBox25.Text = sqlDataReader8["KPI3"].ToString();
              textBox25.Top = index * 20;
              textBox25.Left = 440;
              textBox25.BackColor = Color.White;
              TextBox textBox26 = textBox25;
              this.strMatrix[index, 6] = sqlDataReader8["KPI3"].ToString();
              CheckBox checkBox5 = new CheckBox();
              checkBox5.Name = "ch1" + index.ToString();
              checkBox5.Text = string.Format("{0}", (object) "yes");
              checkBox5.Top = index * 20;
              checkBox5.Left = 580;
              CheckBox checkBox6 = checkBox5;
              checkBox6.Click += new EventHandler(this.chk_Click);
              checkBox6.MouseLeave += new EventHandler(this.chk_MouseMove);
              int num5 = sqlDataReader8["FitStatistically"].ToString() == "true" || sqlDataReader8["FitStatistically"].ToString() == "True" ? 1 : (this.strMatrix[index, 3].ToString() == "100" ? 1 : 0);
              checkBox6.Checked = num5 != 0;
              this.strMatrix[index, 7] = checkBox6.Checked.ToString();
              Button button3 = new Button();
              button3.Name = "b1" + index.ToString();
              button3.Text = string.Format("{0}", (object) "Chart");
              button3.Top = index * 20;
              button3.Left = 680;
              Button button4 = button3;
              this.panel1.Controls.Add((Control) button4);
              button4.Click += new EventHandler(this.ba_Click);
              CheckBox checkBox7 = new CheckBox();
              checkBox7.Name = "ch2" + index.ToString();
              checkBox7.Text = string.Format("{0}", (object) "yes");
              checkBox7.Top = index * 20;
              checkBox7.Left = 770;
              CheckBox checkBox8 = checkBox7;
              checkBox8.Click += new EventHandler(this.chk_Click);
              checkBox8.MouseLeave += new EventHandler(this.chk_MouseMove);
              new ToolTip()
              {
                AutoPopDelay = 5000,
                InitialDelay = 1000,
                ReshowDelay = 500,
                ShowAlways = true
              }.SetToolTip((Control) checkBox8, "This checkbox displays Relevant For Discussion status");
              int num6 = sqlDataReader8["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader8["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox8.Checked = num6 != 0;
              this.strMatrix[index, 8] = checkBox8.Checked.ToString();
              TextBox textBox27 = new TextBox();
              textBox27.Name = "tb7" + index.ToString();
              textBox27.Text = sqlDataReader8["Additional_note"].ToString();
              textBox27.Width = 150;
              textBox27.Top = index * 20;
              textBox27.Left = 900;
              TextBox textBox28 = textBox27;
              this.strMatrix[index, 9] = textBox28.Text;
              textBox28.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel1.Controls.Add((Control) textBox16);
              this.panel1.Controls.Add((Control) textBox18);
              this.panel1.Controls.Add((Control) textBox20);
              this.panel1.Controls.Add((Control) textBox22);
              this.panel1.Controls.Add((Control) textBox24);
              this.panel1.Controls.Add((Control) textBox26);
              this.panel1.Controls.Add((Control) checkBox6);
              this.panel1.Controls.Add((Control) checkBox8);
              this.panel1.Controls.Add((Control) textBox28);
              this.panel1.Controls.Add(this.ba);
              ++index;
            }
            sqlDataReader8.Close();
            string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlDataReader sqlDataReader9 = sqlCommand7.ExecuteReader();
            SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
            int num = 0;
            while (sqlDataReader9.Read())
              ++num;
            this.txtNParameterTotal.Text = num.ToString();
            string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "'";
            sqlDataReader9.Close();
            SqlDataReader sqlDataReader10 = new SqlCommand(cmdText9, connection).ExecuteReader();
            while (sqlDataReader10.Read())
              this.txtTimePointCalc.Text = (string) sqlDataReader10[0];
            string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal = '0'";
            sqlDataReader10.Close();
            SqlDataReader sqlDataReader11 = new SqlCommand(cmdText10, connection).ExecuteReader();
            while (sqlDataReader11.Read())
              this.txtNStatistically.Text = sqlDataReader11[0].ToString();
            string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID.Text + "')";
            sqlDataReader11.Close();
            SqlDataReader sqlDataReader12 = new SqlCommand(cmdText11, connection).ExecuteReader();
            while (sqlDataReader12.Read())
              this.txtPercentStatistically.Text = sqlDataReader12[0].ToString();
            string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID.Text + "' and signal != '0'  ";
            sqlDataReader12.Close();
            SqlDataReader sqlDataReader13 = new SqlCommand(cmdText12, connection).ExecuteReader();
            while (sqlDataReader13.Read())
              this.txtDoNotFitStatistically.Text = sqlDataReader13[0].ToString();
            sqlDataReader13.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source != null)
          {
            int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
          }
        }
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand9 = new SqlCommand(cmdText, connection);
            sqlCommand9.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand9.ExecuteReader();
            SqlCommand sqlCommand10 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
            {
              this.txtNParameterTotal.Text = sqlDataReader["NParameterTotal"].ToString();
              this.txtNStatistically.Text = sqlDataReader["NStatistically"].ToString();
              this.txtPercentStatistically.Text = sqlDataReader["PercentStatistically"].ToString();
              this.txtDoNotFitStatistically.Text = sqlDataReader["DoNotFitStatistically"].ToString();
              this.cmbCalcID.Text = sqlDataReader["CalcID"].ToString();
              this.txtUser.Text = sqlDataReader["User"].ToString();
              this.txtTimePointData.Text = sqlDataReader["TimePointData"].ToString();
              this.txtTimePointCalc.Text = sqlDataReader["TimePointCalc"].ToString();
              this.txtNote.Text = sqlDataReader["Note"].ToString();
              this.chkActive.Checked = !(sqlDataReader["Active"].ToString() == "False");
            }
            sqlDataReader.Close();
            connection.Close();
          }
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + this.cmbCalcID.Text + "'";
            SqlCommand sqlCommand11 = new SqlCommand(cmdText, connection);
            sqlCommand11.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand11.ExecuteReader();
            SqlCommand sqlCommand12 = new SqlCommand(cmdText, connection);
            sqlDataReader.Close();
            connection.Close();
          }
          int num = (int) MessageBox.Show("Loading done!");
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.Message);
        }
      }
    }

    private void txtEntryDate_Validated(object sender, EventArgs e)
    {
      if (!string.IsNullOrEmpty(this.txtTimePointData.Text))
      {
        if (System.DateTime.TryParseExact(this.txtTimePointData.Text, "dd.MM.yyyy", (IFormatProvider) CultureInfo.InvariantCulture, DateTimeStyles.None, out System.DateTime _))
        {
          string text = this.txtTimePointData.Text;
          if (System.DateTime.TryParse(text, out System.DateTime _))
          {
            string[] formats = new string[1]{ "dd.MM.yyyy" };
            System.DateTime result1;
            bool exact = System.DateTime.TryParseExact(text, formats, (IFormatProvider) DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out result1);
            System.DateTime result2;
            exact = System.DateTime.TryParseExact(this.txtTimePointCalc.Text.Substring(0, 10), formats, (IFormatProvider) DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out result2);
            if (System.DateTime.Compare(result1, result2) <= 0)
            {
              Form1.dateIsGood = true;
            }
            else
            {
              Form1.dateIsGood = false;
              int num = (int) MessageBox.Show("The <timepoint data> can not be greater than <timepoint calculation>");
            }
          }
          else
            Form1.dateIsGood = false;
        }
        else
        {
          int num = (int) MessageBox.Show("Invalid date format date must be formatted to dd.mm.yyyy");
          Form1.dateIsGood = false;
          this.txtTimePointData.Focus();
        }
      }
      else
      {
        int num = (int) MessageBox.Show("Please provide entry date in the format of dd.mm.yyyy");
        Form1.dateIsGood = false;
        this.txtTimePointData.Focus();
      }
    }

    private void btnSave_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.txtEntryDate_Validated(sender, e);
        if (!Form1.dateIsGood)
          return;
        this.btnSave_Click(sender, e);
        int num = (int) MessageBox.Show("The data was saved!");
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void rbAll_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.rbAll_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void rbFitStat_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.rbFitStat_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void rbNotFitStat_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.rbNotFitStat_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public void SQLRunFill(string sqlQuery, ComboBox cmbInput)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(sqlQuery, connection);
          sqlCommand.CommandTimeout = 600;
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          cmbInput.Text = "All";
          while (sqlDataReader.Read())
            cmbInput.Items.Add((object) sqlDataReader[0].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public void SQLRunFillChechedListBox(string sqlQuery, CheckedListBox cmbInput)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(sqlQuery, connection);
          sqlCommand.CommandTimeout = 600;
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          cmbInput.Text = "All";
          while (sqlDataReader.Read())
            cmbInput.Items.Add((object) sqlDataReader[0].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public void SQLRunFillCheckedListBox(string sqlQuery, CheckedListBox lsInput)
    {
      try
      {
        lsInput.Items.Clear();
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(sqlQuery, connection);
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          int index = 0;
          while (sqlDataReader.Read())
          {
            lsInput.Items.Add((object) sqlDataReader[0].ToString());
            lsInput.SetItemChecked(index, true);
            ++index;
          }
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public void SQLRunFillListBox(string sqlQuery, ListBox lsInput)
    {
      try
      {
        lsInput.Items.Clear();
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(sqlQuery, connection);
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          lsInput.Items.Add((object) "All");
          lsInput.Text = "All";
          while (sqlDataReader.Read())
            lsInput.Items.Add((object) sqlDataReader[0].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void frmHistResults_Load(object sender, EventArgs e)
    {
      try
      {
        this.chDeActivate.Checked = false;
        this.chDeActivate.Enabled = false;
        this.chActivate.Checked = false;
        this.chActivate.Enabled = false;
        this.dtCalcDateTime.Enabled = false;
        this.clbVirtOzid.Enabled = false;
        this.rbAll1.Enabled = false;
        this.rbActive1.Enabled = false;
        this.rbNotActive1.Enabled = false;
        this.groupFilterSelection.Enabled = false;
        this.panelSelection.Enabled = false;
        this.panelButtons.Enabled = false;
        string[] strArray = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini");
        Form1.connectionString = strArray[0];
        Form1.strRscript = strArray[1];
        Form1.strRpath = strArray[2];
        Form1.strDataDir = strArray[3];
        Form1.strOutputDir = strArray[4];
        Form1.strOutPutPath = strArray[4];
        string sqlQuery1 = "select distinct VIRT_OZID from CalculationRaw order by VIRT_OZID";
        string sqlQuery2 = "select distinct CalcID from CalculationRaw order by CalcID";
        string sqlQuery3 = "select distinct PRODUCTCODE from CalculationRaw order by PRODUCTCODE";
        this.chkVirtOzid2.Items.Clear();
        this.SQLRunFillChechedListBox(sqlQuery1, this.chkVirtOzid2);
        this.cmbProdID2.Items.Clear();
        this.SQLRunFill(sqlQuery3, this.cmbProdID2);
        this.cmbCalcID2.Items.Clear();
        this.SQLRunFill(sqlQuery2, this.cmbCalcID2);
        this.lbOzid.Items.Clear();
        this.SQLRunFillListBox(sqlQuery1, this.lbOzid);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void rbFitStat_CheckedChanged1(object sender, EventArgs e) => Form1.strFilter += " and FitStatistically ='True'";

    private void rbNotFitStat_CheckedChanged1(object sender, EventArgs e) => Form1.strFilter += " and FitStatistically ='False'";

    private void tb_LostFocus1(object sender, EventArgs e)
    {
      try
      {
        TextBox textBox = (TextBox) sender;
        this.strMatrix[(int) short.Parse(textBox.Name.Substring(3, textBox.Name.Length - 3)), 10] = textBox.Text;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void btnGetHistoric_Click(object sender, EventArgs e)
    {
      this.btnZoomIn1.Visible = false;
      this.btnPrint.Visible = false;
      this.pictureBox2.Visible = false;
      this.label82.Visible = false;
      this.label83.Visible = false;
      Form1.strFilter = "";
      if (this.rbActive1.Checked)
        Form1.strFilter += " and Active = 'True'";
      if (this.rbNotActive1.Checked)
        Form1.strFilter += " and Active = 'False'";
      if (this.rbFitstatF.Checked)
        Form1.strFilter += " and FitStatistically ='True'";
      if (this.rbNotFitStatF.Checked)
        Form1.strFilter += " and FitStatistically ='False'";
      if (this.rbKPI0.Checked)
        Form1.strFilter += " and CAST(KPI0 AS real) > 0";
      if (this.rbKPI1.Checked)
        Form1.strFilter += " and CAST(KPI1 AS real) > 0";
      if (this.rbKPI2.Checked)
        Form1.strFilter += " and CAST(KPI2 AS real) > 0";
      if (this.rbKPI3.Checked)
        Form1.strFilter += " and CAST(KPI3 AS real) > 0";
      try
      {
        Form1.strQuery = " (1 = 1)     and 0=0";
        string str1 = " and Virt_Ozid in ('";
        for (int index = 0; index < this.chkVirtOzid2.Items.Count; ++index)
        {
          if (this.chkVirtOzid2.GetItemCheckState(index) == CheckState.Checked)
            str1 = str1 + (string) this.chkVirtOzid2.Items[index] + "','";
        }
        string str2 = this.Left(str1, str1.Length - 2) + ") ";
        if (str2 == " and Virt_Ozid in ) ")
          str2 = "";
        if (str2.IndexOf("()") <= 0)
          Form1.strQuery += str2;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcRow where CalcID = '" + this.cmbCalcID2.Text + "' and " + Form1.strQuery;
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (!Form1.IsCalcIDAvailable)
      {
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel6.Controls.Clear();
            string cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID2.Text + "') and CalcID = '" + this.cmbCalcID2.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
            sqlCommand3.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader1 = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
            int index1 = 0;
            while (sqlDataReader1.Read())
            {
              ++index1;
              TextBox textBox1 = new TextBox();
              textBox1.Name = "tb1" + index1.ToString();
              textBox1.Enabled = false;
              textBox1.Text = sqlDataReader1[1].ToString();
              textBox1.Top = index1 * 20;
              textBox1.Left = 3;
              TextBox textBox2 = textBox1;
              this.strMatrix[index1, 1] = sqlDataReader1[1].ToString();
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index1] = sqlDataReader1[1].ToString();
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 50;
              textBox3.Text = sqlDataReader1[2].ToString();
              textBox3.Top = index1 * 20;
              textBox3.Left = 110;
              TextBox textBox4 = textBox3;
              this.strMatrix[index1, 2] = sqlDataReader1[2].ToString();
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 50;
              textBox5.Text = sqlDataReader1[3].ToString();
              textBox5.Top = index1 * 20;
              textBox5.Left = 190;
              TextBox textBox6 = textBox5;
              this.strMatrix[index1, 3] = sqlDataReader1[3].ToString();
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = sqlDataReader1[4].ToString();
              textBox7.Top = index1 * 20;
              textBox7.Left = 270;
              TextBox textBox8 = textBox7;
              this.strMatrix[index1, 4] = sqlDataReader1[4].ToString();
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = sqlDataReader1[5].ToString();
              textBox9.Top = index1 * 20;
              textBox9.Left = 360;
              TextBox textBox10 = textBox9;
              this.strMatrix[index1, 5] = sqlDataReader1[5].ToString();
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = sqlDataReader1[6].ToString();
              textBox11.Top = index1 * 20;
              textBox11.Left = 440;
              TextBox textBox12 = textBox11;
              this.strMatrix[index1, 6] = sqlDataReader1[6].ToString();
              CheckBox checkBox1 = new CheckBox();
              checkBox1.Name = "ch1" + index1.ToString();
              checkBox1.Text = string.Format("{0}", (object) "yes");
              checkBox1.Top = index1 * 20;
              checkBox1.Left = 580;
              CheckBox checkBox2 = checkBox1;
              checkBox2.MouseLeave += new EventHandler(this.chk_MouseMove);
              if (sqlDataReader1[5].ToString() == "100")
                checkBox2.Checked = true;
              checkBox2.Click += new EventHandler(this.chk_Click);
              checkBox2.MouseLeave += new EventHandler(this.chk_MouseMove);
              string[,] strMatrix1 = this.strMatrix;
              int index2 = index1;
              bool flag = checkBox2.Checked;
              string str3 = flag.ToString();
              strMatrix1[index2, 7] = str3;
              Button button1 = new Button();
              button1.Name = "b1" + index1.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index1 * 20;
              button1.Left = 680;
              Button button2 = button1;
              int num1 = sqlDataReader1["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader1["FitStatistically"].ToString() == "True" ? 1 : 0);
              checkBox2.Checked = num1 != 0;
              this.panel6.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba1_Click);
              CheckBox checkBox3 = new CheckBox();
              checkBox3.Name = "ch2" + index1.ToString();
              checkBox3.Text = string.Format("{0}", (object) "yes");
              checkBox3.Top = index1 * 20;
              checkBox3.Left = 770;
              CheckBox checkBox4 = checkBox3;
              checkBox4.Click += new EventHandler(this.chk_Click);
              checkBox4.MouseLeave += new EventHandler(this.chk_MouseMove);
              string[,] strMatrix2 = this.strMatrix;
              int index3 = index1;
              flag = checkBox4.Checked;
              string str4 = flag.ToString();
              strMatrix2[index3, 8] = str4;
              TextBox textBox13 = new TextBox();
              textBox13.Name = "tb7" + index1.ToString();
              textBox13.Text = sqlDataReader1[12].ToString();
              textBox13.Width = 150;
              textBox13.Top = index1 * 20;
              textBox13.Left = 1050;
              TextBox textBox14 = textBox13;
              textBox14.LostFocus += new EventHandler(this.tb_LostFocus1);
              this.strMatrix[index1, 10] = textBox14.Text;
              int num2 = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox4.Checked = num2 != 0;
              this.panel6.Controls.Add((Control) textBox2);
              this.panel6.Controls.Add((Control) textBox4);
              this.panel6.Controls.Add((Control) textBox6);
              this.panel6.Controls.Add((Control) textBox8);
              this.panel6.Controls.Add((Control) textBox10);
              this.panel6.Controls.Add((Control) textBox12);
              this.panel6.Controls.Add((Control) checkBox2);
              this.panel6.Controls.Add((Control) checkBox4);
              this.panel6.Controls.Add((Control) textBox14);
              this.panel6.Controls.Add(this.ba);
            }
            sqlDataReader1.Close();
            string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID2.Text + "'";
            SqlDataReader sqlDataReader2 = sqlCommand4.ExecuteReader();
            SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
            int num = 0;
            while (sqlDataReader2.Read())
              ++num;
            string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
            sqlDataReader2.Close();
            SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
            string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal = '0'";
            sqlDataReader3.Close();
            SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
            string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID2.Text + "')";
            sqlDataReader4.Close();
            SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
            string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal != '0'  ";
            sqlDataReader5.Close();
            new SqlCommand(cmdText6, connection).ExecuteReader().Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source == null)
            return;
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      else
      {
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            this.panel6.Controls.Clear();
            string cmdText7 = "SELECT [VIRT_OZID],dbo.KPIcount0(calcid,[VIRT_OZID]) as 'c0',dbo.KPIcount1(calcid,[VIRT_OZID]) as 'c1',dbo.KPIcount2(calcid,[VIRT_OZID]) as 'c2',dbo.KPIcount3(calcid,[VIRT_OZID]) as 'c3',[KPI0],[KPI1],[KPI2],[KPI3],[FitStatistically],[RelevantForDiscussion],[GraphID],[Additional_note], Active FROM [VIRT_OZID_per_calculation] where CalcID = '" + this.cmbCalcID2.Text + "' and " + Form1.strQuery + Form1.strFilter;
            SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
            sqlCommand6.ExecuteScalar();
            new DataTable().Columns.Add("VIRT_OZID", typeof (string));
            SqlDataReader sqlDataReader6 = sqlCommand6.ExecuteReader();
            SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
            int index4 = 0;
            while (sqlDataReader6.Read())
            {
              ++index4;
              TextBox textBox15 = new TextBox();
              textBox15.Name = "tb1" + index4.ToString();
              textBox15.Enabled = false;
              textBox15.Width = 50;
              textBox15.Text = index4.ToString();
              textBox15.Top = index4 * 20;
              textBox15.Left = 3;
              TextBox textBox16 = textBox15;
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = false;
              textBox17.Width = 100;
              textBox17.Text = sqlDataReader6[0].ToString();
              textBox17.Top = index4 * 20;
              textBox17.Left = 100;
              TextBox textBox18 = textBox17;
              this.strMatrix[index4, 1] = sqlDataReader6[0].ToString();
              Form1.strOZID = textBox16.Text;
              Form1.strArr2[index4] = sqlDataReader6[0].ToString();
              TextBox textBox19 = new TextBox();
              textBox19.Enabled = false;
              textBox19.Width = 100;
              textBox19.Text = sqlDataReader6[1].ToString() + "," + sqlDataReader6[2].ToString() + "," + sqlDataReader6[3].ToString() + "," + sqlDataReader6[4].ToString();
              textBox19.Top = index4 * 20;
              textBox19.Left = 205;
              TextBox textBox20 = textBox19;
              this.strMatrix[index4, 2] = sqlDataReader6[2].ToString();
              TextBox textBox21 = new TextBox();
              textBox21.Enabled = false;
              textBox21.Width = 110;
              textBox21.Text = this.Left(sqlDataReader6[5].ToString(), 4) + ", " + this.Left(sqlDataReader6[6].ToString(), 4) + ", " + this.Left(sqlDataReader6[7].ToString(), 4) + ", " + this.Left(sqlDataReader6[8].ToString(), 4);
              textBox21.Top = index4 * 20;
              textBox21.Left = 350;
              TextBox textBox22 = textBox21;
              this.strMatrix[index4, 3] = this.Left(sqlDataReader6[5].ToString(), 4) + ", " + this.Left(sqlDataReader6[6].ToString(), 4) + ", " + this.Left(sqlDataReader6[7].ToString(), 4) + ", " + this.Left(sqlDataReader6[8].ToString(), 4);
              CheckBox checkBox5 = new CheckBox();
              checkBox5.Name = "ch1" + index4.ToString();
              checkBox5.Text = string.Format("{0}", (object) "yes");
              checkBox5.Top = index4 * 20;
              checkBox5.Left = 520;
              CheckBox checkBox6 = checkBox5;
              this.strMatrix[index4, 6] = sqlDataReader6[6].ToString();
              checkBox6.Click += new EventHandler(this.chk_Click);
              checkBox6.MouseLeave += new EventHandler(this.chk_MouseMove);
              Button button3 = new Button();
              button3.Name = "b1" + index4.ToString();
              button3.Text = string.Format("{0}", (object) "Chart");
              button3.Top = index4 * 20;
              button3.Left = 600;
              Button button4 = button3;
              int num3 = sqlDataReader6["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader6["FitStatistically"].ToString() == "True" ? 1 : 0);
              checkBox6.Checked = num3 != 0;
              if (this.Left(sqlDataReader6[5].ToString(), 4) == "100")
                checkBox6.Checked = true;
              this.panel6.Controls.Add((Control) button4);
              button4.Click += new EventHandler(this.ba1_Click);
              string[,] strMatrix3 = this.strMatrix;
              int index5 = index4;
              bool flag = checkBox6.Checked;
              string str5 = flag.ToString();
              strMatrix3[index5, 7] = str5;
              CheckBox checkBox7 = new CheckBox();
              checkBox7.Name = "ch2" + index4.ToString();
              checkBox7.Text = string.Format("{0}", (object) "yes");
              checkBox7.Top = index4 * 20;
              checkBox7.Left = 740;
              CheckBox checkBox8 = checkBox7;
              checkBox8.Click += new EventHandler(this.chk_Click);
              checkBox8.MouseLeave += new EventHandler(this.chk_MouseMove);
              CheckBox checkBox9 = new CheckBox();
              checkBox9.Name = "ch3" + index4.ToString();
              checkBox9.Text = string.Format("{0}", (object) "yes");
              checkBox9.Top = index4 * 20;
              checkBox9.Left = 920;
              CheckBox checkBox10 = checkBox9;
              checkBox10.Click += new EventHandler(this.chk_Click);
              checkBox10.MouseLeave += new EventHandler(this.chk_MouseMove);
              checkBox10.Checked = sqlDataReader6["active"].ToString() == "True";
              string[,] strMatrix4 = this.strMatrix;
              int index6 = index4;
              flag = checkBox10.Checked;
              string str6 = flag.ToString();
              strMatrix4[index6, 9] = str6;
              checkBox10.Click += new EventHandler(this.chk_Click);
              TextBox textBox23 = new TextBox();
              textBox23.Name = "tb7" + index4.ToString();
              textBox23.Text = sqlDataReader6[12].ToString();
              textBox23.Width = 150;
              textBox23.Top = index4 * 20;
              textBox23.Left = 1050;
              TextBox textBox24 = textBox23;
              textBox24.LostFocus += new EventHandler(this.tb_LostFocus1);
              this.strMatrix[index4, 10] = textBox24.Text;
              int num4 = sqlDataReader6["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader6["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
              checkBox8.Checked = num4 != 0;
              string[,] strMatrix5 = this.strMatrix;
              int index7 = index4;
              flag = checkBox8.Checked;
              string str7 = flag.ToString();
              strMatrix5[index7, 8] = str7;
              textBox24.LostFocus += new EventHandler(this.tb_LostFocus);
              this.panel6.Controls.Add((Control) textBox16);
              this.panel6.Controls.Add((Control) textBox18);
              this.panel6.Controls.Add((Control) textBox20);
              this.panel6.Controls.Add((Control) textBox22);
              this.panel6.Controls.Add((Control) checkBox6);
              this.panel6.Controls.Add((Control) checkBox8);
              this.panel6.Controls.Add((Control) checkBox10);
              this.panel6.Controls.Add((Control) textBox24);
              this.panel6.Controls.Add(this.ba);
            }
            Form1.strRowsCount = index4;
            sqlDataReader6.Close();
            string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID2.Text + "'";
            SqlDataReader sqlDataReader7 = sqlCommand7.ExecuteReader();
            SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
            int num = 0;
            while (sqlDataReader7.Read())
              ++num;
            string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
            sqlDataReader7.Close();
            SqlDataReader sqlDataReader8 = new SqlCommand(cmdText9, connection).ExecuteReader();
            string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal = '0'";
            sqlDataReader8.Close();
            SqlDataReader sqlDataReader9 = new SqlCommand(cmdText10, connection).ExecuteReader();
            string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID2.Text + "')";
            sqlDataReader9.Close();
            SqlDataReader sqlDataReader10 = new SqlCommand(cmdText11, connection).ExecuteReader();
            string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal != '0'  ";
            sqlDataReader10.Close();
            new SqlCommand(cmdText12, connection).ExecuteReader().Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          if (ex.Source != null)
          {
            int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
          }
        }
      }
    }

    private void cmbProductCode_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.cmbCalcID.Items.Clear();
        if (this.cmbProdID2.Text == "All")
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw";
            SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
            sqlCommand1.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
            SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID2.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + this.cmbProdID2.Text + "'";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
            sqlCommand3.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID2.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void radioButton13_CheckedChanged(object sender, EventArgs e)
    {
    }

    private void rbActive_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and Active ='True'";

    private void rbNotActive_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and Active ='False'";

    private void radioButton9_CheckedChanged(object sender, EventArgs e)
    {
    }

    private void radioButton8_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and Active ='True'";

    private void radioButton7_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and Active ='False'";

    private void radioButton1_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and CAST(KPI0 AS real) > 0";

    private void radioButton2_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and CAST(KPI1 AS real) > 0";

    private void radioButton3_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and CAST(KPI2 AS real) > 0";

    private void radioButton10_CheckedChanged(object sender, EventArgs e) => Form1.strFilter += " and CAST(KPI3 AS real) > 0";

    private void clbVirtOzid_SelectedIndexChanged(object sender, EventArgs e)
    {
      this.lbOzid.Items.Clear();
      this.lbOzid.Items.Add((object) this.clbVirtOzid.Text);
      if (this.cmbCalcID2.Text == "All")
      {
        int num1 = (int) MessageBox.Show("Please select Calculation first!");
      }
      else
      {
        try
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct * from CalcRow where CalcID = '" + this.cmbCalcID2.Text + "' and Virt_Ozid = '" + this.clbVirtOzid.Text + "'";
            SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
            sqlCommand1.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
            SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              Form1.IsCalcIDAvailable = true;
            sqlDataReader.Close();
            connection.Close();
          }
        }
        catch (Exception ex)
        {
          int num2 = (int) MessageBox.Show(ex.Message);
        }
        if (!Form1.IsCalcIDAvailable)
        {
          try
          {
            using (SqlConnection connection = new SqlConnection(Form1.connectionString))
            {
              connection.Open();
              this.panel6.Controls.Clear();
              string cmdText1 = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + this.cmbCalcID2.Text + "') and CalcID = '" + this.cmbCalcID2.Text + "' and Virt_Ozid = '" + this.clbVirtOzid.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + this.cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
              SqlCommand sqlCommand3 = new SqlCommand(cmdText1, connection);
              sqlCommand3.ExecuteScalar();
              new DataTable().Columns.Add("VIRT_OZID", typeof (string));
              SqlDataReader sqlDataReader1 = sqlCommand3.ExecuteReader();
              SqlCommand sqlCommand4 = new SqlCommand(cmdText1, connection);
              int index = 0;
              while (sqlDataReader1.Read())
              {
                ++index;
                TextBox textBox1 = new TextBox();
                textBox1.Name = "tb1" + index.ToString();
                textBox1.Enabled = false;
                textBox1.Text = sqlDataReader1[1].ToString();
                textBox1.Top = index * 20;
                textBox1.Left = 3;
                TextBox textBox2 = textBox1;
                this.strMatrix[index, 1] = sqlDataReader1[1].ToString();
                Form1.strOZID = textBox2.Text;
                Form1.strArr2[index] = sqlDataReader1[1].ToString();
                TextBox textBox3 = new TextBox();
                textBox3.Enabled = false;
                textBox3.Width = 50;
                textBox3.Text = sqlDataReader1[2].ToString();
                textBox3.Top = index * 20;
                textBox3.Left = 110;
                TextBox textBox4 = textBox3;
                this.strMatrix[index, 2] = sqlDataReader1[2].ToString();
                TextBox textBox5 = new TextBox();
                textBox5.Enabled = false;
                textBox5.Width = 50;
                textBox5.Text = sqlDataReader1[3].ToString();
                textBox5.Top = index * 20;
                textBox5.Left = 190;
                TextBox textBox6 = textBox5;
                this.strMatrix[index, 3] = sqlDataReader1[3].ToString();
                TextBox textBox7 = new TextBox();
                textBox7.Enabled = false;
                textBox7.Width = 50;
                textBox7.Text = sqlDataReader1[4].ToString();
                textBox7.Top = index * 20;
                textBox7.Left = 270;
                TextBox textBox8 = textBox7;
                this.strMatrix[index, 4] = sqlDataReader1[4].ToString();
                TextBox textBox9 = new TextBox();
                textBox9.Enabled = false;
                textBox9.Width = 50;
                textBox9.Text = sqlDataReader1[5].ToString();
                textBox9.Top = index * 20;
                textBox9.Left = 360;
                TextBox textBox10 = textBox9;
                this.strMatrix[index, 5] = sqlDataReader1[5].ToString();
                TextBox textBox11 = new TextBox();
                textBox11.Enabled = false;
                textBox11.Width = 50;
                textBox11.Text = sqlDataReader1[6].ToString();
                textBox11.Top = index * 20;
                textBox11.Left = 440;
                TextBox textBox12 = textBox11;
                this.strMatrix[index, 6] = sqlDataReader1[6].ToString();
                CheckBox checkBox1 = new CheckBox();
                checkBox1.Name = "ch1" + index.ToString();
                checkBox1.Text = string.Format("{0}", (object) "yes");
                checkBox1.Top = index * 20;
                checkBox1.Left = 580;
                CheckBox checkBox2 = checkBox1;
                checkBox2.Click += new EventHandler(this.chk_Click);
                this.strMatrix[index, 7] = checkBox2.Checked.ToString();
                Button button1 = new Button();
                button1.Name = "b1" + index.ToString();
                button1.Text = string.Format("{0}", (object) "Chart");
                button1.Top = index * 20;
                button1.Left = 680;
                Button button2 = button1;
                int num3 = sqlDataReader1["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader1["FitStatistically"].ToString() == "True" ? 1 : 0);
                checkBox2.Checked = num3 != 0;
                this.panel6.Controls.Add((Control) button2);
                button2.Click += new EventHandler(this.ba_Click);
                CheckBox checkBox3 = new CheckBox();
                checkBox3.Name = "ch2" + index.ToString();
                checkBox3.Text = string.Format("{0}", (object) "yes");
                checkBox3.Top = index * 20;
                checkBox3.Left = 770;
                CheckBox checkBox4 = checkBox3;
                checkBox4.Click += new EventHandler(this.chk_Click);
                this.strMatrix[index, 8] = checkBox4.Checked.ToString();
                TextBox textBox13 = new TextBox();
                textBox13.Name = "tb7" + index.ToString();
                textBox13.Text = sqlDataReader1[12].ToString();
                textBox13.Width = 150;
                textBox13.Top = index * 20;
                textBox13.Left = 1050;
                TextBox textBox14 = textBox13;
                textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
                this.strMatrix[index, 10] = textBox14.Text;
                int num4 = sqlDataReader1["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader1["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
                checkBox4.Checked = num4 != 0;
                textBox14.LostFocus += new EventHandler(this.tb_LostFocus);
                this.panel6.Controls.Add((Control) textBox2);
                this.panel6.Controls.Add((Control) textBox4);
                this.panel6.Controls.Add((Control) textBox6);
                this.panel6.Controls.Add((Control) textBox8);
                this.panel6.Controls.Add((Control) textBox10);
                this.panel6.Controls.Add((Control) textBox12);
                this.panel6.Controls.Add((Control) checkBox2);
                this.panel6.Controls.Add((Control) checkBox4);
                this.panel6.Controls.Add((Control) textBox14);
                this.panel6.Controls.Add(this.ba);
              }
              sqlDataReader1.Close();
              string cmdText2 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID2.Text + "'";
              SqlDataReader sqlDataReader2 = sqlCommand4.ExecuteReader();
              SqlCommand sqlCommand5 = new SqlCommand(cmdText2, connection);
              int num5 = 0;
              while (sqlDataReader2.Read())
                ++num5;
              string cmdText3 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
              sqlDataReader2.Close();
              SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
              string cmdText4 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal = '0'";
              sqlDataReader3.Close();
              SqlDataReader sqlDataReader4 = new SqlCommand(cmdText4, connection).ExecuteReader();
              string cmdText5 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID2.Text + "')";
              sqlDataReader4.Close();
              SqlDataReader sqlDataReader5 = new SqlCommand(cmdText5, connection).ExecuteReader();
              string cmdText6 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal != '0'  ";
              sqlDataReader5.Close();
              new SqlCommand(cmdText6, connection).ExecuteReader().Close();
              connection.Close();
            }
          }
          catch (Exception ex)
          {
            if (ex.Source != null)
            {
              int num6 = (int) MessageBox.Show("IOException source: {0}", ex.Message);
            }
          }
        }
        else
        {
          try
          {
            using (SqlConnection connection = new SqlConnection(Form1.connectionString))
            {
              connection.Open();
              this.panel6.Controls.Clear();
              string cmdText7 = "SELECT [VIRT_OZID],dbo.KPIcount0(calcid,[VIRT_OZID]) as 'c0',dbo.KPIcount1(calcid,[VIRT_OZID]) as 'c1',dbo.KPIcount2(calcid,[VIRT_OZID]) as 'c2',dbo.KPIcount3(calcid,[VIRT_OZID]) as 'c3',[KPI0],[KPI1],[KPI2],[KPI3],[FitStatistically],[RelevantForDiscussion],[GraphID],[Additional_note], Active FROM[VIRT_OZID_per_calculation] where CalcID = '" + this.cmbCalcID2.Text + "' and Virt_Ozid = '" + this.clbVirtOzid.Text + "'";
              SqlCommand sqlCommand6 = new SqlCommand(cmdText7, connection);
              sqlCommand6.ExecuteScalar();
              new DataTable().Columns.Add("VIRT_OZID", typeof (string));
              SqlDataReader sqlDataReader6 = sqlCommand6.ExecuteReader();
              SqlCommand sqlCommand7 = new SqlCommand(cmdText7, connection);
              int index1 = 0;
              while (sqlDataReader6.Read())
              {
                ++index1;
                TextBox textBox15 = new TextBox();
                textBox15.Name = "tb1" + index1.ToString();
                textBox15.Enabled = false;
                textBox15.Width = 50;
                textBox15.Text = index1.ToString();
                textBox15.Top = index1 * 20;
                textBox15.Left = 3;
                TextBox textBox16 = textBox15;
                TextBox textBox17 = new TextBox();
                textBox17.Enabled = false;
                textBox17.Width = 100;
                textBox17.Text = sqlDataReader6[0].ToString();
                textBox17.Top = index1 * 20;
                textBox17.Left = 100;
                TextBox textBox18 = textBox17;
                this.strMatrix[index1, 1] = sqlDataReader6[0].ToString();
                Form1.strOZID = textBox16.Text;
                Form1.strArr2[index1] = sqlDataReader6[0].ToString();
                TextBox textBox19 = new TextBox();
                textBox19.Enabled = false;
                textBox19.Width = 100;
                textBox19.Text = sqlDataReader6[1].ToString() + "," + sqlDataReader6[2].ToString() + "," + sqlDataReader6[3].ToString() + "," + sqlDataReader6[4].ToString();
                textBox19.Top = index1 * 20;
                textBox19.Left = 205;
                TextBox textBox20 = textBox19;
                this.strMatrix[index1, 2] = sqlDataReader6[2].ToString();
                TextBox textBox21 = new TextBox();
                textBox21.Enabled = false;
                textBox21.Width = 110;
                textBox21.Text = this.Left(sqlDataReader6[5].ToString(), 4) + ", " + this.Left(sqlDataReader6[6].ToString(), 4) + ", " + this.Left(sqlDataReader6[7].ToString(), 4) + ", " + this.Left(sqlDataReader6[8].ToString(), 4);
                textBox21.Top = index1 * 20;
                textBox21.Left = 350;
                TextBox textBox22 = textBox21;
                this.strMatrix[index1, 3] = this.Left(sqlDataReader6[5].ToString(), 4) + ", " + this.Left(sqlDataReader6[6].ToString(), 4) + ", " + this.Left(sqlDataReader6[7].ToString(), 4) + ", " + this.Left(sqlDataReader6[8].ToString(), 4);
                CheckBox checkBox5 = new CheckBox();
                checkBox5.Name = "ch1" + index1.ToString();
                checkBox5.Text = string.Format("{0}", (object) "yes");
                checkBox5.Top = index1 * 20;
                checkBox5.Left = 520;
                CheckBox checkBox6 = checkBox5;
                this.strMatrix[index1, 6] = sqlDataReader6[6].ToString();
                checkBox6.Click += new EventHandler(this.chk_Click);
                Button button3 = new Button();
                button3.Name = "b1" + index1.ToString();
                button3.Text = string.Format("{0}", (object) "Chart");
                button3.Top = index1 * 20;
                button3.Left = 600;
                Button button4 = button3;
                int num7 = sqlDataReader6["FitStatistically"].ToString() == "true" ? 1 : (sqlDataReader6["FitStatistically"].ToString() == "True" ? 1 : 0);
                checkBox6.Checked = num7 != 0;
                this.panel6.Controls.Add((Control) button4);
                button4.Click += new EventHandler(this.ba_Click);
                string[,] strMatrix1 = this.strMatrix;
                int index2 = index1;
                bool flag = checkBox6.Checked;
                string str1 = flag.ToString();
                strMatrix1[index2, 7] = str1;
                CheckBox checkBox7 = new CheckBox();
                checkBox7.Name = "ch2" + index1.ToString();
                checkBox7.Text = string.Format("{0}", (object) "yes");
                checkBox7.Top = index1 * 20;
                checkBox7.Left = 740;
                CheckBox checkBox8 = checkBox7;
                checkBox8.Click += new EventHandler(this.chk_Click);
                CheckBox checkBox9 = new CheckBox();
                checkBox9.Name = "ch3" + index1.ToString();
                checkBox9.Text = string.Format("{0}", (object) "yes");
                checkBox9.Top = index1 * 20;
                checkBox9.Left = 920;
                CheckBox checkBox10 = checkBox9;
                checkBox10.Click += new EventHandler(this.chk_Click);
                checkBox10.Checked = sqlDataReader6["active"].ToString() == "True";
                string[,] strMatrix2 = this.strMatrix;
                int index3 = index1;
                flag = checkBox10.Checked;
                string str2 = flag.ToString();
                strMatrix2[index3, 9] = str2;
                checkBox10.Click += new EventHandler(this.chk_Click);
                TextBox textBox23 = new TextBox();
                textBox23.Name = "tb7" + index1.ToString();
                textBox23.Text = sqlDataReader6[12].ToString();
                textBox23.Width = 150;
                textBox23.Top = index1 * 20;
                textBox23.Left = 1050;
                TextBox textBox24 = textBox23;
                textBox24.LostFocus += new EventHandler(this.tb_LostFocus);
                this.strMatrix[index1, 10] = textBox24.Text;
                int num8 = sqlDataReader6["RelevantForDiscussion"].ToString() == "true" ? 1 : (sqlDataReader6["RelevantForDiscussion"].ToString() == "True" ? 1 : 0);
                checkBox8.Checked = num8 != 0;
                string[,] strMatrix3 = this.strMatrix;
                int index4 = index1;
                flag = checkBox8.Checked;
                string str3 = flag.ToString();
                strMatrix3[index4, 8] = str3;
                textBox24.LostFocus += new EventHandler(this.tb_LostFocus);
                this.panel6.Controls.Add((Control) textBox16);
                this.panel6.Controls.Add((Control) textBox18);
                this.panel6.Controls.Add((Control) textBox20);
                this.panel6.Controls.Add((Control) textBox22);
                this.panel6.Controls.Add((Control) checkBox6);
                this.panel6.Controls.Add((Control) checkBox8);
                this.panel6.Controls.Add((Control) checkBox10);
                this.panel6.Controls.Add((Control) textBox24);
                this.panel6.Controls.Add(this.ba);
              }
              sqlDataReader6.Close();
              string cmdText8 = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + this.cmbCalcID2.Text + "'";
              SqlDataReader sqlDataReader7 = sqlCommand7.ExecuteReader();
              SqlCommand sqlCommand8 = new SqlCommand(cmdText8, connection);
              int num9 = 0;
              while (sqlDataReader7.Read())
                ++num9;
              string cmdText9 = "select distinct CalcDate  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
              sqlDataReader7.Close();
              SqlDataReader sqlDataReader8 = new SqlCommand(cmdText9, connection).ExecuteReader();
              string cmdText10 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal = '0'";
              sqlDataReader8.Close();
              SqlDataReader sqlDataReader9 = new SqlCommand(cmdText10, connection).ExecuteReader();
              string cmdText11 = "select dbo.PerStatisticallyFit  ('" + this.cmbCalcID2.Text + "')";
              sqlDataReader9.Close();
              SqlDataReader sqlDataReader10 = new SqlCommand(cmdText11, connection).ExecuteReader();
              string cmdText12 = "select distinct count(signal)  from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "' and signal != '0'  ";
              sqlDataReader10.Close();
              new SqlCommand(cmdText12, connection).ExecuteReader().Close();
              connection.Close();
            }
          }
          catch (Exception ex)
          {
            if (ex.Source != null)
            {
              int num10 = (int) MessageBox.Show("IOException source: {0}", ex.Message);
            }
          }
        }
      }
    }

    private void button2_Click(object sender, EventArgs e)
    {
      try
      {
        Form1.connectionString = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini")[0];
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
        SqlCommand sqlCommand3 = (SqlCommand) null;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          sqlCommand3 = (SqlCommand) null;
          Form1.CalcID = this.cmbCalcID2.Text;
          Form1.Note = this.txtNote.Text;
          if (this.chActivate.Checked)
            Form1.Active = 1;
          if (this.chDeActivate.Checked)
            Form1.Active = 0;
          SqlCommand sqlCommand4 = new SqlCommand("update CalcResultView " + " set Active  ='" + Form1.Active.ToString() + "'", connection);
          sqlCommand4.Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
          sqlCommand4.ExecuteNonQuery();
          if (Form1.strRowsCount <= 0)
            return;
          for (int index = 1; index <= Form1.strRowsCount; ++index)
          {
            if (!this.rbAll.Checked)
              ;
            Form1.OZID = this.strMatrix[index, 1];
            Form1.CalcID = this.cmbCalcID2.Text;
            Form1.TotalN = this.strMatrix[index, 2];
            Form1.KPI0 = this.strMatrix[index, 3];
            Form1.KPI1 = this.strMatrix[index, 4];
            Form1.KPI2 = this.strMatrix[index, 5];
            Form1.KPI3 = this.strMatrix[index, 6];
            Form1.FitStatistically = this.strMatrix[index, 7];
            Form1.RelevantForDiscussion = this.strMatrix[index, 8];
            string str = this.strMatrix[index, 9];
            Form1.Additional_note = this.strMatrix[index, 10];
            string cmdText = "update VIRT_OZID_per_calculation " + " set FitStatistically='" + Form1.FitStatistically + "'," + " RelevantForDiscussion='" + Form1.RelevantForDiscussion + "'," + " Additional_note='" + Form1.Additional_note + "'," + " Active ='" + str + "' where VIRT_OZID='" + Form1.OZID + "' and CalcID='" + this.cmbCalcID2.Text + "'";
            if (this.strMatrix[index, 1] != null)
            {
              SqlCommand sqlCommand5 = new SqlCommand(cmdText, connection);
              sqlCommand5.Parameters.Add("@VIRT_OZID", SqlDbType.NVarChar, 250).Value = (object) Form1.OZID;
              sqlCommand5.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
              sqlCommand5.Parameters.Add("@Additional_note", SqlDbType.NVarChar, 250).Value = (object) Form1.Additional_note;
              sqlCommand5.Parameters.Add("@FitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.FitStatistically;
              sqlCommand5.Parameters.Add("@RelevantForDiscussion", SqlDbType.NVarChar, 250).Value = (object) Form1.RelevantForDiscussion;
              sqlCommand5.Parameters.Add("@Active", SqlDbType.NVarChar, 250).Value = (object) str;
              sqlCommand5.ExecuteNonQuery();
            }
          }
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chActivate_Click(object sender, EventArgs e)
    {
      this.chDeActivate.Enabled = true;
      this.chDeActivate.Checked = false;
      this.chActivate.Checked = true;
      this.chActivate.Enabled = true;
    }

    private void chDeActivate_Click(object sender, EventArgs e)
    {
      this.chActivate.Enabled = true;
      this.chActivate.Checked = false;
      this.chDeActivate.Checked = true;
      this.chDeActivate.Enabled = true;
    }

    private void chDeActivate_CheckedChanged(object sender, EventArgs e)
    {
    }

    private void btnGetHistoric_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.btnGetHistoric_Click(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.radioButton3_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbProdID2_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.cmbCalcID2.Items.Clear();
        if (this.cmbProdID2.Text == "All")
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw";
            SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
            sqlCommand1.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
            SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID2.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + this.cmbProdID2.Text + "'";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
            sqlCommand3.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID2.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbCalcID2_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.btnZoomIn1.Visible = false;
        this.pictureBox2.Visible = false;
        this.btnPrint.Visible = false;
        this.label82.Visible = false;
        this.label83.Visible = false;
        this.dtCalcDateTime.Enabled = true;
        this.clbVirtOzid.Enabled = true;
        this.rbAll1.Enabled = true;
        this.rbActive1.Enabled = true;
        this.rbNotActive1.Enabled = true;
        this.groupFilterSelection.Enabled = true;
        this.panelSelection.Enabled = true;
        this.panelButtons.Enabled = true;
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand("select distinct active from CalcResultView where calcid='" + this.cmbCalcID2.Text + "'", connection);
          sqlCommand.CommandTimeout = 600;
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          string str = "0";
          while (sqlDataReader.Read())
            str = sqlDataReader[0].ToString();
          if (str == "False")
          {
            this.chDeActivate.Checked = true;
            this.chDeActivate.Enabled = false;
            this.chActivate.Checked = false;
            this.chActivate.Enabled = true;
          }
          else
          {
            this.chActivate.Checked = true;
            this.chActivate.Enabled = false;
            this.chDeActivate.Checked = false;
            this.chDeActivate.Enabled = true;
          }
          sqlDataReader.Close();
          connection.Close();
        }
        string str1 = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID2.Text + "'";
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.txtNote1.Text = sqlDataReader["Note"].ToString();
          sqlDataReader.Close();
          connection.Close();
        }
        string sqlQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + this.cmbCalcID2.Text + "' order by VIRT_OZID";
        this.chkVirtOzid2.Items.Clear();
        this.lbOzid.Items.Clear();
        this.SQLRunFillCheckedListBox(sqlQuery, this.chkVirtOzid2);
        this.SQLRunFillListBox(sqlQuery, this.lbOzid);
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalculationRaw where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
          sqlCommand3.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
          SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.dtCalcDateTime.Value = System.DateTime.Parse(sqlDataReader[1].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      this.btnGetHistoric_Click(sender, e);
    }

    private void rbKPI0_CheckedChanged(object sender, EventArgs e)
    {
    }

    private void chkAllRefCpv_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chkAllRefCpv_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chkAllVirtOzid_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chkAllVirtOzid_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void SearchCalculation_onload(object sender, EventArgs e)
    {
      try
      {
        this.tabMain.SelectTab("SearchCalculation");
        this.frmHistResults_Load(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chkSortDate_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chkSortDate_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chLastN_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chLastN_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chLastM_CheckedChanged_1(object sender, EventArgs e) => this.chLastM_CheckedChanged(sender, e);

    private void chkLaufNr_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chkLaufNr_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void rbDays_CheckedChanged(object sender, EventArgs e)
    {
    }

    private void rbKPI3_CheckedChanged(object sender, EventArgs e)
    {
    }

    public void FillDGV(string queury)
    {
      string str1 = queury;
      string[] strArray = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini");
      Form1.connectionString = strArray[0];
      string str2 = strArray[1];
      string str3 = strArray[2];
      string str4 = strArray[3];
      string str5 = strArray[4];
      Form1.strOutPutPath = strArray[4];
      SqlConnection sqlConnection = new SqlConnection();
      sqlConnection.ConnectionString = Form1.connectionString;
      SqlDataAdapter sqlDataAdapter1 = new SqlDataAdapter(new SqlCommand(str1, sqlConnection));
      DataSet dataSet = new DataSet();
      try
      {
        sqlConnection.Open();
        sqlDataAdapter1.Fill(dataSet);
        sqlConnection.Close();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      if (dataSet.Tables.Count <= 0)
        return;
      SqlDataAdapter sqlDataAdapter2 = new SqlDataAdapter(str1, sqlConnection);
      sqlDataAdapter1.Fill(dataSet, "sql");
      this.dataGridViewRaw.DataSource = (object) dataSet.Tables["sql"];
      sqlConnection.Close();
    }

    private void cmbProductCode_SelectedIndexChanged_1(object sender, EventArgs e)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "SELECT * FROM Products where ProduktCode = '" + this.cmbProductCode.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          sqlDataReader.Close();
          connection.Close();
          for (int index = 0; index <= Form1.intCounterDataGridViews - 1; ++index)
            Form1.arrDataGridView[index].Visible = false;
          Form1.arrDataGridView[Form1.intIndex].Visible = false;
          this.dataGridViewRaw.Visible = true;
          this.FillDGV("select * from products where PRODUKTCODE = '" + this.cmbProductCode.Text + "'");
        }
        Form1.strQuery = "select distinct VIRT_OZID from Products where PRODUKTCODE = '" + this.cmbProductCode.Text + "' order by VIRT_OZID";
        string cmdText1 = "select distinct REFERENCED_CPV from Products where PRODUKTCODE = '" + this.cmbProductCode.Text + "' order by REFERENCED_CPV";
        string str = "";
        str = !(this.cmbProductCode.Text == "") ? "select distinct PRODUKTCODE from Products where PRODUKTCODE!='" + this.cmbProductCode.Text + "'" : "select distinct PRODUKTCODE from Products ";
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(Form1.strQuery, connection);
          sqlCommand.ExecuteScalar();
          sqlCommand.CommandTimeout = 600;
          new DataTable().Columns.Add("VIRT_OZID", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          this.clbVirtOzid.Items.Clear();
          while (sqlDataReader.Read())
          {
            this.clbVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
            this.lstCheckExclVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
          }
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(cmdText1, connection);
          sqlCommand.ExecuteScalar();
          sqlCommand.CommandTimeout = 600;
          new DataTable().Columns.Add("REFERENCED_CPV", typeof (string));
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          this.chkAllRefCpv.Checked = true;
          this.clbRefCpv.Items.Clear();
          while (sqlDataReader.Read())
            this.clbRefCpv.Items.Add((object) sqlDataReader["REFERENCED_CPV"].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
        this.cmbProductCode_SelectedIndexChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void lbOzid_SelectedIndexChanged_1(object sender, EventArgs e)
    {
      try
      {
        string cmdText = "SELECT [GraphName],[VIRT_OZID],[ImageValue],[CalcID],[ID] FROM [Graphs] where calcid='" + this.cmbCalcID2.Text + "' and VIRT_OZID ='" + this.lbOzid.SelectedItem.ToString() + "'";
        this.lbGraph.Items.Clear();
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
          sqlCommand.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          while (sqlDataReader.Read())
            this.lbGraph.Items.Add((object) sqlDataReader["GraphName"].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void lbGraph_SelectedIndexChanged_1(object sender, EventArgs e)
    {
      try
      {
        Form1.blnPressed = true;
        this.btnZoomIn1.Visible = true;
        this.btnPrint.Visible = true;
        this.picGraph.Visible = true;
        this.pictureBox2.Visible = true;
        SqlConnection connection = new SqlConnection(Form1.connectionString);
        connection.Open();
        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID2.Text + "' and VIRT_OZID='" + this.lbOzid.SelectedItem.ToString() + "'", connection));
        DataSet dataSet = new DataSet();
        sqlDataAdapter.Fill(dataSet, "Graphs");
        int count = dataSet.Tables["Graphs"].Rows.Count;
        if (count <= 0)
          return;
        this.pictureBox2.Image = Image.FromStream((Stream) new MemoryStream((byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"]));
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void button3_Click(object sender, EventArgs e)
    {
      try
      {
        Form1.connectionString = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\app.ini")[0];
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID2.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable = true;
          sqlDataReader.Close();
          connection.Close();
        }
        if (!Form1.IsCalcIDAvailable)
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            Form1.NParameterTotal = this.txtNParameterTotal.Text;
            Form1.NStatistically = this.txtNStatistically.Text;
            Form1.PercentStatistically = this.txtPercentStatistically.Text;
            Form1.DoNotFitStatistically = this.txtDoNotFitStatistically.Text;
            Form1.CalcID = this.cmbCalcID2.Text;
            Form1.User = this.txtUser.Text;
            Form1.TimePointData = this.txtTimePointData.Text;
            Form1.TimePointCalc = this.txtTimePointCalc.Text;
            Form1.Note = this.txtNote.Text;
            Form1.Active = !this.chkActive.Checked ? 0 : 1;
            SqlCommand sqlCommand3 = new SqlCommand("insert into CalcResultView (ID, NParameterTotal, " + "NStatistically, PercentStatistically, DoNotFitStatistically, CalcID, [User],TimePointData,TimePointCalc, Note,Active) values(" + " '" + Form1.ID.ToString() + "','" + Form1.NParameterTotal + "','" + Form1.NStatistically + "','" + Form1.PercentStatistically + "','" + Form1.DoNotFitStatistically + "','" + Form1.CalcID + "','" + Form1.User + "','" + Form1.TimePointData + "','" + Form1.TimePointCalc + "','" + Form1.Note + "','" + Form1.Active.ToString() + "')", connection);
            sqlCommand3.Parameters.Add("@ID", SqlDbType.Int).Value = (object) Form1.ID;
            sqlCommand3.Parameters.Add("@NParameterTotal", SqlDbType.NVarChar, 50).Value = (object) Form1.NParameterTotal;
            sqlCommand3.Parameters.Add("@NStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.NStatistically;
            sqlCommand3.Parameters.Add("@PercentStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.PercentStatistically;
            sqlCommand3.Parameters.Add("@DoNotFitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.DoNotFitStatistically;
            sqlCommand3.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
            sqlCommand3.Parameters.Add("@User", SqlDbType.NVarChar, 50).Value = (object) Form1.User;
            sqlCommand3.Parameters.Add("@TimePointData", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointData;
            sqlCommand3.Parameters.Add("@TimePointCalc", SqlDbType.NVarChar, 50).Value = (object) Form1.TimePointCalc;
            sqlCommand3.Parameters.Add("@Note", SqlDbType.NVarChar, 250).Value = (object) Form1.Note;
            sqlCommand3.Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
            sqlCommand3.ExecuteNonQuery();
            if (short.Parse(this.txtNParameterTotal.Text) > (short) 0)
            {
              for (int index = 1; index <= (int) short.Parse(this.txtNParameterTotal.Text); ++index)
              {
                Form1.OZID = this.strMatrix[index, 1];
                Form1.CalcID = this.cmbCalcID2.Text;
                Form1.TotalN = this.strMatrix[index, 2];
                Form1.KPI0 = this.strMatrix[index, 3];
                Form1.KPI1 = this.strMatrix[index, 4];
                Form1.KPI2 = this.strMatrix[index, 5];
                Form1.KPI3 = this.strMatrix[index, 6];
                Form1.FitStatistically = this.strMatrix[index, 7];
                Form1.RelevantForDiscussion = this.strMatrix[index, 8];
                Form1.Additional_note = this.strMatrix[index, 9];
                string cmdText = "select ID from Graphs where VIRT_OZID ='" + Form1.OZID + "' and CalcID = '" + this.cmbCalcID2.Text + "'";
                SqlDataReader sqlDataReader = new SqlCommand(cmdText, connection).ExecuteReader();
                SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
                while (sqlDataReader.Read())
                  Form1.GraphID = sqlDataReader["ID"].ToString();
                sqlDataReader.Close();
                SqlCommand sqlCommand5 = new SqlCommand("insert into VIRT_OZID_per_calculation (VIRT_OZID,CalcID,TotalN,KPI0,KPI1,KPI2,KPI3,Additional_note,FitStatistically,RelevantForDiscussion,GraphID) " + " values(" + " '" + Form1.OZID + "','" + Form1.CalcID + "','" + Form1.TotalN + "','" + Form1.KPI0 + "','" + Form1.KPI1 + "','" + Form1.KPI2 + "','" + Form1.KPI3 + "','" + Form1.Additional_note + "','" + Form1.FitStatistically + "','" + Form1.RelevantForDiscussion + "','" + Form1.GraphID + "')", connection);
                sqlCommand5.Parameters.Add("@VIRT_OZID", SqlDbType.NVarChar, 250).Value = (object) Form1.OZID;
                sqlCommand5.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
                sqlCommand5.Parameters.Add("@TotalN", SqlDbType.NVarChar, 50).Value = (object) Form1.TotalN;
                sqlCommand5.Parameters.Add("@KPI0", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI0;
                sqlCommand5.Parameters.Add("@KPI1", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI1;
                sqlCommand5.Parameters.Add("@KPI2", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI2;
                sqlCommand5.Parameters.Add("@KPI3", SqlDbType.NVarChar, 50).Value = (object) Form1.KPI3;
                sqlCommand5.Parameters.Add("@Additional_note", SqlDbType.NVarChar, 50).Value = (object) Form1.Additional_note;
                sqlCommand5.Parameters.Add("@FitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.FitStatistically;
                sqlCommand5.Parameters.Add("@RelevantForDiscussion", SqlDbType.NVarChar, 250).Value = (object) Form1.RelevantForDiscussion;
                sqlCommand5.Parameters.Add("@GraphID", SqlDbType.NVarChar, 50).Value = (object) Form1.GraphID;
                sqlCommand5.ExecuteNonQuery();
              }
            }
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            Form1.Note = this.txtNote1.Text;
            Form1.Active = !this.chActivate.Checked ? 0 : 1;
            Form1.CalcID = this.cmbCalcID2.Text;
            SqlCommand sqlCommand6 = new SqlCommand("update CalcResultView " + " set Note='" + Form1.Note + "'," + " Active  ='" + Form1.Active.ToString() + "' where calcid='" + Form1.CalcID + "'", connection);
            sqlCommand6.Parameters.Add("@Note", SqlDbType.NVarChar, 250).Value = (object) Form1.Note;
            sqlCommand6.Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
            sqlCommand6.ExecuteNonQuery();
            SqlCommand sqlCommand7 = (SqlCommand) null;
            sqlCommand7 = (SqlCommand) null;
            Form1.CalcID = this.cmbCalcID2.Text;
            Form1.Note = this.txtNote.Text;
            if (this.chActivate.Checked)
              Form1.Active = 1;
            if (this.chDeActivate.Checked)
              Form1.Active = 0;
            new SqlCommand("update CalcResultView " + " set Active  ='" + Form1.Active.ToString() + "'", connection).Parameters.Add("@Active", SqlDbType.NVarChar, 50).Value = (object) Form1.Active;
            if (Form1.strRowsCount > 0)
            {
              string str1 = "";
              for (int index = 1; index <= Form1.strRowsCount; ++index)
              {
                if (!this.rbAll.Checked)
                  ;
                Form1.OZID = this.strMatrix[index, 1];
                Form1.CalcID = this.cmbCalcID.Text;
                Form1.TotalN = this.strMatrix[index, 2];
                Form1.KPI0 = this.strMatrix[index, 3];
                Form1.KPI1 = this.strMatrix[index, 4];
                Form1.KPI2 = this.strMatrix[index, 5];
                Form1.KPI3 = this.strMatrix[index, 6];
                str1 = this.strMatrix[index, 9];
                Form1.FitStatistically = this.strMatrix[index, 7];
                Form1.RelevantForDiscussion = this.strMatrix[index, 8];
                string str2 = this.strMatrix[index, 9];
                Form1.Additional_note = this.strMatrix[index, 10];
                string cmdText = "update VIRT_OZID_per_calculation " + " set FitStatistically='" + Form1.FitStatistically + "'," + " RelevantForDiscussion='" + Form1.RelevantForDiscussion + "'," + " Additional_note='" + Form1.Additional_note + "'," + " Active ='" + str2 + "' where VIRT_OZID='" + Form1.OZID + "' and CalcID='" + this.cmbCalcID2.Text + "'";
                if (this.strMatrix[index, 1] != null)
                {
                  SqlCommand sqlCommand8 = new SqlCommand(cmdText, connection);
                  sqlCommand8.Parameters.Add("@VIRT_OZID", SqlDbType.NVarChar, 250).Value = (object) Form1.OZID;
                  sqlCommand8.Parameters.Add("@CalcID", SqlDbType.NVarChar, 50).Value = (object) Form1.CalcID;
                  sqlCommand8.Parameters.Add("@Additional_note", SqlDbType.NVarChar, 250).Value = (object) Form1.Additional_note;
                  sqlCommand8.Parameters.Add("@FitStatistically", SqlDbType.NVarChar, 50).Value = (object) Form1.FitStatistically;
                  sqlCommand8.Parameters.Add("@RelevantForDiscussion", SqlDbType.NVarChar, 250).Value = (object) Form1.RelevantForDiscussion;
                  sqlCommand8.Parameters.Add("@Active", SqlDbType.NVarChar, 250).Value = (object) str2;
                  sqlCommand8.ExecuteNonQuery();
                }
              }
              connection.Close();
            }
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      int num1 = (int) MessageBox.Show("The data was saved!");
    }

    private void clbRefCpv_SelectedIndexChanged(object sender, EventArgs e)
    {
    }

    private void clbVirtOzid_LostFocus(object sender, EventArgs e)
    {
      try
      {
        this.clbRefCpv.Items.Clear();
        string str1 = " where VIRT_OZID in ('";
        for (int index = 0; index < this.clbVirtOzid.Items.Count; ++index)
        {
          if (this.clbVirtOzid.GetItemCheckState(index) == CheckState.Checked)
            str1 = str1 + (string) this.clbVirtOzid.Items[index] + "','";
        }
        string str2 = this.Left(str1, str1.Length - 2) + ") ";
        str2.IndexOf("()");
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlDataReader sqlDataReader = new SqlCommand("SELECT distinct REFERENCED_CPV FROM [Products] " + str2, connection).ExecuteReader();
          int index = 0;
          this.clbRefCpv.Items.Clear();
          while (sqlDataReader.Read())
          {
            this.clbRefCpv.Items.Add((object) sqlDataReader["REFERENCED_CPV"].ToString());
            this.clbRefCpv.SetItemChecked(index, true);
            ++index;
          }
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void clbRefCpv_LostFocus(object sender, EventArgs e)
    {
      try
      {
        this.clbVirtOzid.Items.Clear();
        string str1 = " where REFERENCED_CPV in ('";
        this.lstCheckExclVirtOzid.Items.Clear();
        for (int index = 0; index < this.clbRefCpv.Items.Count; ++index)
        {
          if (this.clbRefCpv.GetItemCheckState(index) == CheckState.Checked)
          {
            str1 = str1 + (string) this.clbRefCpv.Items[index] + "','";
            this.lstCheckExclVirtOzid.Items.Add((object) (string) this.clbRefCpv.Items[index]);
          }
        }
        string str2 = this.Left(str1, str1.Length - 2) + ") ";
        str2.IndexOf("()");
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlDataReader sqlDataReader = new SqlCommand("SELECT distinct VIRT_OZID FROM [Products] " + str2, connection).ExecuteReader();
          int index = 0;
          this.clbVirtOzid.Items.Clear();
          while (sqlDataReader.Read())
          {
            this.clbVirtOzid.Items.Add((object) sqlDataReader["VIRT_OZID"].ToString());
            this.clbVirtOzid.SetItemChecked(index, true);
            ++index;
          }
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void timer1_Tick_1(object sender, EventArgs e)
    {
    }

    private void chActivate_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        if (this.chActivate.Checked)
        {
          this.chActivate.Enabled = false;
          this.chDeActivate.Enabled = true;
          this.chDeActivate.Checked = false;
        }
        if (!this.chDeActivate.Checked)
          return;
        this.chActivate.Enabled = true;
        this.chDeActivate.Enabled = false;
        this.chDeActivate.Checked = true;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void chDeActivate_CheckedChanged_1(object sender, EventArgs e)
    {
      try
      {
        this.chDeActivate_CheckedChanged(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void label63_Click(object sender, EventArgs e)
    {
    }

    private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
    {
    }

    private void btnCompareCalc_Click(object sender, EventArgs e)
    {
      Form1.strFilter = "";
      this.btnPrint3.Visible = false;
      this.btnPrint4.Visible = false;
      this.pictureBox3.Image = (Image) null;
      this.pictureBox4.Image = (Image) null;
      this.btnZoomOut3.Visible = false;
      this.btnZoomOut4.Visible = false;
      for (int index = 0; index < 500; ++index)
        Form1.strArrToolTip[index] = "";
      try
      {
        Form1.strQuery3 = " (1 = 1)     and 0=0";
        Form1.strQuery4 = " (1 = 1)     and 0=0";
        string str1 = " and Virt_Ozid in ('";
        string str2 = " and Virt_Ozid in ('";
        for (int index = 0; index < this.chkVirtOzid3.Items.Count; ++index)
        {
          if (this.chkVirtOzid3.GetItemCheckState(index) == CheckState.Checked)
            str1 = str1 + (string) this.chkVirtOzid3.Items[index] + "','";
        }
        string str3 = this.Left(str1, str1.Length - 2) + ") ";
        if (str3 == " and Virt_Ozid in ) ")
          str3 = "";
        if (str3.IndexOf("()") <= 0)
          Form1.strQuery3 += str3;
        for (int index = 0; index < this.chkVirtOzid4.Items.Count; ++index)
        {
          if (this.chkVirtOzid4.GetItemCheckState(index) == CheckState.Checked)
            str2 = str2 + (string) this.chkVirtOzid4.Items[index] + "','";
        }
        string str4 = this.Left(str2, str2.Length - 2) + ") ";
        if (str4 == " and Virt_Ozid in ) ")
          str4 = "";
        if (str4.IndexOf("()") <= 0)
          Form1.strQuery4 += str4;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcRow where CalcID = '" + this.cmbCalcID3.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable3 = true;
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcRow where CalcID = '" + this.cmbCalcID4.Text + "' ";
          SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
          sqlCommand3.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
          SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            Form1.IsCalcIDAvailable4 = true;
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          this.panel24.Controls.Clear();
          string[,] strArrMain1 = new string[500, 8];
          string[,] strArrMain2 = new string[500, 8];
          string[,] strArray1 = new string[500, 16];
          string cmdText1 = "  Select VIRT_OZID, dbo.KPIcount0(CalcID, VIRT_OZID), dbo.KPIcount1(CalcID, VIRT_OZID), dbo.KPIcount2(CalcID, VIRT_OZID), dbo.KPIcount3(CalcID, VIRT_OZID), FitStatistically, RelevantForDiscussion,  Additional_note from   VIRT_OZID_per_calculation where calcid = '" + this.cmbCalcID3.Text + "' and " + Form1.strQuery3 + " order by VIRT_OZID";
          string cmdText2 = "  Select VIRT_OZID, dbo.KPIcount0(CalcID, VIRT_OZID), dbo.KPIcount1(CalcID, VIRT_OZID), dbo.KPIcount2(CalcID, VIRT_OZID), dbo.KPIcount3(CalcID, VIRT_OZID), FitStatistically, RelevantForDiscussion,  Additional_note from   VIRT_OZID_per_calculation where calcid = '" + this.cmbCalcID4.Text + "' and " + Form1.strQuery4 + " order by VIRT_OZID";
          string cmdText3 = "select DISTINCT VIRT_OZID FROM [dbo].VIRT_OZID_per_calculation where calcid ='" + this.cmbCalcID3.Text + "' and " + Form1.strQuery3 + " union select DISTINCT VIRT_OZID FROM [dbo].VIRT_OZID_per_calculation where calcid ='" + this.cmbCalcID4.Text + "' and " + Form1.strQuery4;
          SqlDataReader sqlDataReader1 = new SqlCommand(cmdText1, connection).ExecuteReader();
          int index1 = 0;
          try
          {
            while (sqlDataReader1.Read())
            {
              strArrMain1[index1, 0] = sqlDataReader1[0].ToString();
              strArrMain1[index1, 1] = sqlDataReader1[1].ToString();
              strArrMain1[index1, 2] = sqlDataReader1[2].ToString();
              strArrMain1[index1, 3] = sqlDataReader1[3].ToString();
              strArrMain1[index1, 4] = sqlDataReader1[4].ToString();
              strArrMain1[index1, 5] = sqlDataReader1[5].ToString();
              strArrMain1[index1, 6] = sqlDataReader1[6].ToString();
              strArrMain1[index1, 7] = sqlDataReader1[7].ToString();
              ++index1;
            }
          }
          catch (Exception ex)
          {
            int num = (int) MessageBox.Show(ex.Message);
          }
          int k1 = index1;
          sqlDataReader1.Close();
          SqlDataReader sqlDataReader2 = new SqlCommand(cmdText2, connection).ExecuteReader();
          int index2 = 0;
          while (sqlDataReader2.Read())
          {
            strArrMain2[index2, 0] = sqlDataReader2[0].ToString();
            strArrMain2[index2, 1] = sqlDataReader2[1].ToString();
            strArrMain2[index2, 2] = sqlDataReader2[2].ToString();
            strArrMain2[index2, 3] = sqlDataReader2[3].ToString();
            strArrMain2[index2, 4] = sqlDataReader2[4].ToString();
            strArrMain2[index2, 5] = sqlDataReader2[5].ToString();
            strArrMain2[index2, 6] = sqlDataReader2[6].ToString();
            strArrMain2[index2, 7] = sqlDataReader2[7].ToString();
            ++index2;
          }
          int k2 = index2;
          sqlDataReader2.Close();
          int num1 = k1 < k2 ? k2 : k1;
          for (int index3 = 0; index3 <= k1; ++index3)
            Form1.arr1[index3] = strArrMain1[index3, 0];
          for (int index4 = 0; index4 <= k2; ++index4)
            Form1.arr2[index4] = strArrMain2[index4, 0];
          for (int index5 = 0; index5 <= k1 + k2; ++index5)
            Form1.arr3[index5] = strArrMain2[index5, 0];
          SqlDataReader sqlDataReader3 = new SqlCommand(cmdText3, connection).ExecuteReader();
          int index6 = 0;
          string[] strArray2 = new string[7];
          string[] strArray3 = new string[7];
          while (sqlDataReader3.Read())
          {
            if (this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr1) == 1 && this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr2) == 1)
            {
              string[] arrValue1 = this.FindArrValue(sqlDataReader3[0].ToString(), Form1.arr1, strArrMain1, k1);
              string[] arrValue2 = this.FindArrValue(sqlDataReader3[0].ToString(), Form1.arr2, strArrMain2, k2);
              strArray1[index6, 0] = arrValue1[0];
              strArray1[index6, 1] = arrValue1[1];
              strArray1[index6, 2] = arrValue1[2];
              strArray1[index6, 3] = arrValue1[3];
              strArray1[index6, 4] = arrValue1[4];
              strArray1[index6, 5] = arrValue1[5];
              strArray1[index6, 6] = arrValue1[6];
              strArray1[index6, 7] = arrValue1[7];
              strArray1[index6, 8] = arrValue2[0];
              strArray1[index6, 9] = arrValue2[1];
              strArray1[index6, 10] = arrValue2[2];
              strArray1[index6, 11] = arrValue2[3];
              strArray1[index6, 12] = arrValue2[4];
              strArray1[index6, 13] = arrValue2[5];
              strArray1[index6, 14] = arrValue2[6];
              strArray1[index6, 15] = arrValue2[7];
            }
            if (this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr1) == 1 && this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr2) == 0)
            {
              string[] arrValue = this.FindArrValue(sqlDataReader3[0].ToString(), Form1.arr1, strArrMain1, k1);
              strArray1[index6, 0] = arrValue[0];
              strArray1[index6, 1] = arrValue[1];
              strArray1[index6, 2] = arrValue[2];
              strArray1[index6, 3] = arrValue[3];
              strArray1[index6, 4] = arrValue[4];
              strArray1[index6, 5] = arrValue[5];
              strArray1[index6, 6] = arrValue[6];
              strArray1[index6, 7] = arrValue[7];
              strArray1[index6, 8] = "----";
              strArray1[index6, 9] = "----";
              strArray1[index6, 10] = "----";
              strArray1[index6, 11] = "----";
              strArray1[index6, 12] = "----";
              strArray1[index6, 13] = "----";
              strArray1[index6, 14] = "----";
              strArray1[index6, 15] = "----";
            }
            if (this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr1) == 0 && this.FindOzid(sqlDataReader3[0].ToString(), Form1.arr2) == 1)
            {
              string[] arrValue = this.FindArrValue(sqlDataReader3[0].ToString(), Form1.arr2, strArrMain2, k2);
              strArray1[index6, 0] = arrValue[0];
              strArray1[index6, 1] = "----";
              strArray1[index6, 2] = "----";
              strArray1[index6, 3] = "----";
              strArray1[index6, 4] = "----";
              strArray1[index6, 5] = "----";
              strArray1[index6, 6] = "----";
              strArray1[index6, 7] = "----";
              strArray1[index6, 8] = arrValue[0];
              strArray1[index6, 9] = arrValue[1];
              strArray1[index6, 10] = arrValue[2];
              strArray1[index6, 11] = arrValue[3];
              strArray1[index6, 12] = arrValue[4];
              strArray1[index6, 13] = arrValue[5];
              strArray1[index6, 14] = arrValue[6];
              strArray1[index6, 15] = arrValue[7];
            }
            ++index6;
          }
          for (int index7 = 0; index7 < num1; ++index7)
          {
            if (strArray1[index7, 0] != null)
            {
              TextBox textBox1 = new TextBox();
              textBox1.Enabled = false;
              textBox1.Width = 120;
              textBox1.Text = strArray1[index7, 0];
              textBox1.Top = index7 * 20;
              textBox1.Left = 3;
              TextBox textBox2 = textBox1;
              this.strMatrix[index7, 1] = textBox2.Text;
              Form1.strOZID = textBox2.Text;
              Form1.strArr2[index7] = textBox2.Text;
              TextBox textBox3 = new TextBox();
              textBox3.Enabled = false;
              textBox3.Width = 90;
              textBox3.Text = strArray1[index7, 1] + "," + strArray1[index7, 2] + "," + strArray1[index7, 3] + "," + strArray1[index7, 4];
              textBox3.Top = index7 * 20;
              textBox3.Left = 150;
              TextBox textBox4 = textBox3;
              this.strMatrix[index7, 2] = textBox4.Text;
              TextBox textBox5 = new TextBox();
              textBox5.Enabled = false;
              textBox5.Width = 90;
              textBox5.Text = strArray1[index7, 9] + "," + strArray1[index7, 10] + "," + strArray1[index7, 11] + "," + strArray1[index7, 12];
              textBox5.Top = index7 * 20;
              textBox5.Left = 290;
              TextBox textBox6 = textBox5;
              this.strMatrix[index7, 3] = textBox6.Text;
              TextBox textBox7 = new TextBox();
              textBox7.Enabled = false;
              textBox7.Width = 50;
              textBox7.Text = strArray1[index7, 5];
              textBox7.Top = index7 * 20;
              textBox7.Left = 440;
              TextBox textBox8 = textBox7;
              this.strMatrix[index7, 4] = textBox8.Text;
              TextBox textBox9 = new TextBox();
              textBox9.Enabled = false;
              textBox9.Width = 50;
              textBox9.Text = strArray1[index7, 13];
              textBox9.Top = index7 * 20;
              textBox9.Left = 529;
              TextBox textBox10 = textBox9;
              this.strMatrix[index7, 5] = textBox10.Text;
              if (textBox8.Text != textBox10.Text)
                textBox10.BackColor = Color.Red;
              Button button1 = new Button();
              button1.Name = "b1" + index7.ToString();
              button1.Text = string.Format("{0}", (object) "Chart");
              button1.Top = index7 * 20;
              button1.Left = 619;
              Button button2 = button1;
              this.panel24.Controls.Add((Control) button2);
              button2.Click += new EventHandler(this.ba34_Click);
              TextBox textBox11 = new TextBox();
              textBox11.Enabled = false;
              textBox11.Width = 50;
              textBox11.Text = strArray1[index7, 6];
              textBox11.Top = index7 * 20;
              textBox11.Left = 719;
              TextBox textBox12 = textBox11;
              this.strMatrix[index7, 6] = textBox12.Text;
              TextBox textBox13 = new TextBox();
              textBox13.Enabled = false;
              textBox13.Text = strArray1[index7, 14];
              textBox13.Width = 50;
              textBox13.Top = index7 * 20;
              textBox13.Left = 790;
              TextBox textBox14 = textBox13;
              this.strMatrix[index7, 7] = textBox14.Text;
              TextBox textBox15 = new TextBox();
              textBox15.Enabled = true;
              textBox15.Text = strArray1[index7, 7];
              textBox15.Width = 165;
              textBox15.Top = index7 * 20;
              textBox15.Left = 869;
              TextBox textBox16 = textBox15;
              this.strMatrix[index7, 7] = textBox14.Text;
              TextBox textBox17 = new TextBox();
              textBox17.Enabled = true;
              textBox17.Text = strArray1[index7, 15];
              textBox17.Width = 165;
              textBox17.Top = index7 * 20;
              textBox17.Left = 1079;
              TextBox textBox18 = textBox17;
              this.strMatrix[index7, 7] = textBox14.Text;
              this.panel24.Controls.Add((Control) textBox2);
              this.panel24.Controls.Add((Control) textBox4);
              this.panel24.Controls.Add((Control) textBox6);
              this.panel24.Controls.Add((Control) textBox8);
              this.panel24.Controls.Add((Control) textBox10);
              this.panel24.Controls.Add((Control) textBox12);
              this.panel24.Controls.Add((Control) textBox14);
              this.panel24.Controls.Add((Control) textBox16);
              this.panel24.Controls.Add((Control) textBox18);
              this.panel24.Controls.Add(this.ba);
            }
          }
        }
      }
      catch (Exception ex)
      {
        if (ex.Source != null)
        {
          int num = (int) MessageBox.Show("IOException source: {0}", ex.Message);
        }
      }
      int num2 = (int) MessageBox.Show("Comparing is finished!");
    }

    public int FindOzid(string Ozid, string[] arrString)
    {
      try
      {
        int ozid = 0;
        for (int index = 0; index <= arrString.Length; ++index)
        {
          if (Ozid == arrString[index] && arrString[index] != "")
          {
            ozid = 1;
            break;
          }
          if (arrString[index] == null)
            break;
        }
        if (ozid == 0)
          ozid = 0;
        return ozid;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
        return 0;
      }
    }

    public string[] FindArrValue(string Ozid, string[] arrString, string[,] strArrMain, int k)
    {
      string[] arrValue = new string[8];
      bool flag = false;
      for (int index1 = 0; index1 <= arrString.Length; ++index1)
      {
        if (Ozid == arrString[index1])
        {
          for (int index2 = 0; index2 <= k; ++index2)
          {
            if (strArrMain[index2, 0] == Ozid)
            {
              for (int index3 = 0; index3 <= 8 && index3 != 8; ++index3)
                arrValue[index3] = strArrMain[index2, index3];
              flag = true;
              break;
            }
            if (flag)
              break;
          }
          if (flag)
            break;
        }
        if (flag)
          break;
      }
      return arrValue;
    }

    private void CompareCalculation_Click(object sender, EventArgs e)
    {
    }

    private void cmbProdID3_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.cmbCalcID3.Items.Clear();
        if (this.cmbProdID3.Text == "All")
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw";
            SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
            sqlCommand1.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
            SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID3.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + this.cmbProdID3.Text + "'";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
            sqlCommand3.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID3.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbProdID4_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        this.cmbCalcID4.Items.Clear();
        if (this.cmbProdID4.Text == "All")
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw";
            SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
            sqlCommand1.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
            SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID4.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
        else
        {
          using (SqlConnection connection = new SqlConnection(Form1.connectionString))
          {
            connection.Open();
            string cmdText = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + this.cmbProdID4.Text + "'";
            SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
            sqlCommand3.ExecuteScalar();
            SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
            SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
            while (sqlDataReader.Read())
              this.cmbCalcID4.Items.Add((object) sqlDataReader[0].ToString());
            sqlDataReader.Close();
            connection.Close();
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbCalcID3_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalculationRaw where CalcID = '" + this.cmbCalcID3.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.dtCalcDateTime3.Value = System.DateTime.Parse(sqlDataReader[1].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID3.Text + "'";
          SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
          sqlCommand3.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
          SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.txtNote3.Text = sqlDataReader["Note"].ToString();
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand("select distinct active from CalcResultView where calcid='" + this.cmbCalcID3.Text + "'", connection);
          sqlCommand.CommandTimeout = 600;
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          string str = "0";
          while (sqlDataReader.Read())
            str = sqlDataReader[0].ToString();
          if (str == "False")
          {
            this.chDeActivate.Checked = true;
            this.chDeActivate.Enabled = false;
            this.chActivate.Checked = false;
            this.chActivate.Enabled = true;
          }
          else
          {
            this.chActivate.Checked = true;
            this.chActivate.Enabled = false;
            this.chDeActivate.Checked = false;
            this.chDeActivate.Enabled = true;
          }
          sqlDataReader.Close();
          connection.Close();
        }
        string str1 = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID3.Text + "'";
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID3.Text + "'";
          SqlCommand sqlCommand5 = new SqlCommand(cmdText, connection);
          sqlCommand5.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand5.ExecuteReader();
          SqlCommand sqlCommand6 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.txtNote1.Text = sqlDataReader["Note"].ToString();
          sqlDataReader.Close();
          connection.Close();
        }
        string sqlQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + this.cmbCalcID3.Text + "' order by VIRT_OZID";
        this.chkVirtOzid3.Items.Clear();
        this.SQLRunFillCheckedListBox(sqlQuery, this.chkVirtOzid3);
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalculationRaw where CalcID = '" + this.cmbCalcID3.Text + "'";
          SqlCommand sqlCommand7 = new SqlCommand(cmdText, connection);
          sqlCommand7.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand7.ExecuteReader();
          SqlCommand sqlCommand8 = new SqlCommand(cmdText, connection);
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void cmbCalcID4_SelectedIndexChanged(object sender, EventArgs e)
    {
      try
      {
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalculationRaw where CalcID = '" + this.cmbCalcID4.Text + "'";
          SqlCommand sqlCommand1 = new SqlCommand(cmdText, connection);
          sqlCommand1.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand1.ExecuteReader();
          SqlCommand sqlCommand2 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.dtCalcDateTime4.Value = System.DateTime.Parse(sqlDataReader[1].ToString());
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalcResultView where CalcID = '" + this.cmbCalcID4.Text + "'";
          SqlCommand sqlCommand3 = new SqlCommand(cmdText, connection);
          sqlCommand3.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand3.ExecuteReader();
          SqlCommand sqlCommand4 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.txtNote4.Text = sqlDataReader["Note"].ToString();
          sqlDataReader.Close();
          connection.Close();
        }
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          SqlCommand sqlCommand = new SqlCommand("select distinct active from CalcResultView where calcid='" + this.cmbCalcID4.Text + "'", connection);
          sqlCommand.CommandTimeout = 600;
          sqlCommand.ExecuteScalar();
          DataTable dataTable = new DataTable();
          SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
          string str = "0";
          while (sqlDataReader.Read())
            str = sqlDataReader[0].ToString();
          if (str == "False")
          {
            this.chDeActivate.Checked = true;
            this.chDeActivate.Enabled = false;
            this.chActivate.Checked = false;
            this.chActivate.Enabled = true;
          }
          else
          {
            this.chActivate.Checked = true;
            this.chActivate.Enabled = false;
            this.chDeActivate.Checked = false;
            this.chDeActivate.Enabled = true;
          }
          sqlDataReader.Close();
          connection.Close();
        }
        string str1 = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID4.Text + "'";
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select top 1  Note, Active from CalcResultView where calcid='" + this.cmbCalcID4.Text + "'";
          SqlCommand sqlCommand5 = new SqlCommand(cmdText, connection);
          sqlCommand5.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand5.ExecuteReader();
          SqlCommand sqlCommand6 = new SqlCommand(cmdText, connection);
          while (sqlDataReader.Read())
            this.txtNote1.Text = sqlDataReader["Note"].ToString();
          sqlDataReader.Close();
          connection.Close();
        }
        string sqlQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + this.cmbCalcID4.Text + "' order by VIRT_OZID";
        this.chkVirtOzid4.Items.Clear();
        this.SQLRunFillCheckedListBox(sqlQuery, this.chkVirtOzid4);
        using (SqlConnection connection = new SqlConnection(Form1.connectionString))
        {
          connection.Open();
          string cmdText = "select distinct * from CalculationRaw where CalcID = '" + this.cmbCalcID4.Text + "'";
          SqlCommand sqlCommand7 = new SqlCommand(cmdText, connection);
          sqlCommand7.ExecuteScalar();
          SqlDataReader sqlDataReader = sqlCommand7.ExecuteReader();
          SqlCommand sqlCommand8 = new SqlCommand(cmdText, connection);
          sqlDataReader.Close();
          connection.Close();
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void m_Picturebox_Canvas_Paint(object sender, PaintEventArgs e) => e.Graphics.Transform = this.transform;

    private void ZoomScroll(MouseEventArgs e)
    {
      if (e.Delta == 0)
        return;
      if (e.Delta <= 0)
      {
        if (this.picGraph.Width < 50)
          ;
      }
      else
      {
        this.picGraph.Width += Convert.ToInt32(this.picGraph.Width * e.Delta / 1000);
        this.picGraph.Height += Convert.ToInt32(this.picGraph.Height * e.Delta / 1000);
        this.picGraph.Refresh();
      }
    }

    private void btnZoomOut4_Click(object sender, MouseEventArgs ea)
    {
      if (this.pictureBox4.Image == null)
        return;
      if (ea.Delta > 0)
      {
        if (this.pictureBox4.Width < 15 * this.Width && this.pictureBox4.Height < 15 * this.Height)
        {
          this.pictureBox4.Width = (int) ((double) this.pictureBox4.Width * 1.25);
          this.pictureBox4.Height = (int) ((double) this.pictureBox4.Height * 1.25);
          this.pictureBox4.Top = (int) ((double) ea.Y - 1.25 * (double) (ea.Y - this.pictureBox4.Top));
          this.pictureBox4.Left = (int) ((double) ea.X - 1.25 * (double) (ea.X - this.pictureBox4.Left));
        }
      }
      else if (this.pictureBox4.Width > 100 && this.pictureBox4.Height > 100)
      {
        this.pictureBox4.Width = (int) ((double) this.pictureBox4.Width / 1.25);
        this.pictureBox4.Height = (int) ((double) this.pictureBox4.Height / 1.25);
        this.pictureBox4.Top = (int) ((double) ea.Y - 0.8 * (double) (ea.Y - this.pictureBox4.Top));
        this.pictureBox4.Left = (int) ((double) ea.X - 0.8 * (double) (ea.X - this.pictureBox4.Left));
      }
    }

    private void btnZoomOut4_Click(object sender, EventArgs e)
    {
      try
      {
        SqlConnection sqlConnection = new SqlConnection(Form1.connectionString);
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID4.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count <= 0)
            return;
          byte[] buffer = (byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"];
          MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
          PictureBox pictureBox = new PictureBox();
          pictureBox.Image = Image.FromStream((Stream) memoryStream);
          pictureBox.Location = new Point(3, 3);
          pictureBox.Size = new Size(1100, 900);
          pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
          Form form = new Form();
          form.Size = new Size(1200, 1000);
          form.Controls.Add((Control) pictureBox);
          int num = (int) form.ShowDialog();
          connection.Close();
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnZoom1_Click(object sender, EventArgs e)
    {
      try
      {
        this.ba_Zoom(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnZoomIn1_Click(object sender, EventArgs e)
    {
      try
      {
        SqlConnection sqlConnection = new SqlConnection(Form1.connectionString);
        if (Form1.blnPressed)
          Form1.strOZID = this.lbOzid.SelectedItem.ToString();
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID2.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count <= 0)
            return;
          MemoryStream memoryStream = new MemoryStream((byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"]);
          PictureBox pictureBox = new PictureBox();
          pictureBox.Image = Image.FromStream((Stream) memoryStream);
          pictureBox.Location = new Point(3, 3);
          pictureBox.Size = new Size(1100, 900);
          pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
          Form form = new Form();
          form.Size = new Size(1200, 1000);
          form.Controls.Add((Control) pictureBox);
          int num = (int) form.ShowDialog();
          connection.Close();
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnZoomOut3_Click(object sender, EventArgs e)
    {
      try
      {
        SqlConnection sqlConnection = new SqlConnection(Form1.connectionString);
        try
        {
          this.picGraph.Visible = true;
          SqlConnection connection = new SqlConnection(Form1.connectionString);
          connection.Open();
          SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + this.cmbCalcID3.Text + "' and VIRT_OZID='" + Form1.strOZID + "'", connection));
          DataSet dataSet = new DataSet();
          sqlDataAdapter.Fill(dataSet, "Graphs");
          int count = dataSet.Tables["Graphs"].Rows.Count;
          if (count <= 0)
            return;
          byte[] buffer = (byte[]) dataSet.Tables["Graphs"].Rows[count - 1]["ImageValue"];
          MemoryStream memoryStream = new MemoryStream(buffer, 0, buffer.Length);
          PictureBox pictureBox = new PictureBox();
          pictureBox.Image = Image.FromStream((Stream) memoryStream);
          pictureBox.Location = new Point(3, 3);
          pictureBox.Size = new Size(1100, 900);
          pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
          Form form = new Form();
          form.Size = new Size(1200, 1000);
          form.Controls.Add((Control) pictureBox);
          int num = (int) form.ShowDialog();
          connection.Close();
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show(ex.ToString());
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnPrint_Click(object sender, EventArgs e)
    {
      this.printDocument1.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
      this.printDialog1.AllowSomePages = true;
      this.printDialog1.ShowHelp = true;
      this.printDialog1.Document = this.docToPrint;
      if (this.printDialog1.ShowDialog() != DialogResult.OK)
        return;
      this.printDocument1.DefaultPageSettings.Landscape = true;
      this.printDocument1.Print();
    }

    private void printDocument1_PrintPage(object sender, PrintPageEventArgs e) => e.Graphics.DrawImage(this.pictureBox2.Image, 0, 0);

    private void printDocument2_PrintPage(object sender, PrintPageEventArgs e) => e.Graphics.DrawImage(this.picGraph.Image, 0, 0);

    private void btnPrint3_Click(object sender, EventArgs e)
    {
      this.printDocument3.PrintPage += new PrintPageEventHandler(this.printDocument3_PrintPage);
      this.printDialog1.AllowSomePages = true;
      this.printDialog1.ShowHelp = true;
      this.printDialog1.Document = this.docToPrint;
      if (this.printDialog1.ShowDialog() != DialogResult.OK)
        return;
      this.printDocument3.DefaultPageSettings.Landscape = true;
      this.printDocument3.Print();
    }

    private void printDocument3_PrintPage(object sender, PrintPageEventArgs e)
    {
      try
      {
        e.Graphics.DrawImage(this.pictureBox3.Image, 0, 0);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void document_PrintPage(object sender, PrintPageEventArgs e)
    {
      string s = "In document_PrintPage method.";
      Font font = new Font("Arial", 35f, FontStyle.Regular);
      e.Graphics.DrawString(s, font, Brushes.Black, 10f, 10f);
    }

    private void btnPrint2_Click(object sender, EventArgs e)
    {
      this.printDocument2.PrintPage += new PrintPageEventHandler(this.printDocument2_PrintPage);
      this.printDialog1.AllowSomePages = true;
      this.printDialog1.PrinterSettings.DefaultPageSettings.Landscape = true;
      this.printDialog1.ShowHelp = true;
      this.printDialog1.Document = this.docToPrint;
      if (this.printDialog1.ShowDialog() != DialogResult.OK)
        return;
      this.printDocument2.DefaultPageSettings.Landscape = true;
      this.printDocument2.Print();
    }

    public static void PrintToASpecificPrinter(object docToPrint)
    {
      object obj = docToPrint;
      using (PrintDialog printDialog = new PrintDialog())
      {
        printDialog.AllowSomePages = true;
        printDialog.AllowSelection = true;
        if (printDialog.ShowDialog() != DialogResult.OK)
          return;
        Process.Start(new ProcessStartInfo()
        {
          CreateNoWindow = true,
          UseShellExecute = true,
          Verb = "printTo",
          Arguments = "\"" + printDialog.PrinterSettings.PrinterName + "\"",
          WindowStyle = ProcessWindowStyle.Hidden,
          FileName = obj.ToString()
        });
      }
    }

    private void printDocument4_PrintPage(object sender, PrintPageEventArgs e)
    {
      try
      {
        e.Graphics.DrawImage(this.pictureBox4.Image, 0, 0);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnPrint4_Click(object sender, EventArgs e)
    {
      this.printDocument4.PrintPage += new PrintPageEventHandler(this.printDocument4_PrintPage);
      this.printDialog1.AllowSomePages = true;
      this.printDialog1.ShowHelp = true;
      this.printDialog1.Document = this.docToPrint;
      if (this.printDialog1.ShowDialog() != DialogResult.OK)
        return;
      this.printDocument4.DefaultPageSettings.Landscape = true;
      this.printDocument4.Print();
    }

    private void btnPrint3_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.btnPrint3_Click(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void btnPrint4_Click_1(object sender, EventArgs e)
    {
      try
      {
        this.btnPrint4_Click(sender, e);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.ToString());
      }
    }

    private void dataGridViewRaw_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {
    }

    private void Main_Click(object sender, EventArgs e)
    {
    }

    private void button4_Click(object sender, EventArgs e)
    {
      for (int index = 0; index < this.chkVirtOzid3.Items.Count; ++index)
        this.chkVirtOzid3.SetItemChecked(index, false);
    }

    private void button5_Click(object sender, EventArgs e)
    {
      for (int index = 0; index < this.chkVirtOzid4.Items.Count; ++index)
        this.chkVirtOzid4.SetItemChecked(index, false);
    }

    private void Main_Click_1(object sender, EventArgs e)
    {
    }

    private void button2_Click_1(object sender, EventArgs e) => this.tabMain.SelectTab("CompareCalculation");

    private void label23_Click(object sender, EventArgs e)
    {
    }

    private void button6_Click(object sender, EventArgs e)
    {
      for (int index = 0; index < this.chkVirtOzid3.Items.Count; ++index)
        this.chkVirtOzid3.SetItemChecked(index, true);
    }

    private void button7_Click(object sender, EventArgs e)
    {
      for (int index = 0; index < this.chkVirtOzid4.Items.Count; ++index)
        this.chkVirtOzid4.SetItemChecked(index, true);
    }

    protected override void Dispose(bool disposing)
    {
      if (disposing && this.components != null)
        this.components.Dispose();
      base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
      this.components = (IContainer) new System.ComponentModel.Container();
      this.tabMain = new TabControl();
      this.Main = new TabPage();
      this.txProgressBar = new TextBox();
      this.cmbLaufnrMax = new ComboBox();
      this.cmbLaufnrMin = new ComboBox();
      this.cmbProductCode = new ComboBox();
      this.chkWithoutCL = new CheckBox();
      this.label5 = new Label();
      this.chLastN = new CheckBox();
      this.label4 = new Label();
      this.chLastM = new CheckBox();
      this.label3 = new Label();
      this.chkSortDate = new CheckBox();
      this.label2 = new Label();
      this.chkAllVirtOzid = new CheckBox();
      this.chkAllRefCpv = new CheckBox();
      this.clbVirtOzid = new CheckedListBox();
      this.clbRefCpv = new CheckedListBox();
      this.textBox1 = new TextBox();
      this.btnCalculationView = new Button();
      this.pictureBox1 = new PictureBox();
      this.btnCalculationSearch = new Button();
      this.listBox2 = new ListBox();
      this.dataGridViewTemp = new DataGridView();
      this.txtTestSQL = new TextBox();
      this.btnOpenFile = new Button();
      this.rbYears = new RadioButton();
      this.rbMonths = new RadioButton();
      this.rbWeeks = new RadioButton();
      this.rbDays = new RadioButton();
      this.button1 = new Button();
      this.listBox1 = new ListBox();
      this.txtResult = new TextBox();
      this.label1 = new Label();
      this.lstCheckExclVirtOzid = new CheckedListBox();
      this.label58 = new Label();
      this.dataGridViewRaw = new DataGridView();
      this.btnCalculation = new Button();
      this.lblDataGridTitle = new Label();
      this.chkExclVirtOzid = new CheckBox();
      this.lblActiveEvo = new Label();
      this.lblListCheckExclVirtOzid = new Label();
      this.txtLastMdataPoints = new TextBox();
      this.lblLastMdataPoints = new Label();
      this.lblLastNdataPoints = new Label();
      this.btnReset = new Button();
      this.btnFilterData = new Button();
      this.chkLaufNr = new CheckBox();
      this.lblActiveLN = new Label();
      this.dtSortDateTo = new DateTimePicker();
      this.dtSortDateFrom = new DateTimePicker();
      this.label6 = new Label();
      this.lblLaufNRfrom = new Label();
      this.lblLaufNRto = new Label();
      this.lblSortDate = new Label();
      this.lblVirtOzid = new Label();
      this.lblRferencedCPV = new Label();
      this.lblProductCode = new Label();
      this.txtLastNdataPoints = new TextBox();
      this.ViewCalculation = new TabPage();
      this.label87 = new Label();
      this.btnPrint2 = new Button();
      this.btnZoom1 = new Button();
      this.panel2 = new Panel();
      this.label21 = new Label();
      this.label20 = new Label();
      this.label19 = new Label();
      this.label18 = new Label();
      this.label17 = new Label();
      this.label12 = new Label();
      this.label13 = new Label();
      this.label14 = new Label();
      this.label15 = new Label();
      this.label16 = new Label();
      this.dataGridView2 = new DataGridView();
      this.cmbCalcID = new ComboBox();
      this.picGraph = new PictureBox();
      this.panel1 = new Panel();
      this.btnSave = new Button();
      this.chkActive = new CheckBox();
      this.rbAll = new RadioButton();
      this.rbFitStat = new RadioButton();
      this.rbNotFitStat = new RadioButton();
      this.label11 = new Label();
      this.label10 = new Label();
      this.txtNote = new TextBox();
      this.txtUser = new TextBox();
      this.txtTimePointData = new TextBox();
      this.txtTimePointCalc = new TextBox();
      this.txtNParameterTotal = new TextBox();
      this.txtNStatistically = new TextBox();
      this.txtPercentStatistically = new TextBox();
      this.txtDoNotFitStatistically = new TextBox();
      this.label7 = new Label();
      this.label8 = new Label();
      this.label22 = new Label();
      this.label23 = new Label();
      this.label25 = new Label();
      this.label26 = new Label();
      this.label28 = new Label();
      this.label29 = new Label();
      this.label30 = new Label();
      this.dataGridView1 = new DataGridView();
      this.label9 = new Label();
      this.label24 = new Label();
      this.label27 = new Label();
      this.SearchCalculation = new TabPage();
      this.label83 = new Label();
      this.label82 = new Label();
      this.btnPrint = new Button();
      this.btnZoomIn1 = new Button();
      this.pictureBox2 = new PictureBox();
      this.panel4 = new Panel();
      this.label48 = new Label();
      this.panel6 = new Panel();
      this.panel17 = new Panel();
      this.label31 = new Label();
      this.panel16 = new Panel();
      this.label32 = new Label();
      this.panel15 = new Panel();
      this.label33 = new Label();
      this.label34 = new Label();
      this.panel13 = new Panel();
      this.label35 = new Label();
      this.panel12 = new Panel();
      this.label36 = new Label();
      this.panel11 = new Panel();
      this.label37 = new Label();
      this.panel10 = new Panel();
      this.label38 = new Label();
      this.panel5 = new Panel();
      this.label47 = new Label();
      this.panelSelection = new Panel();
      this.label81 = new Label();
      this.label78 = new Label();
      this.label77 = new Label();
      this.label44 = new Label();
      this.lbGraph = new ListBox();
      this.rbKPI3 = new RadioButton();
      this.rbKPI2 = new RadioButton();
      this.rbKPI1 = new RadioButton();
      this.rbKPI0 = new RadioButton();
      this.label45 = new Label();
      this.lbOzid = new ListBox();
      this.label46 = new Label();
      this.panel3 = new Panel();
      this.label74 = new Label();
      this.label75 = new Label();
      this.groupBox1 = new GroupBox();
      this.label79 = new Label();
      this.chkVirtOzid2 = new CheckedListBox();
      this.dtCalcDateTime = new DateTimePicker();
      this.cmbProdID2 = new ComboBox();
      this.rbNotActive1 = new RadioButton();
      this.rbActive1 = new RadioButton();
      this.cmbCalcID2 = new ComboBox();
      this.rbAll1 = new RadioButton();
      this.txtNote1 = new TextBox();
      this.label39 = new Label();
      this.label40 = new Label();
      this.label41 = new Label();
      this.label42 = new Label();
      this.groupFilterSelection = new GroupBox();
      this.label80 = new Label();
      this.label76 = new Label();
      this.btnGetHistoric = new Button();
      this.rbNotFitStatF = new RadioButton();
      this.rbFitstatF = new RadioButton();
      this.rbAllF = new RadioButton();
      this.panelButtons = new Panel();
      this.label86 = new Label();
      this.label85 = new Label();
      this.label84 = new Label();
      this.chActivate = new CheckBox();
      this.chDeActivate = new CheckBox();
      this.button2 = new Button();
      this.button3 = new Button();
      this.CompareCalculation = new TabPage();
      this.button7 = new Button();
      this.button6 = new Button();
      this.label71 = new Label();
      this.button5 = new Button();
      this.button4 = new Button();
      this.btnPrint4 = new Button();
      this.btnPrint3 = new Button();
      this.btnZoomOut3 = new Button();
      this.btnZoomOut4 = new Button();
      this.btnCompareCalc = new Button();
      this.pictureBox4 = new PictureBox();
      this.pictureBox3 = new PictureBox();
      this.panel24 = new Panel();
      this.label68 = new Label();
      this.label67 = new Label();
      this.panel23 = new Panel();
      this.label66 = new Label();
      this.panel22 = new Panel();
      this.label65 = new Label();
      this.panel21 = new Panel();
      this.label64 = new Label();
      this.panel20 = new Panel();
      this.label63 = new Label();
      this.panel19 = new Panel();
      this.label62 = new Label();
      this.panel18 = new Panel();
      this.label61 = new Label();
      this.panel14 = new Panel();
      this.label60 = new Label();
      this.panel9 = new Panel();
      this.label59 = new Label();
      this.panel7 = new Panel();
      this.label56 = new Label();
      this.panel8 = new Panel();
      this.label57 = new Label();
      this.groupBox3 = new GroupBox();
      this.label72 = new Label();
      this.label73 = new Label();
      this.chkVirtOzid4 = new CheckedListBox();
      this.dtCalcDateTime4 = new DateTimePicker();
      this.cmbProdID4 = new ComboBox();
      this.cmbCalcID4 = new ComboBox();
      this.txtNote4 = new TextBox();
      this.label52 = new Label();
      this.label53 = new Label();
      this.label54 = new Label();
      this.label55 = new Label();
      this.groupBox2 = new GroupBox();
      this.chkVirtOzid3 = new CheckedListBox();
      this.label70 = new Label();
      this.label69 = new Label();
      this.dtCalcDateTime3 = new DateTimePicker();
      this.cmbProdID3 = new ComboBox();
      this.cmbCalcID3 = new ComboBox();
      this.txtNote3 = new TextBox();
      this.label43 = new Label();
      this.label49 = new Label();
      this.label50 = new Label();
      this.label51 = new Label();
      this.openFileDialog1 = new OpenFileDialog();
      this.backgroundWorker2 = new BackgroundWorker();
      this.timer1 = new Timer(this.components);
      this.printDocument1 = new PrintDocument();
      this.printDocument2 = new PrintDocument();
      this.printDocument3 = new PrintDocument();
      this.printDocument4 = new PrintDocument();
      this.printDialog1 = new PrintDialog();
      this.tabMain.SuspendLayout();
      this.Main.SuspendLayout();
      ((ISupportInitialize) this.pictureBox1).BeginInit();
      ((ISupportInitialize) this.dataGridViewTemp).BeginInit();
      ((ISupportInitialize) this.dataGridViewRaw).BeginInit();
      this.ViewCalculation.SuspendLayout();
      this.panel2.SuspendLayout();
      ((ISupportInitialize) this.dataGridView2).BeginInit();
      ((ISupportInitialize) this.picGraph).BeginInit();
      ((ISupportInitialize) this.dataGridView1).BeginInit();
      this.SearchCalculation.SuspendLayout();
      ((ISupportInitialize) this.pictureBox2).BeginInit();
      this.panel4.SuspendLayout();
      this.panel17.SuspendLayout();
      this.panel16.SuspendLayout();
      this.panel15.SuspendLayout();
      this.panel13.SuspendLayout();
      this.panel12.SuspendLayout();
      this.panel11.SuspendLayout();
      this.panel10.SuspendLayout();
      this.panel5.SuspendLayout();
      this.panelSelection.SuspendLayout();
      this.panel3.SuspendLayout();
      this.groupBox1.SuspendLayout();
      this.groupFilterSelection.SuspendLayout();
      this.panelButtons.SuspendLayout();
      this.CompareCalculation.SuspendLayout();
      ((ISupportInitialize) this.pictureBox4).BeginInit();
      ((ISupportInitialize) this.pictureBox3).BeginInit();
      this.panel23.SuspendLayout();
      this.panel22.SuspendLayout();
      this.panel21.SuspendLayout();
      this.panel20.SuspendLayout();
      this.panel19.SuspendLayout();
      this.panel18.SuspendLayout();
      this.panel14.SuspendLayout();
      this.panel9.SuspendLayout();
      this.panel7.SuspendLayout();
      this.panel8.SuspendLayout();
      this.groupBox3.SuspendLayout();
      this.groupBox2.SuspendLayout();
      this.SuspendLayout();
      this.tabMain.Controls.Add((Control) this.Main);
      this.tabMain.Controls.Add((Control) this.ViewCalculation);
      this.tabMain.Controls.Add((Control) this.SearchCalculation);
      this.tabMain.Controls.Add((Control) this.CompareCalculation);
      this.tabMain.Location = new Point(0, 0);
      this.tabMain.Name = "tabMain";
      this.tabMain.SelectedIndex = 0;
      this.tabMain.Size = new Size(1418, 1039);
      this.tabMain.TabIndex = 0;
      this.Main.BackColor = Color.White;
      this.Main.Controls.Add((Control) this.txProgressBar);
      this.Main.Controls.Add((Control) this.cmbLaufnrMax);
      this.Main.Controls.Add((Control) this.cmbLaufnrMin);
      this.Main.Controls.Add((Control) this.cmbProductCode);
      this.Main.Controls.Add((Control) this.chkWithoutCL);
      this.Main.Controls.Add((Control) this.label5);
      this.Main.Controls.Add((Control) this.chLastN);
      this.Main.Controls.Add((Control) this.label4);
      this.Main.Controls.Add((Control) this.chLastM);
      this.Main.Controls.Add((Control) this.label3);
      this.Main.Controls.Add((Control) this.chkSortDate);
      this.Main.Controls.Add((Control) this.label2);
      this.Main.Controls.Add((Control) this.chkAllVirtOzid);
      this.Main.Controls.Add((Control) this.chkAllRefCpv);
      this.Main.Controls.Add((Control) this.clbVirtOzid);
      this.Main.Controls.Add((Control) this.clbRefCpv);
      this.Main.Controls.Add((Control) this.textBox1);
      this.Main.Controls.Add((Control) this.btnCalculationView);
      this.Main.Controls.Add((Control) this.pictureBox1);
      this.Main.Controls.Add((Control) this.btnCalculationSearch);
      this.Main.Controls.Add((Control) this.listBox2);
      this.Main.Controls.Add((Control) this.dataGridViewTemp);
      this.Main.Controls.Add((Control) this.txtTestSQL);
      this.Main.Controls.Add((Control) this.btnOpenFile);
      this.Main.Controls.Add((Control) this.rbYears);
      this.Main.Controls.Add((Control) this.rbMonths);
      this.Main.Controls.Add((Control) this.rbWeeks);
      this.Main.Controls.Add((Control) this.rbDays);
      this.Main.Controls.Add((Control) this.button1);
      this.Main.Controls.Add((Control) this.listBox1);
      this.Main.Controls.Add((Control) this.txtResult);
      this.Main.Controls.Add((Control) this.label1);
      this.Main.Controls.Add((Control) this.lstCheckExclVirtOzid);
      this.Main.Controls.Add((Control) this.label58);
      this.Main.Controls.Add((Control) this.dataGridViewRaw);
      this.Main.Controls.Add((Control) this.btnCalculation);
      this.Main.Controls.Add((Control) this.lblDataGridTitle);
      this.Main.Controls.Add((Control) this.chkExclVirtOzid);
      this.Main.Controls.Add((Control) this.lblActiveEvo);
      this.Main.Controls.Add((Control) this.lblListCheckExclVirtOzid);
      this.Main.Controls.Add((Control) this.txtLastMdataPoints);
      this.Main.Controls.Add((Control) this.lblLastMdataPoints);
      this.Main.Controls.Add((Control) this.lblLastNdataPoints);
      this.Main.Controls.Add((Control) this.btnReset);
      this.Main.Controls.Add((Control) this.btnFilterData);
      this.Main.Controls.Add((Control) this.chkLaufNr);
      this.Main.Controls.Add((Control) this.lblActiveLN);
      this.Main.Controls.Add((Control) this.dtSortDateTo);
      this.Main.Controls.Add((Control) this.dtSortDateFrom);
      this.Main.Controls.Add((Control) this.label6);
      this.Main.Controls.Add((Control) this.lblLaufNRfrom);
      this.Main.Controls.Add((Control) this.lblLaufNRto);
      this.Main.Controls.Add((Control) this.lblSortDate);
      this.Main.Controls.Add((Control) this.lblVirtOzid);
      this.Main.Controls.Add((Control) this.lblRferencedCPV);
      this.Main.Controls.Add((Control) this.lblProductCode);
      this.Main.Controls.Add((Control) this.txtLastNdataPoints);
      this.Main.Location = new Point(4, 22);
      this.Main.Name = "Main";
      this.Main.Padding = new Padding(3);
      this.Main.Size = new Size(1410, 1013);
      this.Main.TabIndex = 0;
      this.Main.Tag = (object) "789";
      this.Main.Text = "Main";
      this.Main.Click += new EventHandler(this.Main_Click_1);
      this.txProgressBar.ForeColor = Color.Lime;
      this.txProgressBar.Location = new Point(24, 556);
      this.txProgressBar.Name = "txProgressBar";
      this.txProgressBar.Size = new Size(1282, 20);
      this.txProgressBar.TabIndex = 221;
      this.txProgressBar.Visible = false;
      this.cmbLaufnrMax.FormattingEnabled = true;
      this.cmbLaufnrMax.Location = new Point(752, 151);
      this.cmbLaufnrMax.Name = "cmbLaufnrMax";
      this.cmbLaufnrMax.Size = new Size(194, 21);
      this.cmbLaufnrMax.TabIndex = 220;
      this.cmbLaufnrMin.FormattingEnabled = true;
      this.cmbLaufnrMin.Location = new Point(515, 151);
      this.cmbLaufnrMin.Name = "cmbLaufnrMin";
      this.cmbLaufnrMin.Size = new Size(194, 21);
      this.cmbLaufnrMin.TabIndex = 219;
      this.cmbProductCode.FormattingEnabled = true;
      this.cmbProductCode.Location = new Point(136, 16);
      this.cmbProductCode.Name = "cmbProductCode";
      this.cmbProductCode.Size = new Size(162, 21);
      this.cmbProductCode.TabIndex = 218;
      this.cmbProductCode.SelectedIndexChanged += new EventHandler(this.cmbProductCode_SelectedIndexChanged_1);
      this.chkWithoutCL.AutoSize = true;
      this.chkWithoutCL.Location = new Point(914, 435);
      this.chkWithoutCL.Name = "chkWithoutCL";
      this.chkWithoutCL.Size = new Size(56, 17);
      this.chkWithoutCL.TabIndex = 217;
      this.chkWithoutCL.Text = "Active";
      this.chkWithoutCL.UseVisualStyleBackColor = true;
      this.chkWithoutCL.Visible = false;
      this.label5.AutoSize = true;
      this.label5.Location = new Point(756, 436);
      this.label5.Name = "label5";
      this.label5.Size = new Size(151, 13);
      this.label5.TabIndex = 216;
      this.label5.Text = "Filtering parameters without CL";
      this.label5.Visible = false;
      this.chLastN.AutoSize = true;
      this.chLastN.Location = new Point(719, 187);
      this.chLastN.Name = "chLastN";
      this.chLastN.Size = new Size(15, 14);
      this.chLastN.TabIndex = 215;
      this.chLastN.UseVisualStyleBackColor = true;
      this.chLastN.CheckedChanged += new EventHandler(this.chLastN_CheckedChanged_1);
      this.label4.AutoSize = true;
      this.label4.Location = new Point(740, 188);
      this.label4.Name = "label4";
      this.label4.Size = new Size(37, 13);
      this.label4.TabIndex = 214;
      this.label4.Text = "Active";
      this.chLastM.AutoSize = true;
      this.chLastM.Location = new Point(719, 224);
      this.chLastM.Name = "chLastM";
      this.chLastM.Size = new Size(15, 14);
      this.chLastM.TabIndex = 213;
      this.chLastM.UseVisualStyleBackColor = true;
      this.chLastM.CheckedChanged += new EventHandler(this.chLastM_CheckedChanged_1);
      this.label3.AutoSize = true;
      this.label3.Location = new Point(740, 225);
      this.label3.Name = "label3";
      this.label3.Size = new Size(37, 13);
      this.label3.TabIndex = 212;
      this.label3.Text = "Active";
      this.chkSortDate.AutoSize = true;
      this.chkSortDate.Location = new Point(719, 130);
      this.chkSortDate.Name = "chkSortDate";
      this.chkSortDate.Size = new Size(15, 14);
      this.chkSortDate.TabIndex = 211;
      this.chkSortDate.UseVisualStyleBackColor = true;
      this.chkSortDate.CheckedChanged += new EventHandler(this.chkSortDate_CheckedChanged_1);
      this.label2.AutoSize = true;
      this.label2.Location = new Point(740, 131);
      this.label2.Name = "label2";
      this.label2.Size = new Size(37, 13);
      this.label2.TabIndex = 210;
      this.label2.Text = "Active";
      this.chkAllVirtOzid.AutoSize = true;
      this.chkAllVirtOzid.Checked = true;
      this.chkAllVirtOzid.CheckState = CheckState.Checked;
      this.chkAllVirtOzid.Location = new Point(304, 181);
      this.chkAllVirtOzid.Name = "chkAllVirtOzid";
      this.chkAllVirtOzid.Size = new Size(37, 17);
      this.chkAllVirtOzid.TabIndex = 209;
      this.chkAllVirtOzid.Text = "All";
      this.chkAllVirtOzid.UseVisualStyleBackColor = true;
      this.chkAllVirtOzid.CheckedChanged += new EventHandler(this.chkAllVirtOzid_CheckedChanged_1);
      this.chkAllRefCpv.AutoSize = true;
      this.chkAllRefCpv.Checked = true;
      this.chkAllRefCpv.CheckState = CheckState.Checked;
      this.chkAllRefCpv.Location = new Point(304, 49);
      this.chkAllRefCpv.Name = "chkAllRefCpv";
      this.chkAllRefCpv.Size = new Size(37, 17);
      this.chkAllRefCpv.TabIndex = 208;
      this.chkAllRefCpv.Text = "All";
      this.chkAllRefCpv.UseVisualStyleBackColor = true;
      this.chkAllRefCpv.CheckedChanged += new EventHandler(this.chkAllRefCpv_CheckedChanged_1);
      this.clbVirtOzid.CheckOnClick = true;
      this.clbVirtOzid.FormattingEnabled = true;
      this.clbVirtOzid.Location = new Point(136, 179);
      this.clbVirtOzid.Name = "clbVirtOzid";
      this.clbVirtOzid.Size = new Size(162, 124);
      this.clbVirtOzid.TabIndex = 207;
      this.clbVirtOzid.LostFocus += new EventHandler(this.clbVirtOzid_LostFocus);
      this.clbRefCpv.CheckOnClick = true;
      this.clbRefCpv.FormattingEnabled = true;
      this.clbRefCpv.Location = new Point(136, 47);
      this.clbRefCpv.Name = "clbRefCpv";
      this.clbRefCpv.Size = new Size(162, 124);
      this.clbRefCpv.TabIndex = 206;
      this.clbRefCpv.SelectedIndexChanged += new EventHandler(this.clbRefCpv_SelectedIndexChanged);
      this.clbRefCpv.LostFocus += new EventHandler(this.clbRefCpv_LostFocus);
      this.textBox1.Location = new Point(387, 469);
      this.textBox1.Name = "textBox1";
      this.textBox1.Size = new Size(100, 20);
      this.textBox1.TabIndex = 205;
      this.textBox1.Visible = false;
      this.btnCalculationView.Enabled = false;
      this.btnCalculationView.Location = new Point(1027, 473);
      this.btnCalculationView.Name = "btnCalculationView";
      this.btnCalculationView.Size = new Size(122, 84);
      this.btnCalculationView.TabIndex = 204;
      this.btnCalculationView.Text = "5-Calculation View";
      this.btnCalculationView.UseVisualStyleBackColor = true;
      this.btnCalculationView.Click += new EventHandler(this.btnCalculationView_Click_1);
      this.pictureBox1.Location = new Point(277, 314);
      this.pictureBox1.Name = "pictureBox1";
      this.pictureBox1.Size = new Size(153, 65);
      this.pictureBox1.TabIndex = 203;
      this.pictureBox1.TabStop = false;
      this.btnCalculationSearch.Enabled = false;
      this.btnCalculationSearch.Location = new Point(1155, 473);
      this.btnCalculationSearch.Name = "btnCalculationSearch";
      this.btnCalculationSearch.Size = new Size(122, 84);
      this.btnCalculationSearch.TabIndex = 202;
      this.btnCalculationSearch.Text = "6-Calculation Search";
      this.btnCalculationSearch.UseVisualStyleBackColor = true;
      this.btnCalculationSearch.Click += new EventHandler(this.btnCalculationSearch_Click);
      this.listBox2.FormattingEnabled = true;
      this.listBox2.Location = new Point(746, 12);
      this.listBox2.Name = "listBox2";
      this.listBox2.Size = new Size(120, 43);
      this.listBox2.TabIndex = 201;
      this.listBox2.Visible = false;
      this.dataGridViewTemp.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridViewTemp.Location = new Point(11, 411);
      this.dataGridViewTemp.Name = "dataGridViewTemp";
      this.dataGridViewTemp.Size = new Size(23, 39);
      this.dataGridViewTemp.TabIndex = 200;
      this.dataGridViewTemp.Visible = false;
      this.txtTestSQL.Location = new Point(955, 181);
      this.txtTestSQL.Multiline = true;
      this.txtTestSQL.Name = "txtTestSQL";
      this.txtTestSQL.Size = new Size(243, 34);
      this.txtTestSQL.TabIndex = 199;
      this.txtTestSQL.Visible = false;
      this.btnOpenFile.BackColor = SystemColors.Menu;
      this.btnOpenFile.Location = new Point(495, 472);
      this.btnOpenFile.Name = "btnOpenFile";
      this.btnOpenFile.Size = new Size(113, 81);
      this.btnOpenFile.TabIndex = 198;
      this.btnOpenFile.Text = "1 - Select Data Raw file";
      this.btnOpenFile.UseVisualStyleBackColor = false;
      this.btnOpenFile.Click += new EventHandler(this.btnOpenFile_Click);
      this.rbYears.AutoSize = true;
      this.rbYears.Location = new Point(515, 313);
      this.rbYears.Name = "rbYears";
      this.rbYears.Size = new Size(52, 17);
      this.rbYears.TabIndex = 197;
      this.rbYears.TabStop = true;
      this.rbYears.Text = "Years";
      this.rbYears.UseVisualStyleBackColor = true;
      this.rbMonths.AutoSize = true;
      this.rbMonths.Location = new Point(515, 290);
      this.rbMonths.Name = "rbMonths";
      this.rbMonths.Size = new Size(60, 17);
      this.rbMonths.TabIndex = 196;
      this.rbMonths.TabStop = true;
      this.rbMonths.Text = "Months";
      this.rbMonths.UseVisualStyleBackColor = true;
      this.rbWeeks.AutoSize = true;
      this.rbWeeks.Location = new Point(515, 267);
      this.rbWeeks.Name = "rbWeeks";
      this.rbWeeks.Size = new Size(59, 17);
      this.rbWeeks.TabIndex = 195;
      this.rbWeeks.TabStop = true;
      this.rbWeeks.Text = "Weeks";
      this.rbWeeks.UseVisualStyleBackColor = true;
      this.rbDays.AutoSize = true;
      this.rbDays.Location = new Point(515, 244);
      this.rbDays.Name = "rbDays";
      this.rbDays.Size = new Size(49, 17);
      this.rbDays.TabIndex = 194;
      this.rbDays.TabStop = true;
      this.rbDays.Text = "Days";
      this.rbDays.UseVisualStyleBackColor = true;
      this.rbDays.CheckedChanged += new EventHandler(this.rbDays_CheckedChanged);
      this.button1.Location = new Point(743, 470);
      this.button1.Name = "button1";
      this.button1.Size = new Size(120, 84);
      this.button1.TabIndex = 193;
      this.button1.Text = "3-Export Data to Excel For Calculation";
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new EventHandler(this.button1_Click);
      this.listBox1.FormattingEnabled = true;
      this.listBox1.Location = new Point(285, 397);
      this.listBox1.Name = "listBox1";
      this.listBox1.Size = new Size(56, 95);
      this.listBox1.TabIndex = 192;
      this.listBox1.Visible = false;
      this.txtResult.Location = new Point(955, 242);
      this.txtResult.Multiline = true;
      this.txtResult.Name = "txtResult";
      this.txtResult.Size = new Size(206, 32);
      this.txtResult.TabIndex = 191;
      this.txtResult.Visible = false;
      this.label1.AutoSize = true;
      this.label1.Location = new Point(285, 314);
      this.label1.Name = "label1";
      this.label1.Size = new Size(35, 13);
      this.label1.TabIndex = 190;
      this.label1.Text = "label1";
      this.label1.TextAlign = ContentAlignment.TopCenter;
      this.label1.Visible = false;
      this.lstCheckExclVirtOzid.FormattingEnabled = true;
      this.lstCheckExclVirtOzid.Location = new Point(889, 313);
      this.lstCheckExclVirtOzid.Name = "lstCheckExclVirtOzid";
      this.lstCheckExclVirtOzid.Size = new Size(120, 109);
      this.lstCheckExclVirtOzid.TabIndex = 189;
      this.lstCheckExclVirtOzid.Visible = false;
      this.label58.AutoSize = true;
      this.label58.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label58.Location = new Point(466, 237);
      this.label58.Name = "label58";
      this.label58.Size = new Size(0, 24);
      this.label58.TabIndex = 188;
      this.dataGridViewRaw.AllowUserToOrderColumns = true;
      this.dataGridViewRaw.BorderStyle = BorderStyle.Fixed3D;
      this.dataGridViewRaw.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridViewRaw.Location = new Point(24, 556);
      this.dataGridViewRaw.Name = "dataGridViewRaw";
      this.dataGridViewRaw.Size = new Size(1282, 372);
      this.dataGridViewRaw.TabIndex = 187;
      this.dataGridViewRaw.CellContentClick += new DataGridViewCellEventHandler(this.dataGridViewRaw_CellContentClick);
      this.btnCalculation.Location = new Point(869, 470);
      this.btnCalculation.Name = "btnCalculation";
      this.btnCalculation.Size = new Size(151, 87);
      this.btnCalculation.TabIndex = 186;
      this.btnCalculation.Text = "4-Start calculation";
      this.btnCalculation.UseVisualStyleBackColor = true;
      this.btnCalculation.Click += new EventHandler(this.btnCalculation_Click_1);
      this.lblDataGridTitle.AutoSize = true;
      this.lblDataGridTitle.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.lblDataGridTitle.Location = new Point(20, 520);
      this.lblDataGridTitle.Name = "lblDataGridTitle";
      this.lblDataGridTitle.Size = new Size(448, 20);
      this.lblDataGridTitle.TabIndex = 185;
      this.lblDataGridTitle.Text = "Table of raw data entity (Original: NO  -  Filtered: Yes ) ";
      this.chkExclVirtOzid.AutoSize = true;
      this.chkExclVirtOzid.Location = new Point(363, 394);
      this.chkExclVirtOzid.Name = "chkExclVirtOzid";
      this.chkExclVirtOzid.Size = new Size(15, 14);
      this.chkExclVirtOzid.TabIndex = 184;
      this.chkExclVirtOzid.UseVisualStyleBackColor = true;
      this.chkExclVirtOzid.Visible = false;
      this.lblActiveEvo.AutoSize = true;
      this.lblActiveEvo.Location = new Point(384, 394);
      this.lblActiveEvo.Name = "lblActiveEvo";
      this.lblActiveEvo.Size = new Size(37, 13);
      this.lblActiveEvo.TabIndex = 183;
      this.lblActiveEvo.Text = "Active";
      this.lblActiveEvo.Visible = false;
      this.lblListCheckExclVirtOzid.AutoSize = true;
      this.lblListCheckExclVirtOzid.Location = new Point(753, 313);
      this.lblListCheckExclVirtOzid.Name = "lblListCheckExclVirtOzid";
      this.lblListCheckExclVirtOzid.Size = new Size(113, 13);
      this.lblListCheckExclVirtOzid.TabIndex = 182;
      this.lblListCheckExclVirtOzid.Text = "Excluding VIRT_OZID";
      this.lblListCheckExclVirtOzid.Visible = false;
      this.txtLastMdataPoints.Location = new Point(515, 218);
      this.txtLastMdataPoints.Name = "txtLastMdataPoints";
      this.txtLastMdataPoints.Size = new Size(194, 20);
      this.txtLastMdataPoints.TabIndex = 181;
      this.txtLastMdataPoints.Text = "3";
      this.lblLastMdataPoints.AutoSize = true;
      this.lblLastMdataPoints.Location = new Point(377, 225);
      this.lblLastMdataPoints.Name = "lblLastMdataPoints";
      this.lblLastMdataPoints.Size = new Size(70, 13);
      this.lblLastMdataPoints.TabIndex = 180;
      this.lblLastMdataPoints.Text = "TIMEFRAME";
      this.lblLastNdataPoints.AutoSize = true;
      this.lblLastNdataPoints.Location = new Point(377, 188);
      this.lblLastNdataPoints.Name = "lblLastNdataPoints";
      this.lblLastNdataPoints.Size = new Size(109, 13);
      this.lblLastNdataPoints.TabIndex = 179;
      this.lblLastNdataPoints.Text = "NUMBER_OF_DATA";
      this.btnReset.Location = new Point(608, 519);
      this.btnReset.Name = "btnReset";
      this.btnReset.Size = new Size(134, 33);
      this.btnReset.TabIndex = 178;
      this.btnReset.Text = "7-Reset filter";
      this.btnReset.UseVisualStyleBackColor = true;
      this.btnReset.Click += new EventHandler(this.btnReset_Click_1);
      this.btnFilterData.Location = new Point(607, 472);
      this.btnFilterData.Name = "btnFilterData";
      this.btnFilterData.Size = new Size(135, 41);
      this.btnFilterData.TabIndex = 177;
      this.btnFilterData.Text = "2-Filter data";
      this.btnFilterData.UseVisualStyleBackColor = true;
      this.btnFilterData.Click += new EventHandler(this.btnFilterData_Click);
      this.chkLaufNr.AutoSize = true;
      this.chkLaufNr.Location = new Point(955, 154);
      this.chkLaufNr.Name = "chkLaufNr";
      this.chkLaufNr.Size = new Size(15, 14);
      this.chkLaufNr.TabIndex = 176;
      this.chkLaufNr.UseVisualStyleBackColor = true;
      this.chkLaufNr.CheckedChanged += new EventHandler(this.chkLaufNr_CheckedChanged_1);
      this.lblActiveLN.AutoSize = true;
      this.lblActiveLN.Location = new Point(976, 154);
      this.lblActiveLN.Name = "lblActiveLN";
      this.lblActiveLN.Size = new Size(37, 13);
      this.lblActiveLN.TabIndex = 175;
      this.lblActiveLN.Text = "Active";
      this.dtSortDateTo.Format = DateTimePickerFormat.Short;
      this.dtSortDateTo.Location = new Point(626, 124);
      this.dtSortDateTo.Name = "dtSortDateTo";
      this.dtSortDateTo.Size = new Size(83, 20);
      this.dtSortDateTo.TabIndex = 174;
      this.dtSortDateFrom.Format = DateTimePickerFormat.Short;
      this.dtSortDateFrom.Location = new Point(515, 123);
      this.dtSortDateFrom.Name = "dtSortDateFrom";
      this.dtSortDateFrom.Size = new Size(81, 20);
      this.dtSortDateFrom.TabIndex = 173;
      this.label6.AutoSize = true;
      this.label6.Location = new Point(604, 130);
      this.label6.Name = "label6";
      this.label6.Size = new Size(16, 13);
      this.label6.TabIndex = 172;
      this.label6.Text = "to";
      this.lblLaufNRfrom.AutoSize = true;
      this.lblLaufNRfrom.Location = new Point(377, 158);
      this.lblLaufNRfrom.Name = "lblLaufNRfrom";
      this.lblLaufNRfrom.Size = new Size(73, 13);
      this.lblLaufNRfrom.TabIndex = 171;
      this.lblLaufNRfrom.Text = "LAUFNR from";
      this.lblLaufNRto.AutoSize = true;
      this.lblLaufNRto.Location = new Point(730, 154);
      this.lblLaufNRto.Name = "lblLaufNRto";
      this.lblLaufNRto.Size = new Size(16, 13);
      this.lblLaufNRto.TabIndex = 170;
      this.lblLaufNRto.Text = "to";
      this.lblSortDate.AutoSize = true;
      this.lblSortDate.Location = new Point(374, 131);
      this.lblSortDate.Name = "lblSortDate";
      this.lblSortDate.Size = new Size(95, 13);
      this.lblSortDate.TabIndex = 169;
      this.lblSortDate.Text = "SORT_DATE from";
      this.lblVirtOzid.AutoSize = true;
      this.lblVirtOzid.Location = new Point(21, 179);
      this.lblVirtOzid.Name = "lblVirtOzid";
      this.lblVirtOzid.Size = new Size(64, 13);
      this.lblVirtOzid.TabIndex = 168;
      this.lblVirtOzid.Text = "VIRT_OZID";
      this.lblRferencedCPV.AutoSize = true;
      this.lblRferencedCPV.Location = new Point(21, 49);
      this.lblRferencedCPV.Name = "lblRferencedCPV";
      this.lblRferencedCPV.Size = new Size(115, 13);
      this.lblRferencedCPV.TabIndex = 167;
      this.lblRferencedCPV.Text = "REFERRENCED_CPV";
      this.lblProductCode.AutoSize = true;
      this.lblProductCode.Location = new Point(21, 19);
      this.lblProductCode.Name = "lblProductCode";
      this.lblProductCode.Size = new Size(90, 13);
      this.lblProductCode.TabIndex = 166;
      this.lblProductCode.Text = "PRODUCTCODE";
      this.txtLastNdataPoints.Location = new Point(515, 181);
      this.txtLastNdataPoints.Name = "txtLastNdataPoints";
      this.txtLastNdataPoints.Size = new Size(194, 20);
      this.txtLastNdataPoints.TabIndex = 165;
      this.txtLastNdataPoints.Text = "5000";
      this.ViewCalculation.BackColor = Color.White;
      this.ViewCalculation.Controls.Add((Control) this.label87);
      this.ViewCalculation.Controls.Add((Control) this.btnPrint2);
      this.ViewCalculation.Controls.Add((Control) this.btnZoom1);
      this.ViewCalculation.Controls.Add((Control) this.panel2);
      this.ViewCalculation.Controls.Add((Control) this.dataGridView2);
      this.ViewCalculation.Controls.Add((Control) this.cmbCalcID);
      this.ViewCalculation.Controls.Add((Control) this.picGraph);
      this.ViewCalculation.Controls.Add((Control) this.panel1);
      this.ViewCalculation.Controls.Add((Control) this.btnSave);
      this.ViewCalculation.Controls.Add((Control) this.chkActive);
      this.ViewCalculation.Controls.Add((Control) this.rbAll);
      this.ViewCalculation.Controls.Add((Control) this.rbFitStat);
      this.ViewCalculation.Controls.Add((Control) this.rbNotFitStat);
      this.ViewCalculation.Controls.Add((Control) this.label11);
      this.ViewCalculation.Controls.Add((Control) this.label10);
      this.ViewCalculation.Controls.Add((Control) this.txtNote);
      this.ViewCalculation.Controls.Add((Control) this.txtUser);
      this.ViewCalculation.Controls.Add((Control) this.txtTimePointData);
      this.ViewCalculation.Controls.Add((Control) this.txtTimePointCalc);
      this.ViewCalculation.Controls.Add((Control) this.txtNParameterTotal);
      this.ViewCalculation.Controls.Add((Control) this.txtNStatistically);
      this.ViewCalculation.Controls.Add((Control) this.txtPercentStatistically);
      this.ViewCalculation.Controls.Add((Control) this.txtDoNotFitStatistically);
      this.ViewCalculation.Controls.Add((Control) this.label7);
      this.ViewCalculation.Controls.Add((Control) this.label8);
      this.ViewCalculation.Controls.Add((Control) this.label22);
      this.ViewCalculation.Controls.Add((Control) this.label23);
      this.ViewCalculation.Controls.Add((Control) this.label25);
      this.ViewCalculation.Controls.Add((Control) this.label26);
      this.ViewCalculation.Controls.Add((Control) this.label28);
      this.ViewCalculation.Controls.Add((Control) this.label29);
      this.ViewCalculation.Controls.Add((Control) this.label30);
      this.ViewCalculation.Controls.Add((Control) this.dataGridView1);
      this.ViewCalculation.Controls.Add((Control) this.label9);
      this.ViewCalculation.Controls.Add((Control) this.label24);
      this.ViewCalculation.Controls.Add((Control) this.label27);
      this.ViewCalculation.Location = new Point(4, 22);
      this.ViewCalculation.Name = "ViewCalculation";
      this.ViewCalculation.Padding = new Padding(3);
      this.ViewCalculation.Size = new Size(1410, 1013);
      this.ViewCalculation.TabIndex = 1;
      this.ViewCalculation.Tag = (object) "790";
      this.ViewCalculation.Text = "Calculation View";
      this.label87.AutoSize = true;
      this.label87.Location = new Point(98, 83);
      this.label87.Name = "label87";
      this.label87.Size = new Size(0, 13);
      this.label87.TabIndex = 99;
      this.btnPrint2.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnPrint2.Location = new Point(145, 530);
      this.btnPrint2.Name = "btnPrint2";
      this.btnPrint2.Size = new Size(97, 23);
      this.btnPrint2.TabIndex = 98;
      this.btnPrint2.Text = "Plot Print";
      this.btnPrint2.UseVisualStyleBackColor = true;
      this.btnPrint2.Visible = false;
      this.btnPrint2.Click += new EventHandler(this.btnPrint2_Click);
      this.btnZoom1.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnZoom1.Location = new Point(145, 501);
      this.btnZoom1.Name = "btnZoom1";
      this.btnZoom1.Size = new Size(97, 23);
      this.btnZoom1.TabIndex = 97;
      this.btnZoom1.Text = "Zoom In";
      this.btnZoom1.UseVisualStyleBackColor = true;
      this.btnZoom1.Visible = false;
      this.btnZoom1.Click += new EventHandler(this.btnZoom1_Click);
      this.panel2.BackColor = SystemColors.InactiveCaption;
      this.panel2.Controls.Add((Control) this.label21);
      this.panel2.Controls.Add((Control) this.label20);
      this.panel2.Controls.Add((Control) this.label19);
      this.panel2.Controls.Add((Control) this.label18);
      this.panel2.Controls.Add((Control) this.label17);
      this.panel2.Controls.Add((Control) this.label12);
      this.panel2.Controls.Add((Control) this.label13);
      this.panel2.Controls.Add((Control) this.label14);
      this.panel2.Controls.Add((Control) this.label15);
      this.panel2.Controls.Add((Control) this.label16);
      this.panel2.Location = new Point(100, 274);
      this.panel2.Name = "panel2";
      this.panel2.Size = new Size(1138, 21);
      this.panel2.TabIndex = 96;
      this.label21.AutoSize = true;
      this.label21.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label21.Location = new Point(3, 0);
      this.label21.Name = "label21";
      this.label21.Size = new Size(92, 20);
      this.label21.TabIndex = 34;
      this.label21.Text = "Parameter";
      this.label20.AutoSize = true;
      this.label20.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label20.Location = new Point(111, 1);
      this.label20.Name = "label20";
      this.label20.Size = new Size(66, 20);
      this.label20.TabIndex = 33;
      this.label20.Text = "Total N";
      this.label19.AutoSize = true;
      this.label19.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label19.Location = new Point(193, 1);
      this.label19.Name = "label19";
      this.label19.Size = new Size(62, 20);
      this.label19.TabIndex = 32;
      this.label19.Text = "KPI0%";
      this.label18.AutoSize = true;
      this.label18.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label18.Location = new Point(272, 1);
      this.label18.Name = "label18";
      this.label18.Size = new Size(62, 20);
      this.label18.TabIndex = 31;
      this.label18.Text = "KPI1%";
      this.label17.AutoSize = true;
      this.label17.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label17.Location = new Point(360, 1);
      this.label17.Name = "label17";
      this.label17.Size = new Size(62, 20);
      this.label17.TabIndex = 30;
      this.label17.Text = "KPI2%";
      this.label12.AutoSize = true;
      this.label12.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label12.Location = new Point(898, 1);
      this.label12.Name = "label12";
      this.label12.Size = new Size(47, 20);
      this.label12.TabIndex = 25;
      this.label12.Text = "Note";
      this.label13.AutoSize = true;
      this.label13.Font = new Font("Microsoft Sans Serif", 11.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label13.Location = new Point(749, 3);
      this.label13.Name = "label13";
      this.label13.Size = new Size(140, 18);
      this.label13.TabIndex = 26;
      this.label13.Text = "Relevant for disc.";
      this.label14.AutoSize = true;
      this.label14.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label14.Location = new Point(675, 1);
      this.label14.Name = "label14";
      this.label14.Size = new Size(53, 20);
      this.label14.TabIndex = 27;
      this.label14.Text = "Chart";
      this.label15.AutoSize = true;
      this.label15.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label15.Location = new Point(538, 1);
      this.label15.Name = "label15";
      this.label15.Size = new Size(119, 20);
      this.label15.TabIndex = 28;
      this.label15.Text = "fit statistically";
      this.label16.AutoSize = true;
      this.label16.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label16.Location = new Point(442, 1);
      this.label16.Name = "label16";
      this.label16.Size = new Size(62, 20);
      this.label16.TabIndex = 29;
      this.label16.Text = "KPI3%";
      this.dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView2.Location = new Point(85, -29);
      this.dataGridView2.Name = "dataGridView2";
      this.dataGridView2.Size = new Size(60, 23);
      this.dataGridView2.TabIndex = 95;
      this.dataGridView2.Visible = false;
      this.cmbCalcID.AutoCompleteCustomSource.AddRange(new string[1]
      {
        "Select Calculation ID"
      });
      this.cmbCalcID.FormattingEnabled = true;
      this.cmbCalcID.Location = new Point(510, 35);
      this.cmbCalcID.Name = "cmbCalcID";
      this.cmbCalcID.Size = new Size(121, 21);
      this.cmbCalcID.TabIndex = 94;
      this.cmbCalcID.SelectedIndexChanged += new EventHandler(this.cmbCalcID_SelectedIndexChanged_1);
      this.picGraph.Location = new Point(297, 501);
      this.picGraph.Name = "picGraph";
      this.picGraph.Size = new Size(941, 506);
      this.picGraph.TabIndex = 93;
      this.picGraph.TabStop = false;
      this.panel1.AutoScroll = true;
      this.panel1.BackColor = Color.White;
      this.panel1.Location = new Point(101, 297);
      this.panel1.Name = "panel1";
      this.panel1.Size = new Size(1137, 190);
      this.panel1.TabIndex = 92;
      this.btnSave.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnSave.Location = new Point(1077, 172);
      this.btnSave.Name = "btnSave";
      this.btnSave.Size = new Size(161, 50);
      this.btnSave.TabIndex = 91;
      this.btnSave.Text = "2-Save results";
      this.btnSave.UseVisualStyleBackColor = true;
      this.btnSave.Click += new EventHandler(this.btnSave_Click_1);
      this.chkActive.AutoSize = true;
      this.chkActive.Location = new Point(654, 172);
      this.chkActive.Name = "chkActive";
      this.chkActive.Size = new Size(56, 17);
      this.chkActive.TabIndex = 90;
      this.chkActive.Text = "Active";
      this.chkActive.UseVisualStyleBackColor = true;
      this.rbAll.AutoSize = true;
      this.rbAll.Enabled = false;
      this.rbAll.Location = new Point(206, 205);
      this.rbAll.Name = "rbAll";
      this.rbAll.Size = new Size(36, 17);
      this.rbAll.TabIndex = 89;
      this.rbAll.TabStop = true;
      this.rbAll.Text = "All";
      this.rbAll.UseVisualStyleBackColor = true;
      this.rbAll.CheckedChanged += new EventHandler(this.rbAll_CheckedChanged_1);
      this.rbFitStat.AutoSize = true;
      this.rbFitStat.Enabled = false;
      this.rbFitStat.Location = new Point(206, 228);
      this.rbFitStat.Name = "rbFitStat";
      this.rbFitStat.Size = new Size(86, 17);
      this.rbFitStat.TabIndex = 88;
      this.rbFitStat.TabStop = true;
      this.rbFitStat.Text = "fit statistically";
      this.rbFitStat.UseVisualStyleBackColor = true;
      this.rbFitStat.CheckedChanged += new EventHandler(this.rbFitStat_CheckedChanged_1);
      this.rbNotFitStat.AutoSize = true;
      this.rbNotFitStat.Enabled = false;
      this.rbNotFitStat.Location = new Point(206, 251);
      this.rbNotFitStat.Name = "rbNotFitStat";
      this.rbNotFitStat.Size = new Size(119, 17);
      this.rbNotFitStat.TabIndex = 87;
      this.rbNotFitStat.TabStop = true;
      this.rbNotFitStat.Text = "do not fit statistically";
      this.rbNotFitStat.UseVisualStyleBackColor = true;
      this.rbNotFitStat.CheckedChanged += new EventHandler(this.rbNotFitStat_CheckedChanged_1);
      this.label11.AutoSize = true;
      this.label11.Location = new Point(97, 205);
      this.label11.Name = "label11";
      this.label11.Size = new Size(80, 13);
      this.label11.TabIndex = 86;
      this.label11.Text = "Filter table view";
      this.label10.AutoSize = true;
      this.label10.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label10.ImageAlign = ContentAlignment.TopLeft;
      this.label10.Location = new Point(96, 165);
      this.label10.Name = "label10";
      this.label10.Size = new Size(215, 24);
      this.label10.TabIndex = 85;
      this.label10.Text = "Results per parameter";
      this.txtNote.Location = new Point(654, 31);
      this.txtNote.Multiline = true;
      this.txtNote.Name = "txtNote";
      this.txtNote.Size = new Size(584, 116);
      this.txtNote.TabIndex = 84;
      this.txtUser.Location = new Point(510, 63);
      this.txtUser.Name = "txtUser";
      this.txtUser.Size = new Size(100, 20);
      this.txtUser.TabIndex = 83;
      this.txtTimePointData.Location = new Point(510, 93);
      this.txtTimePointData.Name = "txtTimePointData";
      this.txtTimePointData.Size = new Size(100, 20);
      this.txtTimePointData.TabIndex = 82;
      this.txtTimePointCalc.Enabled = false;
      this.txtTimePointCalc.Location = new Point(510, (int) sbyte.MaxValue);
      this.txtTimePointCalc.Name = "txtTimePointCalc";
      this.txtTimePointCalc.Size = new Size(100, 20);
      this.txtTimePointCalc.TabIndex = 81;
      this.txtNParameterTotal.Enabled = false;
      this.txtNParameterTotal.Location = new Point(270, 35);
      this.txtNParameterTotal.Name = "txtNParameterTotal";
      this.txtNParameterTotal.Size = new Size(100, 20);
      this.txtNParameterTotal.TabIndex = 80;
      this.txtNStatistically.Enabled = false;
      this.txtNStatistically.Location = new Point(270, 63);
      this.txtNStatistically.Name = "txtNStatistically";
      this.txtNStatistically.Size = new Size(100, 20);
      this.txtNStatistically.TabIndex = 79;
      this.txtPercentStatistically.Enabled = false;
      this.txtPercentStatistically.Location = new Point(270, 93);
      this.txtPercentStatistically.Name = "txtPercentStatistically";
      this.txtPercentStatistically.Size = new Size(100, 20);
      this.txtPercentStatistically.TabIndex = 78;
      this.txtDoNotFitStatistically.Enabled = false;
      this.txtDoNotFitStatistically.Location = new Point(270, (int) sbyte.MaxValue);
      this.txtDoNotFitStatistically.Name = "txtDoNotFitStatistically";
      this.txtDoNotFitStatistically.Size = new Size(100, 20);
      this.txtDoNotFitStatistically.TabIndex = 77;
      this.label7.AutoSize = true;
      this.label7.Location = new Point(97, 42);
      this.label7.Name = "label7";
      this.label7.Size = new Size(98, 13);
      this.label7.TabIndex = 76;
      this.label7.Text = "N \"parameter total\"";
      this.label8.AutoSize = true;
      this.label8.Location = new Point(97, 70);
      this.label8.Name = "label8";
      this.label8.Size = new Size(89, 13);
      this.label8.TabIndex = 75;
      this.label8.Text = "N \"fit statistically\"";
      this.label22.AutoSize = true;
      this.label22.Location = new Point(98, 100);
      this.label22.Name = "label22";
      this.label22.Size = new Size(89, 13);
      this.label22.TabIndex = 74;
      this.label22.Text = "% \"fit statistically\"";
      this.label23.AutoSize = true;
      this.label23.Location = new Point(97, 130);
      this.label23.Name = "label23";
      this.label23.Size = new Size(122, 13);
      this.label23.TabIndex = 73;
      this.label23.Text = "N \"do not fit statistically\"";
      this.label23.Click += new EventHandler(this.label23_Click);
      this.label25.AutoSize = true;
      this.label25.Location = new Point(400, 63);
      this.label25.Name = "label25";
      this.label25.Size = new Size(29, 13);
      this.label25.TabIndex = 72;
      this.label25.Text = "User";
      this.label26.AutoSize = true;
      this.label26.Location = new Point(400, 38);
      this.label26.Name = "label26";
      this.label26.Size = new Size(88, 13);
      this.label26.TabIndex = 71;
      this.label26.Text = "1 - Calculation ID";
      this.label28.AutoSize = true;
      this.label28.Location = new Point(400, 96);
      this.label28.Name = "label28";
      this.label28.Size = new Size(77, 13);
      this.label28.TabIndex = 70;
      this.label28.Text = "Timepoint data";
      this.label29.AutoSize = true;
      this.label29.Location = new Point(400, 134);
      this.label29.Name = "label29";
      this.label29.Size = new Size(107, 13);
      this.label29.TabIndex = 69;
      this.label29.Text = "Timepoint calculation";
      this.label30.AutoSize = true;
      this.label30.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label30.ImageAlign = ContentAlignment.TopLeft;
      this.label30.Location = new Point(96, -3);
      this.label30.Name = "label30";
      this.label30.Size = new Size(231, 24);
      this.label30.TabIndex = 68;
      this.label30.Text = "Results per calculations";
      this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView1.Location = new Point(85, -82);
      this.dataGridView1.Name = "dataGridView1";
      this.dataGridView1.Size = new Size(60, 23);
      this.dataGridView1.TabIndex = 67;
      this.dataGridView1.Visible = false;
      this.label9.AutoSize = true;
      this.label9.Location = new Point(97, -11);
      this.label9.Name = "label9";
      this.label9.Size = new Size(98, 13);
      this.label9.TabIndex = 48;
      this.label9.Text = "N \"parameter total\"";
      this.label24.AutoSize = true;
      this.label24.Location = new Point(400, -15);
      this.label24.Name = "label24";
      this.label24.Size = new Size(73, 13);
      this.label24.TabIndex = 43;
      this.label24.Text = "Calculation ID";
      this.label27.AutoSize = true;
      this.label27.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label27.ImageAlign = ContentAlignment.TopLeft;
      this.label27.Location = new Point(96, -56);
      this.label27.Name = "label27";
      this.label27.Size = new Size(231, 24);
      this.label27.TabIndex = 40;
      this.label27.Text = "Results per calculations";
      this.SearchCalculation.Controls.Add((Control) this.label83);
      this.SearchCalculation.Controls.Add((Control) this.label82);
      this.SearchCalculation.Controls.Add((Control) this.btnPrint);
      this.SearchCalculation.Controls.Add((Control) this.btnZoomIn1);
      this.SearchCalculation.Controls.Add((Control) this.pictureBox2);
      this.SearchCalculation.Controls.Add((Control) this.panel4);
      this.SearchCalculation.Controls.Add((Control) this.panel6);
      this.SearchCalculation.Controls.Add((Control) this.panel17);
      this.SearchCalculation.Controls.Add((Control) this.panel16);
      this.SearchCalculation.Controls.Add((Control) this.panel15);
      this.SearchCalculation.Controls.Add((Control) this.label34);
      this.SearchCalculation.Controls.Add((Control) this.panel13);
      this.SearchCalculation.Controls.Add((Control) this.panel12);
      this.SearchCalculation.Controls.Add((Control) this.panel11);
      this.SearchCalculation.Controls.Add((Control) this.panel10);
      this.SearchCalculation.Controls.Add((Control) this.panel5);
      this.SearchCalculation.Controls.Add((Control) this.panelSelection);
      this.SearchCalculation.Controls.Add((Control) this.panel3);
      this.SearchCalculation.Controls.Add((Control) this.groupFilterSelection);
      this.SearchCalculation.Controls.Add((Control) this.panelButtons);
      this.SearchCalculation.ForeColor = Color.FromArgb(64, 0, 64);
      this.SearchCalculation.Location = new Point(4, 22);
      this.SearchCalculation.Name = "SearchCalculation";
      this.SearchCalculation.Padding = new Padding(3);
      this.SearchCalculation.Size = new Size(1410, 1013);
      this.SearchCalculation.TabIndex = 2;
      this.SearchCalculation.Tag = (object) "791";
      this.SearchCalculation.Text = "Calculation Search";
      this.SearchCalculation.Click += new EventHandler(this.SearchCalculation_onload);
      this.label83.AutoSize = true;
      this.label83.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label83.Location = new Point(41, 521);
      this.label83.Name = "label83";
      this.label83.Size = new Size(24, 16);
      this.label83.TabIndex = 47;
      this.label83.Text = "10";
      this.label83.Visible = false;
      this.label82.AutoSize = true;
      this.label82.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label82.Location = new Point(41, 492);
      this.label82.Name = "label82";
      this.label82.Size = new Size(16, 16);
      this.label82.TabIndex = 46;
      this.label82.Text = "9";
      this.label82.Visible = false;
      this.btnPrint.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnPrint.Location = new Point(66, 516);
      this.btnPrint.Name = "btnPrint";
      this.btnPrint.Size = new Size(87, 23);
      this.btnPrint.TabIndex = 1;
      this.btnPrint.Text = "Plot Print";
      this.btnPrint.UseVisualStyleBackColor = true;
      this.btnPrint.Visible = false;
      this.btnPrint.Click += new EventHandler(this.btnPrint_Click);
      this.btnZoomIn1.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnZoomIn1.Location = new Point(66, 487);
      this.btnZoomIn1.Name = "btnZoomIn1";
      this.btnZoomIn1.Size = new Size(87, 23);
      this.btnZoomIn1.TabIndex = 43;
      this.btnZoomIn1.Text = "Zoom In";
      this.btnZoomIn1.UseVisualStyleBackColor = true;
      this.btnZoomIn1.Visible = false;
      this.btnZoomIn1.Click += new EventHandler(this.btnZoomIn1_Click);
      this.pictureBox2.Location = new Point(186, 487);
      this.pictureBox2.Name = "pictureBox2";
      this.pictureBox2.Size = new Size(1085, 520);
      this.pictureBox2.TabIndex = 42;
      this.pictureBox2.TabStop = false;
      this.panel4.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel4.Controls.Add((Control) this.label48);
      this.panel4.Location = new Point(533, 274);
      this.panel4.Name = "panel4";
      this.panel4.Size = new Size(81, 21);
      this.panel4.TabIndex = 41;
      this.label48.AutoSize = true;
      this.label48.Location = new Point(3, 4);
      this.label48.Name = "label48";
      this.label48.Size = new Size(71, 13);
      this.label48.TabIndex = 7;
      this.label48.Text = "Fit statistically";
      this.panel6.AutoScroll = true;
      this.panel6.BackColor = Color.White;
      this.panel6.Location = new Point(30, 313);
      this.panel6.Name = "panel6";
      this.panel6.Size = new Size(1241, 158);
      this.panel6.TabIndex = 40;
      this.panel17.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel17.Controls.Add((Control) this.label31);
      this.panel17.Location = new Point(125, 274);
      this.panel17.Name = "panel17";
      this.panel17.Size = new Size(95, 21);
      this.panel17.TabIndex = 39;
      this.label31.AutoSize = true;
      this.label31.Location = new Point(3, 4);
      this.label31.Name = "label31";
      this.label31.Size = new Size(64, 13);
      this.label31.TabIndex = 7;
      this.label31.Text = "VIRT_OZID";
      this.panel16.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel16.Controls.Add((Control) this.label32);
      this.panel16.Location = new Point(226, 274);
      this.panel16.Name = "panel16";
      this.panel16.Size = new Size(147, 21);
      this.panel16.TabIndex = 38;
      this.label32.AutoSize = true;
      this.label32.Location = new Point(3, 4);
      this.label32.Name = "label32";
      this.label32.Size = new Size(60, 13);
      this.label32.TabIndex = 6;
      this.label32.Text = "KPI (count)";
      this.panel15.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel15.Controls.Add((Control) this.label33);
      this.panel15.Location = new Point(379, 274);
      this.panel15.Name = "panel15";
      this.panel15.Size = new Size(151, 21);
      this.panel15.TabIndex = 37;
      this.label33.AutoSize = true;
      this.label33.Location = new Point(0, 4);
      this.label33.Name = "label33";
      this.label33.Size = new Size(38, 13);
      this.label33.TabIndex = 5;
      this.label33.Text = "KPI(%)";
      this.label34.AutoSize = true;
      this.label34.Location = new Point(23, 24);
      this.label34.Name = "label34";
      this.label34.Size = new Size(0, 13);
      this.label34.TabIndex = 31;
      this.panel13.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel13.Controls.Add((Control) this.label35);
      this.panel13.Location = new Point(617, 274);
      this.panel13.Name = "panel13";
      this.panel13.Size = new Size(99, 21);
      this.panel13.TabIndex = 36;
      this.label35.AutoSize = true;
      this.label35.Location = new Point(3, 4);
      this.label35.Name = "label35";
      this.label35.Size = new Size(32, 13);
      this.label35.TabIndex = 3;
      this.label35.Text = "Chart";
      this.panel12.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel12.Controls.Add((Control) this.label36);
      this.panel12.Location = new Point(721, 274);
      this.panel12.Name = "panel12";
      this.panel12.Size = new Size(202, 21);
      this.panel12.TabIndex = 35;
      this.label36.AutoSize = true;
      this.label36.Location = new Point(3, 4);
      this.label36.Name = "label36";
      this.label36.Size = new Size(117, 13);
      this.label36.TabIndex = 2;
      this.label36.Text = "Relevant for discussion";
      this.panel11.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel11.Controls.Add((Control) this.label37);
      this.panel11.Location = new Point(929, 274);
      this.panel11.Name = "panel11";
      this.panel11.Size = new Size(99, 21);
      this.panel11.TabIndex = 34;
      this.label37.AutoSize = true;
      this.label37.Location = new Point(3, 4);
      this.label37.Name = "label37";
      this.label37.Size = new Size(37, 13);
      this.label37.TabIndex = 1;
      this.label37.Text = "Active";
      this.panel10.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel10.Controls.Add((Control) this.label38);
      this.panel10.Location = new Point(1031, 274);
      this.panel10.Name = "panel10";
      this.panel10.Size = new Size(240, 21);
      this.panel10.TabIndex = 33;
      this.label38.AutoSize = true;
      this.label38.Location = new Point(3, 4);
      this.label38.Name = "label38";
      this.label38.Size = new Size(30, 13);
      this.label38.TabIndex = 8;
      this.label38.Text = "Note";
      this.panel5.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel5.Controls.Add((Control) this.label47);
      this.panel5.Location = new Point(30, 274);
      this.panel5.Name = "panel5";
      this.panel5.Size = new Size(89, 21);
      this.panel5.TabIndex = 32;
      this.label47.AutoSize = true;
      this.label47.Location = new Point(3, 4);
      this.label47.Name = "label47";
      this.label47.Size = new Size(21, 13);
      this.label47.TabIndex = 0;
      this.label47.Text = "No";
      this.panelSelection.Controls.Add((Control) this.label81);
      this.panelSelection.Controls.Add((Control) this.label78);
      this.panelSelection.Controls.Add((Control) this.label77);
      this.panelSelection.Controls.Add((Control) this.label44);
      this.panelSelection.Controls.Add((Control) this.lbGraph);
      this.panelSelection.Controls.Add((Control) this.rbKPI3);
      this.panelSelection.Controls.Add((Control) this.rbKPI2);
      this.panelSelection.Controls.Add((Control) this.rbKPI1);
      this.panelSelection.Controls.Add((Control) this.rbKPI0);
      this.panelSelection.Controls.Add((Control) this.label45);
      this.panelSelection.Controls.Add((Control) this.lbOzid);
      this.panelSelection.Controls.Add((Control) this.label46);
      this.panelSelection.Location = new Point(659, 9);
      this.panelSelection.Name = "panelSelection";
      this.panelSelection.Size = new Size(373, 246);
      this.panelSelection.TabIndex = 26;
      this.label81.AutoSize = true;
      this.label81.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label81.Location = new Point(31, 125);
      this.label81.Name = "label81";
      this.label81.Size = new Size(16, 16);
      this.label81.TabIndex = 48;
      this.label81.Text = "5";
      this.label78.AutoSize = true;
      this.label78.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label78.Location = new Point(336, 28);
      this.label78.Name = "label78";
      this.label78.Size = new Size(16, 16);
      this.label78.TabIndex = 47;
      this.label78.Text = "8";
      this.label77.AutoSize = true;
      this.label77.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label77.Location = new Point(179, 28);
      this.label77.Name = "label77";
      this.label77.Size = new Size(16, 16);
      this.label77.TabIndex = 46;
      this.label77.Text = "7";
      this.label44.AutoSize = true;
      this.label44.Location = new Point(210, 8);
      this.label44.Name = "label44";
      this.label44.Size = new Size(58, 13);
      this.label44.TabIndex = 9;
      this.label44.Text = "Select Plot";
      this.lbGraph.BackColor = Color.White;
      this.lbGraph.FormattingEnabled = true;
      this.lbGraph.Location = new Point(210, 23);
      this.lbGraph.Name = "lbGraph";
      this.lbGraph.Size = new Size(151, 199);
      this.lbGraph.TabIndex = 8;
      this.lbGraph.SelectedIndexChanged += new EventHandler(this.lbGraph_SelectedIndexChanged_1);
      this.rbKPI3.AutoSize = true;
      this.rbKPI3.BackColor = SystemColors.Control;
      this.rbKPI3.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
      this.rbKPI3.Location = new Point(10, 101);
      this.rbKPI3.Name = "rbKPI3";
      this.rbKPI3.Size = new Size(69, 19);
      this.rbKPI3.TabIndex = 7;
      this.rbKPI3.TabStop = true;
      this.rbKPI3.Text = "KPI(%)3";
      this.rbKPI3.UseVisualStyleBackColor = false;
      this.rbKPI3.CheckedChanged += new EventHandler(this.rbKPI3_CheckedChanged);
      this.rbKPI2.AutoSize = true;
      this.rbKPI2.BackColor = SystemColors.Control;
      this.rbKPI2.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
      this.rbKPI2.Location = new Point(10, 75);
      this.rbKPI2.Name = "rbKPI2";
      this.rbKPI2.Size = new Size(69, 19);
      this.rbKPI2.TabIndex = 6;
      this.rbKPI2.TabStop = true;
      this.rbKPI2.Text = "KPI(%)2";
      this.rbKPI2.UseVisualStyleBackColor = false;
      this.rbKPI1.AutoSize = true;
      this.rbKPI1.BackColor = SystemColors.Control;
      this.rbKPI1.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
      this.rbKPI1.Location = new Point(10, 49);
      this.rbKPI1.Name = "rbKPI1";
      this.rbKPI1.Size = new Size(69, 19);
      this.rbKPI1.TabIndex = 5;
      this.rbKPI1.TabStop = true;
      this.rbKPI1.Text = "KPI(%)1";
      this.rbKPI1.UseVisualStyleBackColor = false;
      this.rbKPI0.AutoSize = true;
      this.rbKPI0.BackColor = SystemColors.Control;
      this.rbKPI0.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
      this.rbKPI0.ForeColor = SystemColors.ControlText;
      this.rbKPI0.Location = new Point(10, 28);
      this.rbKPI0.Name = "rbKPI0";
      this.rbKPI0.Size = new Size(69, 19);
      this.rbKPI0.TabIndex = 4;
      this.rbKPI0.TabStop = true;
      this.rbKPI0.Text = "KPI(%)0";
      this.rbKPI0.UseVisualStyleBackColor = false;
      this.rbKPI0.CheckedChanged += new EventHandler(this.rbKPI0_CheckedChanged);
      this.label45.AutoSize = true;
      this.label45.Font = new Font("Segoe UI", 12f);
      this.label45.Location = new Point(81, 3);
      this.label45.Name = "label45";
      this.label45.Size = new Size(131, 21);
      this.label45.TabIndex = 3;
      this.label45.Text = "Select parameter ";
      this.lbOzid.BackColor = Color.White;
      this.lbOzid.FormattingEnabled = true;
      this.lbOzid.Location = new Point(85, 23);
      this.lbOzid.Name = "lbOzid";
      this.lbOzid.Size = new Size(118, 199);
      this.lbOzid.TabIndex = 2;
      this.lbOzid.SelectedIndexChanged += new EventHandler(this.lbOzid_SelectedIndexChanged_1);
      this.label46.AutoSize = true;
      this.label46.Font = new Font("Segoe UI", 12f);
      this.label46.Location = new Point(10, 3);
      this.label46.Name = "label46";
      this.label46.Size = new Size(77, 21);
      this.label46.TabIndex = 1;
      this.label46.Text = "Select KPI";
      this.panel3.Controls.Add((Control) this.label74);
      this.panel3.Controls.Add((Control) this.label75);
      this.panel3.Controls.Add((Control) this.groupBox1);
      this.panel3.Location = new Point(30, 6);
      this.panel3.Name = "panel3";
      this.panel3.Size = new Size(321, 249);
      this.panel3.TabIndex = 27;
      this.label74.AutoSize = true;
      this.label74.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label74.Location = new Point(299, 27);
      this.label74.Name = "label74";
      this.label74.Size = new Size(16, 16);
      this.label74.TabIndex = 46;
      this.label74.Text = "1";
      this.label75.AutoSize = true;
      this.label75.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label75.Location = new Point(299, 58);
      this.label75.Name = "label75";
      this.label75.Size = new Size(16, 16);
      this.label75.TabIndex = 45;
      this.label75.Text = "2";
      this.groupBox1.Controls.Add((Control) this.label79);
      this.groupBox1.Controls.Add((Control) this.chkVirtOzid2);
      this.groupBox1.Controls.Add((Control) this.dtCalcDateTime);
      this.groupBox1.Controls.Add((Control) this.cmbProdID2);
      this.groupBox1.Controls.Add((Control) this.rbNotActive1);
      this.groupBox1.Controls.Add((Control) this.rbActive1);
      this.groupBox1.Controls.Add((Control) this.cmbCalcID2);
      this.groupBox1.Controls.Add((Control) this.rbAll1);
      this.groupBox1.Controls.Add((Control) this.txtNote1);
      this.groupBox1.Controls.Add((Control) this.label39);
      this.groupBox1.Controls.Add((Control) this.label40);
      this.groupBox1.Controls.Add((Control) this.label41);
      this.groupBox1.Controls.Add((Control) this.label42);
      this.groupBox1.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.groupBox1.Location = new Point(0, 0);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new Size(306, 246);
      this.groupBox1.TabIndex = 0;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = " ";
      this.label79.AutoSize = true;
      this.label79.Location = new Point(91, 153);
      this.label79.Name = "label79";
      this.label79.Size = new Size(19, 21);
      this.label79.TabIndex = 46;
      this.label79.Text = "3";
      this.chkVirtOzid2.CheckOnClick = true;
      this.chkVirtOzid2.FormattingEnabled = true;
      this.chkVirtOzid2.Location = new Point(147, 117);
      this.chkVirtOzid2.Name = "chkVirtOzid2";
      this.chkVirtOzid2.Size = new Size(150, 76);
      this.chkVirtOzid2.TabIndex = 16;
      this.dtCalcDateTime.CalendarFont = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.dtCalcDateTime.Location = new Point(147, 81);
      this.dtCalcDateTime.Name = "dtCalcDateTime";
      this.dtCalcDateTime.Size = new Size(150, 29);
      this.dtCalcDateTime.TabIndex = 15;
      this.cmbProdID2.Font = new Font("Segoe UI", 12f);
      this.cmbProdID2.FormattingEnabled = true;
      this.cmbProdID2.Location = new Point(131, 18);
      this.cmbProdID2.Name = "cmbProdID2";
      this.cmbProdID2.Size = new Size(166, 29);
      this.cmbProdID2.TabIndex = 13;
      this.cmbProdID2.SelectedIndexChanged += new EventHandler(this.cmbProdID2_SelectedIndexChanged);
      this.rbNotActive1.AutoSize = true;
      this.rbNotActive1.BackColor = SystemColors.Control;
      this.rbNotActive1.Font = new Font("Segoe UI", 12f);
      this.rbNotActive1.Location = new Point(16, 200);
      this.rbNotActive1.Name = "rbNotActive1";
      this.rbNotActive1.Size = new Size(98, 25);
      this.rbNotActive1.TabIndex = 12;
      this.rbNotActive1.TabStop = true;
      this.rbNotActive1.Text = "Not active";
      this.rbNotActive1.UseVisualStyleBackColor = false;
      this.rbActive1.AutoSize = true;
      this.rbActive1.BackColor = SystemColors.Control;
      this.rbActive1.Font = new Font("Segoe UI", 12f);
      this.rbActive1.Location = new Point(16, 176);
      this.rbActive1.Name = "rbActive1";
      this.rbActive1.Size = new Size(70, 25);
      this.rbActive1.TabIndex = 11;
      this.rbActive1.TabStop = true;
      this.rbActive1.Text = "Active";
      this.rbActive1.UseVisualStyleBackColor = false;
      this.cmbCalcID2.Font = new Font("Segoe UI", 12f);
      this.cmbCalcID2.FormattingEnabled = true;
      this.cmbCalcID2.Location = new Point(131, 50);
      this.cmbCalcID2.Name = "cmbCalcID2";
      this.cmbCalcID2.Size = new Size(166, 29);
      this.cmbCalcID2.TabIndex = 10;
      this.cmbCalcID2.SelectedIndexChanged += new EventHandler(this.cmbCalcID2_SelectedIndexChanged);
      this.rbAll1.AutoSize = true;
      this.rbAll1.BackColor = SystemColors.Control;
      this.rbAll1.Font = new Font("Segoe UI", 12f);
      this.rbAll1.Location = new Point(16, 153);
      this.rbAll1.Name = "rbAll1";
      this.rbAll1.Size = new Size(46, 25);
      this.rbAll1.TabIndex = 10;
      this.rbAll1.TabStop = true;
      this.rbAll1.Text = "All";
      this.rbAll1.UseVisualStyleBackColor = false;
      this.txtNote1.Font = new Font("Segoe UI", 12f);
      this.txtNote1.Location = new Point(147, 199);
      this.txtNote1.Multiline = true;
      this.txtNote1.Name = "txtNote1";
      this.txtNote1.Size = new Size(150, 41);
      this.txtNote1.TabIndex = 8;
      this.label39.AutoSize = true;
      this.label39.Font = new Font("Segoe UI", 12f);
      this.label39.Location = new Point(3, 117);
      this.label39.Name = "label39";
      this.label39.Size = new Size(74, 21);
      this.label39.TabIndex = 3;
      this.label39.Text = "Virt_Ozid";
      this.label40.AutoSize = true;
      this.label40.Font = new Font("Segoe UI", 12f);
      this.label40.Location = new Point(3, 80);
      this.label40.Name = "label40";
      this.label40.Size = new Size(142, 21);
      this.label40.TabIndex = 2;
      this.label40.Text = "Date of calculation ";
      this.label41.AutoSize = true;
      this.label41.Font = new Font("Segoe UI", 12f);
      this.label41.Location = new Point(3, 58);
      this.label41.Name = "label41";
      this.label41.Size = new Size(106, 21);
      this.label41.TabIndex = 1;
      this.label41.Text = "Calculation ID";
      this.label42.AutoSize = true;
      this.label42.Font = new Font("Segoe UI", 12f);
      this.label42.Location = new Point(3, 26);
      this.label42.Name = "label42";
      this.label42.Size = new Size(97, 21);
      this.label42.TabIndex = 0;
      this.label42.Text = "Productcode";
      this.groupFilterSelection.Controls.Add((Control) this.label80);
      this.groupFilterSelection.Controls.Add((Control) this.label76);
      this.groupFilterSelection.Controls.Add((Control) this.btnGetHistoric);
      this.groupFilterSelection.Controls.Add((Control) this.rbNotFitStatF);
      this.groupFilterSelection.Controls.Add((Control) this.rbFitstatF);
      this.groupFilterSelection.Controls.Add((Control) this.rbAllF);
      this.groupFilterSelection.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.groupFilterSelection.Location = new Point(357, 11);
      this.groupFilterSelection.Name = "groupFilterSelection";
      this.groupFilterSelection.Size = new Size(296, 244);
      this.groupFilterSelection.TabIndex = 28;
      this.groupFilterSelection.TabStop = false;
      this.groupFilterSelection.Text = "Filter selection";
      this.label80.AutoSize = true;
      this.label80.Location = new Point(169, 47);
      this.label80.Name = "label80";
      this.label80.Size = new Size(19, 21);
      this.label80.TabIndex = 47;
      this.label80.Text = "4";
      this.label76.AutoSize = true;
      this.label76.Location = new Point(21, 167);
      this.label76.Name = "label76";
      this.label76.Size = new Size(19, 21);
      this.label76.TabIndex = 46;
      this.label76.Text = "6";
      this.btnGetHistoric.Location = new Point(15, 157);
      this.btnGetHistoric.Name = "btnGetHistoric";
      this.btnGetHistoric.Size = new Size(265, 64);
      this.btnGetHistoric.TabIndex = 26;
      this.btnGetHistoric.Text = "Get historic data";
      this.btnGetHistoric.UseVisualStyleBackColor = true;
      this.btnGetHistoric.Click += new EventHandler(this.btnGetHistoric_Click_1);
      this.rbNotFitStatF.AutoSize = true;
      this.rbNotFitStatF.BackColor = SystemColors.Control;
      this.rbNotFitStatF.Font = new Font("Segoe UI", 12f);
      this.rbNotFitStatF.Location = new Point(15, 97);
      this.rbNotFitStatF.Name = "rbNotFitStatF";
      this.rbNotFitStatF.Size = new Size(173, 25);
      this.rbNotFitStatF.TabIndex = 9;
      this.rbNotFitStatF.TabStop = true;
      this.rbNotFitStatF.Text = "Do not fit statistically";
      this.rbNotFitStatF.UseVisualStyleBackColor = false;
      this.rbFitstatF.AutoSize = true;
      this.rbFitstatF.BackColor = SystemColors.Control;
      this.rbFitstatF.Font = new Font("Segoe UI", 12f);
      this.rbFitstatF.Location = new Point(15, 72);
      this.rbFitstatF.Name = "rbFitstatF";
      this.rbFitstatF.Size = new Size(125, 25);
      this.rbFitstatF.TabIndex = 8;
      this.rbFitstatF.TabStop = true;
      this.rbFitstatF.Text = "Fit statistically";
      this.rbFitstatF.UseVisualStyleBackColor = false;
      this.rbAllF.AutoSize = true;
      this.rbAllF.BackColor = SystemColors.Control;
      this.rbAllF.Font = new Font("Segoe UI", 12f);
      this.rbAllF.Location = new Point(15, 49);
      this.rbAllF.Name = "rbAllF";
      this.rbAllF.Size = new Size(46, 25);
      this.rbAllF.TabIndex = 7;
      this.rbAllF.TabStop = true;
      this.rbAllF.Text = "All";
      this.rbAllF.UseVisualStyleBackColor = false;
      this.rbAllF.CheckedChanged += new EventHandler(this.radioButton3_CheckedChanged_1);
      this.panelButtons.Controls.Add((Control) this.label86);
      this.panelButtons.Controls.Add((Control) this.label85);
      this.panelButtons.Controls.Add((Control) this.label84);
      this.panelButtons.Controls.Add((Control) this.chActivate);
      this.panelButtons.Controls.Add((Control) this.chDeActivate);
      this.panelButtons.Controls.Add((Control) this.button2);
      this.panelButtons.Controls.Add((Control) this.button3);
      this.panelButtons.Location = new Point(1038, 9);
      this.panelButtons.Name = "panelButtons";
      this.panelButtons.Size = new Size(233, 246);
      this.panelButtons.TabIndex = 29;
      this.label86.AutoSize = true;
      this.label86.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label86.Location = new Point(206, 134);
      this.label86.Name = "label86";
      this.label86.Size = new Size(24, 16);
      this.label86.TabIndex = 50;
      this.label86.Text = "13";
      this.label85.AutoSize = true;
      this.label85.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label85.Location = new Point(206, 59);
      this.label85.Name = "label85";
      this.label85.Size = new Size(24, 16);
      this.label85.TabIndex = 49;
      this.label85.Text = "12";
      this.label84.AutoSize = true;
      this.label84.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label84.Location = new Point(206, 8);
      this.label84.Name = "label84";
      this.label84.Size = new Size(24, 16);
      this.label84.TabIndex = 48;
      this.label84.Text = "11";
      this.chActivate.AutoSize = true;
      this.chActivate.BackColor = SystemColors.Control;
      this.chActivate.Location = new Point(24, 7);
      this.chActivate.Name = "chActivate";
      this.chActivate.Size = new Size(119, 17);
      this.chActivate.TabIndex = 0;
      this.chActivate.Text = "Activate calculation";
      this.chActivate.UseVisualStyleBackColor = false;
      this.chActivate.CheckedChanged += new EventHandler(this.chActivate_CheckedChanged_1);
      this.chActivate.Click += new EventHandler(this.chActivate_Click);
      this.chDeActivate.AutoSize = true;
      this.chDeActivate.BackColor = SystemColors.Control;
      this.chDeActivate.Location = new Point(24, 29);
      this.chDeActivate.Name = "chDeActivate";
      this.chDeActivate.Size = new Size(132, 17);
      this.chDeActivate.TabIndex = 1;
      this.chDeActivate.Text = "Deactivate calculation";
      this.chDeActivate.UseVisualStyleBackColor = false;
      this.chDeActivate.CheckedChanged += new EventHandler(this.chDeActivate_CheckedChanged_1);
      this.chDeActivate.Click += new EventHandler(this.chDeActivate_Click);
      this.button2.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.button2.Location = new Point(24, 125);
      this.button2.Name = "button2";
      this.button2.Size = new Size(184, 50);
      this.button2.TabIndex = 2;
      this.button2.Text = "Compare to active data";
      this.button2.UseVisualStyleBackColor = true;
      this.button2.Click += new EventHandler(this.button2_Click_1);
      this.button3.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.button3.Location = new Point(24, 57);
      this.button3.Name = "button3";
      this.button3.Size = new Size(184, 51);
      this.button3.TabIndex = 3;
      this.button3.Text = "Save status and notes";
      this.button3.UseVisualStyleBackColor = true;
      this.button3.Click += new EventHandler(this.button3_Click);
      this.CompareCalculation.Controls.Add((Control) this.button7);
      this.CompareCalculation.Controls.Add((Control) this.button6);
      this.CompareCalculation.Controls.Add((Control) this.label71);
      this.CompareCalculation.Controls.Add((Control) this.button5);
      this.CompareCalculation.Controls.Add((Control) this.button4);
      this.CompareCalculation.Controls.Add((Control) this.btnPrint4);
      this.CompareCalculation.Controls.Add((Control) this.btnPrint3);
      this.CompareCalculation.Controls.Add((Control) this.btnZoomOut3);
      this.CompareCalculation.Controls.Add((Control) this.btnZoomOut4);
      this.CompareCalculation.Controls.Add((Control) this.btnCompareCalc);
      this.CompareCalculation.Controls.Add((Control) this.pictureBox4);
      this.CompareCalculation.Controls.Add((Control) this.pictureBox3);
      this.CompareCalculation.Controls.Add((Control) this.panel24);
      this.CompareCalculation.Controls.Add((Control) this.label68);
      this.CompareCalculation.Controls.Add((Control) this.label67);
      this.CompareCalculation.Controls.Add((Control) this.panel23);
      this.CompareCalculation.Controls.Add((Control) this.panel22);
      this.CompareCalculation.Controls.Add((Control) this.panel21);
      this.CompareCalculation.Controls.Add((Control) this.panel20);
      this.CompareCalculation.Controls.Add((Control) this.panel19);
      this.CompareCalculation.Controls.Add((Control) this.panel18);
      this.CompareCalculation.Controls.Add((Control) this.panel14);
      this.CompareCalculation.Controls.Add((Control) this.panel9);
      this.CompareCalculation.Controls.Add((Control) this.panel7);
      this.CompareCalculation.Controls.Add((Control) this.panel8);
      this.CompareCalculation.Controls.Add((Control) this.groupBox3);
      this.CompareCalculation.Controls.Add((Control) this.groupBox2);
      this.CompareCalculation.Location = new Point(4, 22);
      this.CompareCalculation.Name = "CompareCalculation";
      this.CompareCalculation.Size = new Size(1410, 1013);
      this.CompareCalculation.TabIndex = 3;
      this.CompareCalculation.Text = "Compare Calculations";
      this.CompareCalculation.UseVisualStyleBackColor = true;
      this.CompareCalculation.Click += new EventHandler(this.CompareCalculation_Click);
      this.button7.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.button7.Location = new Point(1167, 190);
      this.button7.Name = "button7";
      this.button7.Size = new Size(114, 23);
      this.button7.TabIndex = 66;
      this.button7.Text = "Select All";
      this.button7.UseVisualStyleBackColor = true;
      this.button7.Click += new EventHandler(this.button7_Click);
      this.button6.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.button6.Location = new Point(498, 181);
      this.button6.Name = "button6";
      this.button6.Size = new Size(114, 23);
      this.button6.TabIndex = 65;
      this.button6.Text = "Select All";
      this.button6.UseVisualStyleBackColor = true;
      this.button6.Click += new EventHandler(this.button6_Click);
      this.label71.AutoSize = true;
      this.label71.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label71.Location = new Point(507, 223);
      this.label71.Name = "label71";
      this.label71.Size = new Size(14, 13);
      this.label71.TabIndex = 64;
      this.label71.Text = "5";
      this.button5.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.button5.Location = new Point(1167, 170);
      this.button5.Name = "button5";
      this.button5.Size = new Size(114, 23);
      this.button5.TabIndex = 61;
      this.button5.Text = "DeSelect All";
      this.button5.UseVisualStyleBackColor = true;
      this.button5.Click += new EventHandler(this.button5_Click);
      this.button4.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.button4.Location = new Point(498, 161);
      this.button4.Name = "button4";
      this.button4.Size = new Size(114, 23);
      this.button4.TabIndex = 60;
      this.button4.Text = "DeSelect All";
      this.button4.UseVisualStyleBackColor = true;
      this.button4.Click += new EventHandler(this.button4_Click);
      this.btnPrint4.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnPrint4.Location = new Point(794, 509);
      this.btnPrint4.Name = "btnPrint4";
      this.btnPrint4.Size = new Size(142, 23);
      this.btnPrint4.TabIndex = 59;
      this.btnPrint4.Text = "Plot Print";
      this.btnPrint4.UseVisualStyleBackColor = true;
      this.btnPrint4.Visible = false;
      this.btnPrint4.Click += new EventHandler(this.btnPrint4_Click_1);
      this.btnPrint3.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnPrint3.Location = new Point(128, 509);
      this.btnPrint3.Name = "btnPrint3";
      this.btnPrint3.Size = new Size(134, 23);
      this.btnPrint3.TabIndex = 58;
      this.btnPrint3.Text = "Plot Print";
      this.btnPrint3.UseVisualStyleBackColor = true;
      this.btnPrint3.Visible = false;
      this.btnPrint3.Click += new EventHandler(this.btnPrint3_Click_1);
      this.btnZoomOut3.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnZoomOut3.Location = new Point(47, 509);
      this.btnZoomOut3.Name = "btnZoomOut3";
      this.btnZoomOut3.Size = new Size(115, 23);
      this.btnZoomOut3.TabIndex = 57;
      this.btnZoomOut3.Text = "Zoom In";
      this.btnZoomOut3.UseVisualStyleBackColor = true;
      this.btnZoomOut3.Visible = false;
      this.btnZoomOut3.Click += new EventHandler(this.btnZoomOut3_Click);
      this.btnZoomOut4.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.btnZoomOut4.Location = new Point(713, 509);
      this.btnZoomOut4.Name = "btnZoomOut4";
      this.btnZoomOut4.Size = new Size(115, 23);
      this.btnZoomOut4.TabIndex = 56;
      this.btnZoomOut4.Text = "Zoom In";
      this.btnZoomOut4.UseVisualStyleBackColor = true;
      this.btnZoomOut4.Visible = false;
      this.btnZoomOut4.Click += new EventHandler(this.btnZoomOut4_Click);
      this.btnCompareCalc.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.btnCompareCalc.Location = new Point(498, 210);
      this.btnCompareCalc.Name = "btnCompareCalc";
      this.btnCompareCalc.Size = new Size(200, 79);
      this.btnCompareCalc.TabIndex = 55;
      this.btnCompareCalc.Text = "Compare calculations";
      this.btnCompareCalc.UseVisualStyleBackColor = true;
      this.btnCompareCalc.Click += new EventHandler(this.btnCompareCalc_Click);
      this.pictureBox4.Location = new Point(713, 531);
      this.pictureBox4.Name = "pictureBox4";
      this.pictureBox4.Size = new Size(615, 466);
      this.pictureBox4.TabIndex = 54;
      this.pictureBox4.TabStop = false;
      this.pictureBox3.Location = new Point(47, 531);
      this.pictureBox3.Name = "pictureBox3";
      this.pictureBox3.Size = new Size(615, 466);
      this.pictureBox3.TabIndex = 53;
      this.pictureBox3.TabStop = false;
      this.panel24.AutoScroll = true;
      this.panel24.BackColor = Color.White;
      this.panel24.Location = new Point(23, 350);
      this.panel24.Name = "panel24";
      this.panel24.Size = new Size(1305, 158);
      this.panel24.TabIndex = 52;
      this.label68.AutoSize = true;
      this.label68.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.label68.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.label68.Location = new Point(44, 14);
      this.label68.Name = "label68";
      this.label68.Size = new Size(268, 24);
      this.label68.TabIndex = 51;
      this.label68.Text = "Historic Calculation parameters";
      this.label67.AutoSize = true;
      this.label67.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.label67.Font = new Font("Microsoft Sans Serif", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.label67.Location = new Point(712, 14);
      this.label67.Name = "label67";
      this.label67.Size = new Size(265, 24);
      this.label67.TabIndex = 50;
      this.label67.Text = "Current calculation parameters";
      this.panel23.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel23.Controls.Add((Control) this.label66);
      this.panel23.Location = new Point(1091, 323);
      this.panel23.Name = "panel23";
      this.panel23.Size = new Size(237, 21);
      this.panel23.TabIndex = 49;
      this.label66.AutoSize = true;
      this.label66.Location = new Point(3, 4);
      this.label66.Name = "label66";
      this.label66.Size = new Size(30, 13);
      this.label66.TabIndex = 8;
      this.label66.Text = "Note";
      this.panel22.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.panel22.Controls.Add((Control) this.label65);
      this.panel22.Location = new Point(877, 323);
      this.panel22.Name = "panel22";
      this.panel22.Size = new Size(248, 21);
      this.panel22.TabIndex = 48;
      this.label65.AutoSize = true;
      this.label65.Location = new Point(3, 4);
      this.label65.Name = "label65";
      this.label65.Size = new Size(30, 13);
      this.label65.TabIndex = 8;
      this.label65.Text = "Note";
      this.panel21.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel21.Controls.Add((Control) this.label64);
      this.panel21.Location = new Point(805, 323);
      this.panel21.Name = "panel21";
      this.panel21.Size = new Size(106, 21);
      this.panel21.TabIndex = 47;
      this.label64.AutoSize = true;
      this.label64.Location = new Point(3, 4);
      this.label64.Name = "label64";
      this.label64.Size = new Size(23, 13);
      this.label64.TabIndex = 2;
      this.label64.Text = "Rel";
      this.panel20.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.panel20.Controls.Add((Control) this.label63);
      this.panel20.Location = new Point(733, 323);
      this.panel20.Name = "panel20";
      this.panel20.Size = new Size(106, 21);
      this.panel20.TabIndex = 46;
      this.label63.AutoSize = true;
      this.label63.Location = new Point(3, 4);
      this.label63.Name = "label63";
      this.label63.Size = new Size(23, 13);
      this.label63.TabIndex = 2;
      this.label63.Text = "Rel";
      this.label63.Click += new EventHandler(this.label63_Click);
      this.panel19.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel19.Controls.Add((Control) this.label62);
      this.panel19.Location = new Point(628, 323);
      this.panel19.Name = "panel19";
      this.panel19.Size = new Size(139, 21);
      this.panel19.TabIndex = 45;
      this.label62.AutoSize = true;
      this.label62.Location = new Point(3, 4);
      this.label62.Name = "label62";
      this.label62.Size = new Size(32, 13);
      this.label62.TabIndex = 3;
      this.label62.Text = "Chart";
      this.panel18.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel18.Controls.Add((Control) this.label61);
      this.panel18.Location = new Point(541, 323);
      this.panel18.Name = "panel18";
      this.panel18.Size = new Size(121, 21);
      this.panel18.TabIndex = 44;
      this.label61.AutoSize = true;
      this.label61.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.label61.Location = new Point(3, 4);
      this.label61.Name = "label61";
      this.label61.Size = new Size(71, 13);
      this.label61.TabIndex = 7;
      this.label61.Text = "Fit statistically";
      this.panel14.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.panel14.Controls.Add((Control) this.label60);
      this.panel14.Location = new Point(454, 323);
      this.panel14.Name = "panel14";
      this.panel14.Size = new Size(121, 21);
      this.panel14.TabIndex = 43;
      this.label60.AutoSize = true;
      this.label60.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.label60.Location = new Point(3, 4);
      this.label60.Name = "label60";
      this.label60.Size = new Size(71, 13);
      this.label60.TabIndex = 7;
      this.label60.Text = "Fit statistically";
      this.panel9.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel9.Controls.Add((Control) this.label59);
      this.panel9.Location = new Point(301, 323);
      this.panel9.Name = "panel9";
      this.panel9.Size = new Size(187, 21);
      this.panel9.TabIndex = 42;
      this.label59.AutoSize = true;
      this.label59.Location = new Point(3, 4);
      this.label59.Name = "label59";
      this.label59.Size = new Size(60, 13);
      this.label59.TabIndex = 6;
      this.label59.Text = "KPI (count)";
      this.panel7.BackColor = Color.FromArgb((int) byte.MaxValue, 224, 192);
      this.panel7.Controls.Add((Control) this.label56);
      this.panel7.Location = new Point(148, 323);
      this.panel7.Name = "panel7";
      this.panel7.Size = new Size(187, 21);
      this.panel7.TabIndex = 40;
      this.label56.AutoSize = true;
      this.label56.Location = new Point(3, 4);
      this.label56.Name = "label56";
      this.label56.Size = new Size(60, 13);
      this.label56.TabIndex = 6;
      this.label56.Text = "KPI (count)";
      this.panel8.BackColor = Color.FromArgb(192, 192, (int) byte.MaxValue);
      this.panel8.Controls.Add((Control) this.label57);
      this.panel8.Location = new Point(23, 323);
      this.panel8.Name = "panel8";
      this.panel8.Size = new Size(124, 21);
      this.panel8.TabIndex = 41;
      this.label57.AutoSize = true;
      this.label57.Location = new Point(3, 4);
      this.label57.Name = "label57";
      this.label57.Size = new Size(64, 13);
      this.label57.TabIndex = 7;
      this.label57.Text = "VIRT_OZID";
      this.groupBox3.Controls.Add((Control) this.label72);
      this.groupBox3.Controls.Add((Control) this.label73);
      this.groupBox3.Controls.Add((Control) this.chkVirtOzid4);
      this.groupBox3.Controls.Add((Control) this.dtCalcDateTime4);
      this.groupBox3.Controls.Add((Control) this.cmbProdID4);
      this.groupBox3.Controls.Add((Control) this.cmbCalcID4);
      this.groupBox3.Controls.Add((Control) this.txtNote4);
      this.groupBox3.Controls.Add((Control) this.label52);
      this.groupBox3.Controls.Add((Control) this.label53);
      this.groupBox3.Controls.Add((Control) this.label54);
      this.groupBox3.Controls.Add((Control) this.label55);
      this.groupBox3.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.groupBox3.Location = new Point(716, 52);
      this.groupBox3.Name = "groupBox3";
      this.groupBox3.Size = new Size(445, 246);
      this.groupBox3.TabIndex = 2;
      this.groupBox3.TabStop = false;
      this.groupBox3.Text = " ";
      this.label72.AutoSize = true;
      this.label72.Location = new Point(312, 21);
      this.label72.Name = "label72";
      this.label72.Size = new Size(19, 21);
      this.label72.TabIndex = 64;
      this.label72.Text = "3";
      this.label73.AutoSize = true;
      this.label73.Location = new Point(312, 53);
      this.label73.Name = "label73";
      this.label73.Size = new Size(19, 21);
      this.label73.TabIndex = 65;
      this.label73.Text = "4";
      this.chkVirtOzid4.CheckOnClick = true;
      this.chkVirtOzid4.FormattingEnabled = true;
      this.chkVirtOzid4.Location = new Point(147, 117);
      this.chkVirtOzid4.Name = "chkVirtOzid4";
      this.chkVirtOzid4.Size = new Size(281, 76);
      this.chkVirtOzid4.TabIndex = 16;
      this.chkVirtOzid4.SelectedIndexChanged += new EventHandler(this.checkedListBox2_SelectedIndexChanged);
      this.dtCalcDateTime4.CalendarFont = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.dtCalcDateTime4.Enabled = false;
      this.dtCalcDateTime4.Location = new Point(147, 81);
      this.dtCalcDateTime4.Name = "dtCalcDateTime4";
      this.dtCalcDateTime4.Size = new Size(281, 29);
      this.dtCalcDateTime4.TabIndex = 15;
      this.cmbProdID4.Font = new Font("Segoe UI", 12f);
      this.cmbProdID4.FormattingEnabled = true;
      this.cmbProdID4.Location = new Point(147, 18);
      this.cmbProdID4.Name = "cmbProdID4";
      this.cmbProdID4.Size = new Size(159, 29);
      this.cmbProdID4.TabIndex = 13;
      this.cmbProdID4.SelectedIndexChanged += new EventHandler(this.cmbProdID4_SelectedIndexChanged);
      this.cmbCalcID4.Font = new Font("Segoe UI", 12f);
      this.cmbCalcID4.FormattingEnabled = true;
      this.cmbCalcID4.Location = new Point(147, 50);
      this.cmbCalcID4.Name = "cmbCalcID4";
      this.cmbCalcID4.Size = new Size(159, 29);
      this.cmbCalcID4.TabIndex = 10;
      this.cmbCalcID4.SelectedIndexChanged += new EventHandler(this.cmbCalcID4_SelectedIndexChanged);
      this.txtNote4.Font = new Font("Segoe UI", 12f);
      this.txtNote4.Location = new Point(147, 199);
      this.txtNote4.Multiline = true;
      this.txtNote4.Name = "txtNote4";
      this.txtNote4.Size = new Size(281, 41);
      this.txtNote4.TabIndex = 8;
      this.label52.AutoSize = true;
      this.label52.Font = new Font("Segoe UI", 12f);
      this.label52.Location = new Point(3, 117);
      this.label52.Name = "label52";
      this.label52.Size = new Size(74, 21);
      this.label52.TabIndex = 3;
      this.label52.Text = "Virt_Ozid";
      this.label53.AutoSize = true;
      this.label53.Font = new Font("Segoe UI", 12f);
      this.label53.Location = new Point(3, 80);
      this.label53.Name = "label53";
      this.label53.Size = new Size(142, 21);
      this.label53.TabIndex = 2;
      this.label53.Text = "Date of calculation ";
      this.label54.AutoSize = true;
      this.label54.Font = new Font("Segoe UI", 12f);
      this.label54.Location = new Point(3, 58);
      this.label54.Name = "label54";
      this.label54.Size = new Size(106, 21);
      this.label54.TabIndex = 1;
      this.label54.Text = "Calculation ID";
      this.label55.AutoSize = true;
      this.label55.Font = new Font("Segoe UI", 12f);
      this.label55.Location = new Point(3, 26);
      this.label55.Name = "label55";
      this.label55.Size = new Size(97, 21);
      this.label55.TabIndex = 0;
      this.label55.Text = "Productcode";
      this.groupBox2.Controls.Add((Control) this.chkVirtOzid3);
      this.groupBox2.Controls.Add((Control) this.label70);
      this.groupBox2.Controls.Add((Control) this.label69);
      this.groupBox2.Controls.Add((Control) this.dtCalcDateTime3);
      this.groupBox2.Controls.Add((Control) this.cmbProdID3);
      this.groupBox2.Controls.Add((Control) this.cmbCalcID3);
      this.groupBox2.Controls.Add((Control) this.txtNote3);
      this.groupBox2.Controls.Add((Control) this.label43);
      this.groupBox2.Controls.Add((Control) this.label49);
      this.groupBox2.Controls.Add((Control) this.label50);
      this.groupBox2.Controls.Add((Control) this.label51);
      this.groupBox2.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.groupBox2.Location = new Point(47, 43);
      this.groupBox2.Name = "groupBox2";
      this.groupBox2.Size = new Size(445, 246);
      this.groupBox2.TabIndex = 1;
      this.groupBox2.TabStop = false;
      this.groupBox2.Text = " ";
      this.chkVirtOzid3.CheckOnClick = true;
      this.chkVirtOzid3.FormattingEnabled = true;
      this.chkVirtOzid3.Location = new Point(147, 117);
      this.chkVirtOzid3.Name = "chkVirtOzid3";
      this.chkVirtOzid3.Size = new Size(281, 76);
      this.chkVirtOzid3.TabIndex = 16;
      this.label70.AutoSize = true;
      this.label70.Location = new Point(312, 53);
      this.label70.Name = "label70";
      this.label70.Size = new Size(19, 21);
      this.label70.TabIndex = 63;
      this.label70.Text = "2";
      this.label69.AutoSize = true;
      this.label69.Location = new Point(312, 18);
      this.label69.Name = "label69";
      this.label69.Size = new Size(19, 21);
      this.label69.TabIndex = 62;
      this.label69.Text = "1";
      this.dtCalcDateTime3.CalendarFont = new Font("Segoe UI", 12f, FontStyle.Bold);
      this.dtCalcDateTime3.Enabled = false;
      this.dtCalcDateTime3.Location = new Point(147, 81);
      this.dtCalcDateTime3.Name = "dtCalcDateTime3";
      this.dtCalcDateTime3.Size = new Size(281, 29);
      this.dtCalcDateTime3.TabIndex = 15;
      this.cmbProdID3.Font = new Font("Segoe UI", 12f);
      this.cmbProdID3.FormattingEnabled = true;
      this.cmbProdID3.Location = new Point(147, 18);
      this.cmbProdID3.Name = "cmbProdID3";
      this.cmbProdID3.Size = new Size(159, 29);
      this.cmbProdID3.TabIndex = 13;
      this.cmbProdID3.SelectedIndexChanged += new EventHandler(this.cmbProdID3_SelectedIndexChanged);
      this.cmbCalcID3.Font = new Font("Segoe UI", 12f);
      this.cmbCalcID3.FormattingEnabled = true;
      this.cmbCalcID3.Location = new Point(147, 50);
      this.cmbCalcID3.Name = "cmbCalcID3";
      this.cmbCalcID3.Size = new Size(159, 29);
      this.cmbCalcID3.TabIndex = 10;
      this.cmbCalcID3.SelectedIndexChanged += new EventHandler(this.cmbCalcID3_SelectedIndexChanged);
      this.txtNote3.Font = new Font("Segoe UI", 12f);
      this.txtNote3.Location = new Point(147, 199);
      this.txtNote3.Multiline = true;
      this.txtNote3.Name = "txtNote3";
      this.txtNote3.Size = new Size(281, 41);
      this.txtNote3.TabIndex = 8;
      this.label43.AutoSize = true;
      this.label43.Font = new Font("Segoe UI", 12f);
      this.label43.Location = new Point(3, 117);
      this.label43.Name = "label43";
      this.label43.Size = new Size(74, 21);
      this.label43.TabIndex = 3;
      this.label43.Text = "Virt_Ozid";
      this.label49.AutoSize = true;
      this.label49.Font = new Font("Segoe UI", 12f);
      this.label49.Location = new Point(3, 80);
      this.label49.Name = "label49";
      this.label49.Size = new Size(142, 21);
      this.label49.TabIndex = 2;
      this.label49.Text = "Date of calculation ";
      this.label50.AutoSize = true;
      this.label50.Font = new Font("Segoe UI", 12f);
      this.label50.Location = new Point(3, 58);
      this.label50.Name = "label50";
      this.label50.Size = new Size(106, 21);
      this.label50.TabIndex = 1;
      this.label50.Text = "Calculation ID";
      this.label51.AutoSize = true;
      this.label51.Font = new Font("Segoe UI", 12f);
      this.label51.Location = new Point(3, 26);
      this.label51.Name = "label51";
      this.label51.Size = new Size(97, 21);
      this.label51.TabIndex = 0;
      this.label51.Text = "Productcode";
      this.openFileDialog1.FileName = "openFileDialog1";
      this.timer1.Enabled = true;
      this.timer1.Interval = 1000;
      this.timer1.Tick += new EventHandler(this.timer1_Tick_1);
      this.printDialog1.UseEXDialog = true;
      this.AutoScaleDimensions = new SizeF(6f, 13f);
      this.AutoScaleMode = AutoScaleMode.Font;
      this.ClientSize = new Size(1446, 1051);
      this.Controls.Add((Control) this.tabMain);
      this.Name = nameof (Form1);
      this.Text = " BI CPV tool";
      this.Load += new EventHandler(this.Form1_Load);
      this.Resize += new EventHandler(this.Form1_Resize);
      this.tabMain.ResumeLayout(false);
      this.Main.ResumeLayout(false);
      this.Main.PerformLayout();
      ((ISupportInitialize) this.pictureBox1).EndInit();
      ((ISupportInitialize) this.dataGridViewTemp).EndInit();
      ((ISupportInitialize) this.dataGridViewRaw).EndInit();
      this.ViewCalculation.ResumeLayout(false);
      this.ViewCalculation.PerformLayout();
      this.panel2.ResumeLayout(false);
      this.panel2.PerformLayout();
      ((ISupportInitialize) this.dataGridView2).EndInit();
      ((ISupportInitialize) this.picGraph).EndInit();
      ((ISupportInitialize) this.dataGridView1).EndInit();
      this.SearchCalculation.ResumeLayout(false);
      this.SearchCalculation.PerformLayout();
      ((ISupportInitialize) this.pictureBox2).EndInit();
      this.panel4.ResumeLayout(false);
      this.panel4.PerformLayout();
      this.panel17.ResumeLayout(false);
      this.panel17.PerformLayout();
      this.panel16.ResumeLayout(false);
      this.panel16.PerformLayout();
      this.panel15.ResumeLayout(false);
      this.panel15.PerformLayout();
      this.panel13.ResumeLayout(false);
      this.panel13.PerformLayout();
      this.panel12.ResumeLayout(false);
      this.panel12.PerformLayout();
      this.panel11.ResumeLayout(false);
      this.panel11.PerformLayout();
      this.panel10.ResumeLayout(false);
      this.panel10.PerformLayout();
      this.panel5.ResumeLayout(false);
      this.panel5.PerformLayout();
      this.panelSelection.ResumeLayout(false);
      this.panelSelection.PerformLayout();
      this.panel3.ResumeLayout(false);
      this.panel3.PerformLayout();
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      this.groupFilterSelection.ResumeLayout(false);
      this.groupFilterSelection.PerformLayout();
      this.panelButtons.ResumeLayout(false);
      this.panelButtons.PerformLayout();
      this.CompareCalculation.ResumeLayout(false);
      this.CompareCalculation.PerformLayout();
      ((ISupportInitialize) this.pictureBox4).EndInit();
      ((ISupportInitialize) this.pictureBox3).EndInit();
      this.panel23.ResumeLayout(false);
      this.panel23.PerformLayout();
      this.panel22.ResumeLayout(false);
      this.panel22.PerformLayout();
      this.panel21.ResumeLayout(false);
      this.panel21.PerformLayout();
      this.panel20.ResumeLayout(false);
      this.panel20.PerformLayout();
      this.panel19.ResumeLayout(false);
      this.panel19.PerformLayout();
      this.panel18.ResumeLayout(false);
      this.panel18.PerformLayout();
      this.panel14.ResumeLayout(false);
      this.panel14.PerformLayout();
      this.panel9.ResumeLayout(false);
      this.panel9.PerformLayout();
      this.panel7.ResumeLayout(false);
      this.panel7.PerformLayout();
      this.panel8.ResumeLayout(false);
      this.panel8.PerformLayout();
      this.groupBox3.ResumeLayout(false);
      this.groupBox3.PerformLayout();
      this.groupBox2.ResumeLayout(false);
      this.groupBox2.PerformLayout();
      this.ResumeLayout(false);
    }
  }
}
