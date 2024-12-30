using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using RDotNet;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;


namespace BI_CPV_tool
{
// Interface for reading Excel files
public interface IExcelReader
{
    DataTable ReadExcel(string path);
}

// Interface for processing the data table
public interface IDataProcessor
{
    void ProcessData(DataTable dataTable);
}
    // Concrete implementation of IExcelReader
public class ExcelReader : IExcelReader
{
    public DataTable ReadExcel(string path)
    {
        try
        {
            using (Stream inputStream = File.OpenRead(path))
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    IWorkbook workbook = application.Workbooks.Open(inputStream);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    return worksheet.ExportDataTable(worksheet.UsedRange["A1:AA300000"], ExcelExportDataTableOptions.ColumnNames);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
            return null;
        }
    }
}
// Concrete implementation of IDataProcessor
public class DataProcessor : IDataProcessor
{
    public void ProcessData(DataTable dataTable)
    {
        if (dataTable.Columns[0].ColumnName == "PRODUKTCODE")
        {
            dataTable.Columns.Add("SORT_DATE1", typeof(DateTime)).SetOrdinal(1);
            dataTable.Columns.Add("TS_ABS1", typeof(DateTime)).SetOrdinal(2);
        }
        else
        {
            dataTable.Columns.Add("SORT_DATE1", typeof(DateTime)).SetOrdinal(2);
            dataTable.Columns.Add("TS_ABS1", typeof(DateTime)).SetOrdinal(3);
        }

        foreach (DataRow r1 in dataTable.Rows)
        {
            string f = r1["SORT_DATE"].ToString();
            string h = r1["TS_ABS"].ToString();

            if (f == "01.01.0001 00:00:00") f = r1["TS_ABS"].ToString();
            if (h == "01.01.0001 00:00:00") h = r1["TS_ABS"].ToString();

            r1["TS_ABS1"] = DateTime.Parse(h);
            r1["SORT_DATE1"] = DateTime.Parse(f);
        }

        dataTable.Columns.Remove("TS_ABS");
        dataTable.Columns.Remove("SORT_DATE");

        dataTable.Columns["TS_ABS1"].ColumnName = "TS_ABS";
        dataTable.Columns["SORT_DATE1"].ColumnName = "SORT_DATE";
    }
}
// High-level class that orchestrates reading and processing
public class ExcelService
{
    private readonly IExcelReader _excelReader;
    private readonly IDataProcessor _dataProcessor;

    public ExcelService(IExcelReader excelReader, IDataProcessor dataProcessor)
    {
        _excelReader = excelReader;
        _dataProcessor = dataProcessor;
    }

    public DataTable ReadAndProcessExcel(string path)
    {
        DataTable dataTable = _excelReader.ReadExcel(path);
        if (dataTable == null) return null;

        _dataProcessor.ProcessData(dataTable);
        return dataTable;
    }
}
    
public partial class Form1 : Form
{
    // R.net engine
    internal REngine engine; // R.net object working with R code libraries

    // UI-related variables
    private System.Drawing.Printing.PrintDocument docToPrint = new System.Drawing.Printing.PrintDocument();
    public static FormWindowState LastWindowState;

    // Connection and Database-related variables
    public static string connectionString = ""; // The string of connection to DB
    public static string strSQLfiltered = "";  // SQL query after data filter
    public static string strQuery = "";        // SQL query string
    public static string strQuery2 = "";       // Another SQL query string
    public static string strQuery3 = "";       // Third query string
    public static string strQuery4 = "";       // Fourth query string

    // Excel and Data Table variables
    public static string strExcelFileName = ""; // The name of Excel file after filtering data
    public static DataTable dt = null; // Data Table Object where input Excel file is exported
    public static string strFileName = ""; // Input Excel file name
    public static string[] arrDataGridView = new DataGridView[1000]; // Array of DataGridViews linked to the number of loaded input files
    public static int intCounterDataGridViews = 0; // Counter for the number of loaded input files

    // File path and directory-related variables
    public static string strRscript = ""; // First line from app.ini file - real MS SQL connection string
    public static string strRpath = ""; // Second line from app.ini file (not used anymore)
    public static string strDataDir = ""; // Third line from app.ini file - Input R code file
    public static string strOutputDir = ""; // Forth line from app.ini file - Output calculation folder
    public static string strCalcID = ""; // Fifth line from app.ini file - Output calculation folder with plots

    // Calculation and filter-related variables
    public static string strLaufNRDateMin = ""; // Minimal LaufNR
    public static string strLaufNRDateMax = ""; // Maximal LaufNR
    public static bool dateIsGood = false; // Check date on dd.mm.yyyy
    public static bool filterOn = false; // Verifying if the filter button was pressed
    public static string strFilterExclude = ""; // Excluding OZID filter (not used)
    public static string strHasID = ""; // The flag defining the first Excel file column "ID"
    public static string strFilterLaufnr = ""; // Used in filter
    public static string strOZID = ""; // Saving Virt_Ozid in calculation search

    // Flags and states
    public static bool filterIsOn = false; // The flag defining the filter is enabled
    public static bool blnFlag = false; // Boolean flag
    public static bool stopCalc = false; // Stop calculation flag
    public static bool blnPressed = false; // Used in plot reflection

    // Parameters for calculation and logic
    public static int last = 1; // Used in progress bar
    public static int intMaxPoints = 200000; // The number of max rows imported from input Excel file
    public static string[] strArrToolTip = new string[500];
    public static string[] strArg = { "", "", "" }; // Input arguments for R code file
    public static string[] strArr = { "" }; // Not used
    public static string[] strArr2 = new string[500]; // Used in calculation compare
    public static string[] arr1 = new string[500]; // Used in logic of calculation compare
    public static string[] arr2 = new string[500]; // Used in logic of calculation compare
    public static string[] arr3 = new string[500]; // Used in logic of calculation compare
    public static string[] strArrToolTip = new string[500]; // Tooltip array for UI elements

    // Calculation view and result variables
    public static int ID = 0; // Field in table CalcResultView
    public static int intIndex = 0; // Used just as 0
    public static string NParameterTotal = ""; // NParameterTotal parameter in Calculation View
    public static string NStatistically = ""; // NStatistically parameter in Calculation View
    public static string PercentStatistically = ""; // PercentStatistically parameter in Calculation View
    public static string DoNotFitStatistically = ""; // DoNotFitStatistically parameter in Calculation View
    public static string CalcID = ""; // CalcID field in many tables

    // Output and status-related variables
    public static string strOutPutPath = ""; // Output path
    public static string TimePointData = ""; // TimePointData field
    public static string TimePointCalc = ""; // TimePointCalc field
    public static string Note = ""; // Note field

    // Status tracking variables
    public static string Status_fit_statistically = ""; // Status_fit_statistically
    public static string Num_VIRT_OZID_not_fit_stat = ""; // Num_VIRT_OZID_not_fit_stat
    public static string Percent_of_values_status_KPI0_KPI3 = ""; // Percent_of_values_status_KPI0_KPI3
    public static string Num_of_values_status_KPI0_KPI3 = ""; // Num_of_values_status_KPI0_KPI3
    public static string RelevantForDiscussion = ""; // Relevant For Discussion flag
    public static string KPI0 = ""; // KPI0 field in table VIRT_OZID_per_calculation
    public static string KPI1 = ""; // KPI1 field in table VIRT_OZID_per_calculation
    public static string KPI2 = ""; // KPI2 field in table VIRT_OZID_per_calculation
    public static string KPI3 = ""; // KPI3 field in table VIRT_OZID_per_calculation
    public static string GraphID = ""; // GraphID in table VIRT_OZID_per_calculation

    // Miscellaneous
    public static string strInitialDate = ""; // "from" date
    public static string strEndDate = ""; // "to" date
    public static string Additional_note = ""; // Field in table VIRT_OZID_per_calculation
    public static int[,] intParameters = new int[200, 3]; // Array of control sizes and coordinates in minimal form
    public string[,] strMatrix = new string[500, 15]; // The matrix contains dynamic rows of data  
    public static int strRowsCount = 0; // Quantity of rows in dynamic data block
    public static string strFilter = ""; // The variable of filter conditions

    public Boolean IsInteger(string strNum) // Defining if the string is integer
    {
        return int.TryParse(strNum, out _);
    }
}

       
        
          


        private static void DirectoryCopy(
        string sourceDirName, string destDirName, bool copySubDirs) // Copying folders to other places
        {
            try { 
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }


            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = Path.Combine(destDirName, file.Name);

                // Copy the file.
                file.CopyTo(temppath, false);
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {

                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
}
        public string Left(string value, int maxLength) // left signs in string
        {
            if (string.IsNullOrEmpty(value)) return value;
            maxLength = Math.Abs(maxLength);

            return (value.Length <= maxLength
                   ? value
                   : value.Substring(0, maxLength)
                   );
        }



        public void CopyToSQLUniversal(DataGridView dtGrid, string[] arr, string TableName) // bulk copy to sql table
        {
            try { 
            string connection = connectionString;
            using (var conn = new SqlConnection(connection))
            {
                conn.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
                {

                    for (int j = 0; j < arr.Count(); ++j)
                        bulkCopy.ColumnMappings.Add(j, arr[j]);



                    bulkCopy.BatchSize = 800000;
                    bulkCopy.DestinationTableName = TableName;
                    bulkCopy.BulkCopyTimeout = 600;
                    bulkCopy.WriteToServer(dt.CreateDataReader());
                }
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


       



        public void CopyToSQL(DataGridView dtGrid) //bulk copy to sql table
        {
            try { 
            DataTable dt = (DataTable)dtGrid.DataSource;
            string connection = connectionString;
            using (var conn = new SqlConnection(connection))
            {
                conn.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
                {

                    bulkCopy.ColumnMappings.Add(0, "ID");
                    bulkCopy.ColumnMappings.Add(1, "PRODUKTCODE");
                    bulkCopy.ColumnMappings.Add(2, "SORT_DATE");
                    bulkCopy.ColumnMappings.Add(3, "TS_ABS");
                    bulkCopy.ColumnMappings.Add(4, "LAUFNR");
                    bulkCopy.ColumnMappings.Add(5, "CHNR_ENDPRODUKT");
                    bulkCopy.ColumnMappings.Add(6, "PROCESS_CODE");
                    bulkCopy.ColumnMappings.Add(7, "PROCESS_CODE_NAME");
                    bulkCopy.ColumnMappings.Add(8, "PARAMETER_NAME");
                    bulkCopy.ColumnMappings.Add(9, "ASSAY");
                    bulkCopy.ColumnMappings.Add(10, "VIRT_OZID");
                    bulkCopy.ColumnMappings.Add(11, "TREND_WERT");
                    bulkCopy.ColumnMappings.Add(12, "TREND_WERT_2");
                    bulkCopy.ColumnMappings.Add(13, "ISTWERT_LIMS");
                    bulkCopy.ColumnMappings.Add(14, "LCL");
                    bulkCopy.ColumnMappings.Add(15, "UCL");
                    bulkCopy.ColumnMappings.Add(16, "CL");
                    bulkCopy.ColumnMappings.Add(17, "UAL");
                    bulkCopy.ColumnMappings.Add(18, "LAL");
                    bulkCopy.ColumnMappings.Add(19, "DECIMAL_PLACES_XCL_SUBSTITUTED");
                    bulkCopy.ColumnMappings.Add(20, "DECIMAL_PLACES_AL");
                    bulkCopy.ColumnMappings.Add(21, "DATA_TYPE");
                    bulkCopy.ColumnMappings.Add(22, "SOURCE_SYSTEM");
                    bulkCopy.ColumnMappings.Add(23, "EXCURSION");
                    bulkCopy.ColumnMappings.Add(24, "REFERENCED_CPV");
                    bulkCopy.ColumnMappings.Add(25, "IS_IN_RUN_NUMBER_RANGE");
                    bulkCopy.ColumnMappings.Add(26, "LOCATION");

                    bulkCopy.BatchSize = 800000;
                    bulkCopy.DestinationTableName = "Products";
                    bulkCopy.BulkCopyTimeout = 600;
                    bulkCopy.WriteToServer(dt.CreateDataReader());
                }
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void CopyToSQL2(DataGridView dtGrid)//bulk copy to sql table
        {
            try { 
            DataTable dt = (DataTable)dtGrid.DataSource;
            string connection = connectionString;
            using (var conn = new SqlConnection(connection))
            {
                conn.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
                {


                    bulkCopy.ColumnMappings.Add(0, "PRODUKTCODE");
                    bulkCopy.ColumnMappings.Add(1, "SORT_DATE");
                    bulkCopy.ColumnMappings.Add(2, "TS_ABS");
                    bulkCopy.ColumnMappings.Add(3, "LAUFNR");
                    bulkCopy.ColumnMappings.Add(4, "CHNR_ENDPRODUKT");
                    bulkCopy.ColumnMappings.Add(5, "PROCESS_CODE");
                    bulkCopy.ColumnMappings.Add(6, "PROCESS_CODE_NAME");
                    bulkCopy.ColumnMappings.Add(7, "PARAMETER_NAME");
                    bulkCopy.ColumnMappings.Add(8, "ASSAY");
                    bulkCopy.ColumnMappings.Add(9, "VIRT_OZID");
                    bulkCopy.ColumnMappings.Add(10, "TREND_WERT");
                    bulkCopy.ColumnMappings.Add(11, "TREND_WERT_2");
                    bulkCopy.ColumnMappings.Add(12, "ISTWERT_LIMS");
                    bulkCopy.ColumnMappings.Add(13, "LCL");
                    bulkCopy.ColumnMappings.Add(14, "UCL");
                    bulkCopy.ColumnMappings.Add(15, "CL");
                    bulkCopy.ColumnMappings.Add(16, "UAL");
                    bulkCopy.ColumnMappings.Add(17, "LAL");
                    bulkCopy.ColumnMappings.Add(18, "DECIMAL_PLACES_XCL_SUBSTITUTED");
                    bulkCopy.ColumnMappings.Add(19, "DECIMAL_PLACES_AL");
                    bulkCopy.ColumnMappings.Add(20, "DATA_TYPE");
                    bulkCopy.ColumnMappings.Add(21, "SOURCE_SYSTEM");
                    bulkCopy.ColumnMappings.Add(22, "EXCURSION");
                    bulkCopy.ColumnMappings.Add(23, "REFERENCED_CPV");
                    bulkCopy.ColumnMappings.Add(24, "IS_IN_RUN_NUMBER_RANGE");
                    bulkCopy.ColumnMappings.Add(25, "LOCATION");

                    bulkCopy.BatchSize = 800000;
                    bulkCopy.DestinationTableName = "Products";
                    bulkCopy.BulkCopyTimeout = 600;
                        DataColumnCollection columns = dt.Columns;
                        if (columns.Contains("Column1"))
                        {
                           dt.Columns.Remove("Column1");
                        }
                        if (columns.Contains("Column2"))
                        {
                            dt.Columns.Remove("Column2");
                        }
                        if (columns.Contains("ID"))
                        {
                            dt.Columns.Remove("ID");
                        }
                        bulkCopy.WriteToServer(dt.CreateDataReader());
                }
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void timer1_Tick(object sender, EventArgs e) //ProgressBar action
        {
            try { 
            if (stopCalc) 
            { 
            if (txProgressBar.Text.Length == 100)
            {

                //timer1.Stop();
                txProgressBar.Visible = false;
                txProgressBar.Text =  "";
                //MessageBox.Show("Counted!");

            }
            else
            {
                txProgressBar.Visible = true;
                if (last == 1)
                {
                    txProgressBar.Text = txProgressBar.Text + "█";
                    last = 2;
                }
                else
                {
                    txProgressBar.Text = txProgressBar.Text + "█";
                    last = 1;
                }
            }
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void strH()//insert data to calculation tables used in Calculation search and calculation form save
        {
            try { 
            var strOutputDirMod = strOutputDir.Replace("/", "\\");
            var strFolderPath = @strOutputDirMod + strCalcID + "\\";
            string FolderPath = @strOutputDirMod + strCalcID + "\\";



            DirectoryInfo di = new DirectoryInfo(FolderPath);

            //Get All xlsx Files  
            List<string> getAllCSVFiles = new List<string>();
            getAllCSVFiles = di.GetFiles("*.xlsx")
                                 .Where(file => file.Name.EndsWith(".xlsx"))
                                 .Select(file => file.Name).ToList();

            int count = getAllCSVFiles.Count();
            for (int j = 0; j < count; ++j)
            {
                if (getAllCSVFiles[j].ToString().Length < 35)
                    listBox2.Items.Add(getAllCSVFiles[j].ToString());
            }


            string[] INIfolderPath;


            string dir = System.IO.Directory.GetCurrentDirectory();

            var strFile = strOutputDirMod + strCalcID + "\\" + getAllCSVFiles[0].ToString();

        IExcelReader excelReader = new ExcelReader();
        IDataProcessor dataProcessor = new DataProcessor();
        ExcelService excelService = new ExcelService(excelReader, dataProcessor);

        
        dt = excelService.ReadAndProcessExcel(strFile);
            
            //if (dt.Columns.Count > 24) { 
            //dt.Columns.RemoveAt(25); dt.Columns.RemoveAt(26);
            //}
            string[] strFields = { "PRODUKTCODE",
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
                "lag", "Column1","Column2" };
            //PRODUKTCODE     TREND_WERT  TREND_WERT_2   CL    LCL    UCL    LAL    UAL    TS_ABS     SORT_DATE   LAUFNR    EXCURSION    VIRT_OZID    VALUE    BatchID    lowSD    uppSD    mu    sigma    upp    delta    rSigma    valid    signal    lag

            var strTableName = "CalculationResult";

            //dataGridViewTemp.DataSource = dataGridViewRaw.DataSource;
            dataGridViewTemp.Columns.Clear();
            dataGridViewTemp.Columns.Add("PRODUKTCODE", "");
            dataGridViewTemp.Columns.Add("TREND_WERT", "");
            dataGridViewTemp.Columns.Add("TREND_WERT_2", "");
            dataGridViewTemp.Columns.Add("CL", "");
            dataGridViewTemp.Columns.Add("LCL", "");
            dataGridViewTemp.Columns.Add("UCL", "");
            dataGridViewTemp.Columns.Add("LAL", "");
            dataGridViewTemp.Columns.Add("UAL", "");
            dataGridViewTemp.Columns.Add("TS_ABS", "");
            dataGridViewTemp.Columns.Add("SORT_DATE", "");
            dataGridViewTemp.Columns.Add("LAUFNR", "");
            dataGridViewTemp.Columns.Add("EXCURSION", "");
            dataGridViewTemp.Columns.Add("VIRT_OZID", "");
            dataGridViewTemp.Columns.Add("VALUE", "");
            dataGridViewTemp.Columns.Add("BatchID", "");
            dataGridViewTemp.Columns.Add("lowSD", "");
            dataGridViewTemp.Columns.Add("uppSD", "");
            dataGridViewTemp.Columns.Add("mu", "");
            dataGridViewTemp.Columns.Add("sigma", "");
            dataGridViewTemp.Columns.Add("upp", "");
            dataGridViewTemp.Columns.Add("delta", "");
            dataGridViewTemp.Columns.Add("rSigma", "");
            dataGridViewTemp.Columns.Add("valid", "");
            dataGridViewTemp.Columns.Add("signal", "");
            dataGridViewTemp.Columns.Add("lag", "");
            dataGridViewTemp.Columns.Add("Column1", "");
            dataGridViewTemp.Columns.Add("Column2", "");

            //CopyDataTableToSQL(dataGridViewTemp, strFields, strTableName);
            CopyToSQLUniversal(dataGridViewTemp, strFields, strTableName);

            SqlConnection CN = new SqlConnection(connectionString);
            CN.Open();

            DateTime utcDate = DateTime.UtcNow;

            string qry = "";
            qry = "insert into CalculationRaw select '" + strCalcID + "','" + utcDate.ToString() + "'," + " [PRODUKTCODE],[TREND_WERT],[TREND_WERT_2],[CL],[LCL],[UCL],[LAL],[UAL],[TS_ABS],[SORT_DATE],[LAUFNR],[EXCURSION],[VIRT_OZID],[VALUE],[BatchID],[lowSD],[uppSD],[mu],[sigma],[upp],[delta],[rSigma],[valid],[signal],[lag] from CalculationResult";


            SqlCommand cmd = new SqlCommand(qry, CN);
            cmd = new SqlCommand(qry, CN);
            cmd.ExecuteNonQuery();

            qry = "insert into[dbo].[CalcRow] SELECT distinct  calcid,VIRT_OZID, count(signal) as totaln,dbo.PercentStatisticallyFit0_Ozid(calcid, VIRT_OZID) as procent0, dbo.PercentStatisticallyFit1_Ozid(calcid, VIRT_OZID) as procent1,dbo.PercentStatisticallyFit2_Ozid(calcid, VIRT_OZID) as procent2, dbo.PercentStatisticallyFit3_Ozid(calcid, VIRT_OZID) as procent3, 'False',  'False',''  FROM [dbo].[CalculationRaw]  where calcid = '" + strCalcID + "' group by VIRT_OZID, calcid";

            cmd = new SqlCommand(qry, CN);
            cmd.ExecuteNonQuery();
            qry = "insert into[dbo].[CalcRowSearch] SELECT distinct  calcid,VIRT_OZID, dbo.KPIcount0(calcid, VIRT_OZID) as totaln0,dbo.KPIcount1(calcid, VIRT_OZID) as totaln1,dbo.KPIcount2(calcid, VIRT_OZID) as totaln2,dbo.KPIcount3(calcid, VIRT_OZID) as totaln3,dbo.PercentStatisticallyFit0_Ozid(calcid, VIRT_OZID) as procent0, dbo.PercentStatisticallyFit1_Ozid(calcid, VIRT_OZID) as procent1,dbo.PercentStatisticallyFit2_Ozid(calcid, VIRT_OZID) as procent2, dbo.PercentStatisticallyFit3_Ozid(calcid, VIRT_OZID) as procent3, 'False',  'False','False',''   FROM[dbo].[CalculationRaw]  where calcid = '" + strCalcID + "' group by VIRT_OZID, calcid";
            cmd = new SqlCommand(qry, CN);
            cmd.ExecuteNonQuery();
            CN.Close();
            CultureInfo culture = new CultureInfo("de-DE");
            
            stopCalc = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }


        public static byte[] GetPhoto(string filePath) // get image for sql table Graphs
        {
            try { 
            FileStream stream = new FileStream(
                filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            byte[] photo = reader.ReadBytes((int)stream.Length);

            reader.Close();
            stream.Close();

            return photo;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        
        
        private void btnCalculation_Click(object sender, EventArgs e) //Action after starting calculation
        {
            try { 
            //var lastCulture = Thread.CurrentThread.CurrentCulture;
            //Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            DialogResult dialogResult = MessageBox.Show("Are you sure to start calculation Yes/No", "", MessageBoxButtons.YesNo);
            
            if (dialogResult == DialogResult.Yes)
            {
                       
            try
            {
                    txProgressBar.Visible = true;
                    txProgressBar.Text = "";
                    timer1.Tick += new EventHandler(timer1_Tick);
                    
                    timer1.Start();
                    backgroundWorker2.RunWorkerAsync();

                // SQL connection
                    txtResult.Visible = false;
                txtResult.Text = "";
                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();
                // Current date and time
                DateTime utcDate = DateTime.UtcNow;
                //Delete content of temporary calculation table
                string qry = "";
                qry = "delete from CalculationResult";
                SqlCommand cmd = new SqlCommand(qry, CN);
                cmd.ExecuteNonQuery();
                CN.Close();

                var rpath = strRscript;
                var scriptpath = @strRscript;



                string file = strRpath;
                string result = string.Empty;

                try
                {
                    if (strExcelFileName != null)
                    {
                        var info = new ProcessStartInfo();
                        info.FileName = rpath;
                        info.WorkingDirectory = Path.GetDirectoryName(file);

                        utcDate = DateTime.UtcNow;
                        var strDate =        utcDate.ToString("yyyy-MM-dd'T'HH:mm:ss",CultureInfo.InvariantCulture);

                                strCalcID = strDate.ToString().Replace("_", "").Replace(" ", "").Replace("-", "").Replace(".", "").Replace(":", "");
                         

                        //The below code will create a folder if the folder is not exists in C#.Net.            
                        var folderPath = strOutputDir + strCalcID;
                        DirectoryInfo folder = Directory.CreateDirectory(folderPath);

                        string[] args = { strExcelFileName.Replace(@"//", "/"), cmbProductCode.Text, folderPath + "/" };



                        args[0] = @strExcelFileName.Replace(@"//", "/");
                        //Pass parameters to R code
                        var subString = @scriptpath.Replace(@"\\", @"\");
                        info.Arguments = file + " " + args[0] + " " + args[1] + " " + args[2];




                        using (StreamWriter writer = new StreamWriter(file, true))
                        {

                            strFileName = openFileDialog1.FileName;

                            {

                                btnOpenFile.BackColor = Color.LightBlue;
                                String name = "Items";
                                DataTable dtRanges = new DataTable();
                                DataTable dtTable = new DataTable();
                                dt = (DataTable)dataGridViewRaw.DataSource;
                                dynamic dWorkSheet = null;
                                IExcelReader excelReader = new ExcelReader();
                                IDataProcessor dataProcessor = new DataProcessor();
                                ExcelService excelService = new ExcelService(excelReader, dataProcessor);

        
                                dt = excelService.ReadAndProcessExcel(strExcelFileName.Replace("//", "/"));
                               
                                dt.Columns.Add("UserID"); dt.Columns.Add("ModifiedDate"); dt.Columns.Add("GraphID"); dt.Columns.Add("CalcID"); dt.Columns.Add("FilterID");
                            }

                        }
                        string[] strFields = {"ID","PRODUKTCODE","SORT_DATE","TS_ABS","LAUFNR","CHNR_ENDPRODUKT","PROCESS_CODE","PROCESS_CODE_NAME","PARAMETER_NAME","ASSAY",
                          "VIRT_OZID","TREND_WERT","TREND_WERT_2","ISTWERT_LIMS","LCL","UCL","CL","UAL","LAL","DECIMAL_PLACES_XCL_SUBSTITUTED","DECIMAL_PLACES_AL" ,
                          "DATA_TYPE","SOURCE_SYSTEM","EXCURSION" ,"REFERENCED_CPV","IS_IN_RUN_NUMBER_RANGE","LOCATION","UserID","ModifiedDate","GraphID","CalcID","FilterID"  };


                        //if (strHasID == "1") { 
                        //strFields = {"ID","PRODUKTCODE","SORT_DATE","TS_ABS","LAUFNR","CHNR_ENDPRODUKT","PROCESS_CODE","PROCESS_CODE_NAME","PARAMETER_NAME","ASSAY",
                        //  "VIRT_OZID","TREND_WERT","TREND_WERT_2","ISTWERT_LIMS","LCL","UCL","CL","UAL","LAL","DECIMAL_PLACES_XCL_SUBSTITUTED","DECIMAL_PLACES_AL" ,
                        //  "DATA_TYPE","SOURCE_SYSTEM","EXCURSION" ,"REFERENCED_CPV","IS_IN_RUN_NUMBER_RANGE","LOCATION","UserID","ModifiedDate","GraphID","CalcID","FilterID"  };
                        //}
                        //if (strHasID == "0")
                        //{
                        //  strFields = {"PRODUKTCODE","SORT_DATE","TS_ABS","LAUFNR","CHNR_ENDPRODUKT","PROCESS_CODE","PROCESS_CODE_NAME","PARAMETER_NAME","ASSAY",
                        //  "VIRT_OZID","TREND_WERT","TREND_WERT_2","ISTWERT_LIMS","LCL","UCL","CL","UAL","LAL","DECIMAL_PLACES_XCL_SUBSTITUTED","DECIMAL_PLACES_AL" ,
                        //  "DATA_TYPE","SOURCE_SYSTEM","EXCURSION" ,"REFERENCED_CPV","IS_IN_RUN_NUMBER_RANGE","LOCATION","UserID","ModifiedDate","GraphID","CalcID","FilterID"  };
                        //}
                        var strTableName = "ProductsFilteredTemp";

                        //dataGridViewTemp.DataSource = dataGridViewRaw.DataSource;
                        if (strHasID == "1")
                        {
                            dataGridViewTemp.Columns.Add("ID", "");
                        }

                        dataGridViewTemp.Columns.Add("PRODUKTCODE", "");
                        dataGridViewTemp.Columns.Add("SORT_DATE", "");
                        dataGridViewTemp.Columns.Add("TS_ABS", "");
                        dataGridViewTemp.Columns.Add("LAUFNR", "");
                        dataGridViewTemp.Columns.Add("CHNR_ENDPRODUKT", "");
                        dataGridViewTemp.Columns.Add("PROCESS_CODE", "");
                        dataGridViewTemp.Columns.Add("PROCESS_CODE_NAME", "");
                        dataGridViewTemp.Columns.Add("PARAMETER_NAME", "");
                        dataGridViewTemp.Columns.Add("ASSAY", "");
                        dataGridViewTemp.Columns.Add("VIRT_OZID", "");
                        dataGridViewTemp.Columns.Add("TREND_WERT", "");
                        dataGridViewTemp.Columns.Add("TREND_WERT_2", "");
                        dataGridViewTemp.Columns.Add("ISTWERT_LIMS", "");
                        dataGridViewTemp.Columns.Add("LCL", "");
                        dataGridViewTemp.Columns.Add("UCL", "");
                        dataGridViewTemp.Columns.Add("CL", "");
                        dataGridViewTemp.Columns.Add("UAL", "");
                        dataGridViewTemp.Columns.Add("LAL", "");
                        dataGridViewTemp.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", "");
                        dataGridViewTemp.Columns.Add("DECIMAL_PLACES_AL", "");
                        dataGridViewTemp.Columns.Add("DATA_TYPE", "");
                        dataGridViewTemp.Columns.Add("SOURCE_SYSTEM", "");
                        dataGridViewTemp.Columns.Add("EXCURSION", "");
                        dataGridViewTemp.Columns.Add("REFERENCED_CPV", ""); 
                        dataGridViewTemp.Columns.Add("IS_IN_RUN_NUMBER_RANGE", "");
                        dataGridViewTemp.Columns.Add("LOCATION", "");
                        dataGridViewTemp.Columns.Add("UserID", "");
                        dataGridViewTemp.Columns.Add("ModifiedDate", "");
                        dataGridViewTemp.Columns.Add("GraphID", "");
                        dataGridViewTemp.Columns.Add("CalcID", strCalcID);
                        dataGridViewTemp.Columns.Add("FilterID", "");

                        CopyToSQLUniversal(dataGridViewTemp, strFields, strTableName);
                        SqlConnection connection = new SqlConnection(connectionString);
                        CN.Open();
                        connection.Open();
                        string query = "update ProductsFilteredTemp set CalcID = '" + strCalcID + "'";



                        SqlCommand command2 = new SqlCommand(query, connection);
                        var r2 = command2.ExecuteNonQuery();


                        var constr = connectionString;
                        SqlConnection connection2 = new SqlConnection(connectionString);



                        connection2.Open();

                        strTableName = "ProductsFiltered";
                        if (strHasID == "1")
                        {
                            string[] strFields1 = { "ID", "PRODUKTCODE", "SORT_DATE", "TS_ABS", "LAUFNR", "CHNR_ENDPRODUKT", "PROCESS_CODE", "PROCESS_CODE_NAME", "PARAMETER_NAME", "ASSAY", "VIRT_OZID", "TREND_WERT", "TREND_WERT_2", "ISTWERT_LIMS", "LCL", "UCL", "CL", "UAL", "LAL", "DECIMAL_PLACES_XCL_SUBSTITUTED", "DECIMAL_PLACES_AL", "DATA_TYPE", "SOURCE_SYSTEM", "EXCURSION", "REFERENCED_CPV", "IS_IN_RUN_NUMBER_RANGE", "LOCATION", "UserID", "ModifiedDate", "GraphID", "CalcID", "FilterID" };
                        }
                        if (strHasID == "0")
                        {
                            string[] strFields1 = { "PRODUKTCODE", "SORT_DATE", "TS_ABS", "LAUFNR", "CHNR_ENDPRODUKT", "PROCESS_CODE", "PROCESS_CODE_NAME", "PARAMETER_NAME", "ASSAY", "VIRT_OZID", "TREND_WERT", "TREND_WERT_2", "ISTWERT_LIMS", "LCL", "UCL", "CL", "UAL", "LAL", "DECIMAL_PLACES_XCL_SUBSTITUTED", "DECIMAL_PLACES_AL", "DATA_TYPE", "SOURCE_SYSTEM", "EXCURSION", "REFERENCED_CPV", "IS_IN_RUN_NUMBER_RANGE", "LOCATION", "UserID", "ModifiedDate", "GraphID", "CalcID", "FilterID" };
                        }
                        cmd = new SqlCommand("dataIn", connection2);
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.ExecuteNonQuery();
                        connection2.Close();



                        query = "delete from ProductsFilteredTemp ";

                        command2 = new SqlCommand(query, connection);
                        r2 = command2.ExecuteNonQuery();

                        info.RedirectStandardInput = false;
                        info.RedirectStandardOutput = true;
                        info.UseShellExecute = false;
                        info.CreateNoWindow = true;

                        CN = new SqlConnection(connectionString);
                        CN.Open();



                        qry = "";
                        qry = "delete from TempParams";
                        cmd = new SqlCommand(qry, CN);
                        cmd.ExecuteNonQuery();
                        qry = "insert into TempParams select '" + args[0] + "','" + args[1] + "','" + args[2] + "'";
                        cmd = new SqlCommand(qry, CN);
                        cmd.ExecuteNonQuery();
                        CN.Close();

                        //Run the calculation script
                        //Run the calculation script
                        //try
                        //{
                        //    Process process = Process.Start(@"RcodeApp.exe");
                        //    int id = process.Id;
                        //    Process tempProc = Process.GetProcessById(id);
                        //    this.Visible = true;
                        //    tempProc.WaitForExit();
                        //    this.Visible = true;


                        //    process.WaitForExit();
                        //}
                        //catch (Exception ex)
                        //{


                        //    MessageBox.Show(ex.Message);
                        //}
                        try
                        {
                            string INIfolderPath;

                            //Reading data from app.ini file
                            INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                            INIfolderPath = INIfolderPath + "\\app.ini";

                            string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                            connectionString = lines[0];
                            strRscript = lines[1];
                            strRpath = lines[2];
                            var RDirectory = strRpath;
                            strDataDir = lines[3];
                            strOutputDir = lines[4];



                            var pathToFile = @lines[2];
                            string[] arrLine = File.ReadAllLines(pathToFile);


                            using (var conn = new SqlConnection(connectionString))
                            {
                                conn.Open();
                                qry = "";
                                qry = "select distinct * from TempParams";
                                SqlCommand command1 = new SqlCommand(qry, conn);
                                var cmdSelectFromProduct = command1.ExecuteScalar();
                                SqlDataReader dr = command1.ExecuteReader();
                                cmd = new SqlCommand(qry, conn);
                                int i = 0;
                                while (dr.Read())
                                {
                                    strArg[0] = dr[0].ToString();
                                    strArg[1] = dr[1].ToString();
                                    strArg[2] = dr[2].ToString();
                                }
                                dr.Close();
                                conn.Close();
                                try { Console.WriteLine("started"); }
                                catch (Exception ex)
                                { Console.WriteLine(ex.Message); }

                            }
                            //string[] args = { @strArg[0], @strArg[1], @strArg[2] };
                            arrLine[15] = @"filename <-paste0('" + @strArg[0].ToString() + "')";
                            arrLine[16] = @"prodcode <-paste0('" + @strArg[1].ToString() + "')";
                            arrLine[17] = @"OutputDir <-paste0('" + @strArg[2].ToString() + "')";
                            File.WriteAllLines(pathToFile, arrLine);
                            //MessageBox.Show("Calculation is started");
                            //var pBar = new ProgressBar();
                            //pBar.Location = new System.Drawing.Point(50, 200);
                            //pBar.Name = "Calculation";
                            //pBar.Width = 200;
                            //pBar.Height = 30;
                            //Controls.Add(pBar);
                            //pBar.Dock = DockStyle.Top;
                            //pBar.Minimum = 0;
                            //pBar.Maximum = 100;
                            //pBar.Value = 70;

                            REngine.SetEnvironmentVariables(Directory.GetCurrentDirectory() + @"\bin\x64", Directory.GetCurrentDirectory());
                            engine = REngine.GetInstance();

                            // Prepare the fetching of the libraries
                            string RLibraryDirectory = RDirectory + @"\R_LIBS_USER";
                            RLibraryDirectory = RLibraryDirectory.Replace(@"\", @"/");

                            engine.Evaluate("source('" + @pathToFile.Replace("\\", "/") + "')");
                            //pBar.Value = 100;
                            //MessageBox.Show("Calculation is finished, please press any key");
                            txtResult.Text = "Calculation is finished, the number of calculation is " + strCalcID;
                            txProgressBar.Visible = false;
                            btnCalculationView.Enabled = true;
                            btnCalculationSearch.Enabled = true;

                            //pBar.Dispose(); 


                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }




                        try
                        {

                            //ProcessStartInfo p = new ProcessStartInfo();
                            //p.CreateNoWindow = true;
                            //p.UseShellExecute = false;
                            //p.RedirectStandardOutput = true;
                            //p.FileName = "rasdial";
                            //p.Arguments = string.Format("\x22{0}\x22", @"RcodeApp.exe");
                            //p.WindowStyle = ProcessWindowStyle.Hidden;

                            //Process process = Process.Start(p);
                            //process.WaitForExit();
                            strH();
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }


                                //
                                //using (var proc = new Process())
                                //{
                                //proc.StartInfo = info;
                                //proc.Start();
                                //result = proc.StandardOutput.ReadToEnd();
                                //Console.WriteLine(result);
                                //REngine.SetEnvironmentVariables(@"C:\BI_CPV_tool1\bin\x64", Directory.GetCurrentDirectory());
                                //var engine = REngine.GetInstance();
                                ////REngine.SetEnvironmentVariables();
                                ////engine = REngine.GetInstance();
                                //// REngine requires explicit initialization.
                                //// You can set some parameters.
                                //engine.Initialize();
                                //engine.Evaluate(info.Arguments);
                                //var pathToFile = @"C:\BI_CPV_tool1\BI_code1.r";
                                //string[] arrLine = File.ReadAllLines(pathToFile);


                                ////string[] args = { @"C:/OutPut/ExcelFile06-12-202201-52-31.xlsx", @"ASW-76", @"C:/output/plots/06122022005257/" };
                                //arrLine[6] = @"filename <-paste0('" + @args[0] + "')";
                                //arrLine[7] = @"prodcode <-paste0('" + @args[1] + "')";
                                //arrLine[8] = @"OutputDir <-paste0('" + @args[2] + "')";
                                //File.WriteAllLines(pathToFile, arrLine);
                                //engine.Evaluate("source('" + @pathToFile.Replace("\\", "/") + "')");
                                //Console.ReadKey();


                                //strH();

                                //}





                                txtResult.Visible = true;
                                result = result.Replace("# A tibble: 0 Ã—", "");
                        string[] strFiles = { "" };
                        string sName;
                        string prod;
                        string ozid;
                        string dtDateTime;
                        string strPlotName = "";
                        byte[] btImage;
                        CN = new SqlConnection(connectionString);
                        CN.Open();
                            //int fileCount = Directory.GetFiles(@args[2], "*.jpeg", SearchOption.AllDirectories).Length;
                            //if (fileCount > 0)
                            //{
                            //    strFiles = Directory.GetFiles(@args[2], "*.jpeg");
                            //    string s = "";
                            //        string t = "";
                            //        int index = 0;
                            //    for (int j = 0; j < fileCount; j++)
                            //    {

                            //        sName = strFiles[j];
                            //            //Testprod_DP_ANALYTICS_23-01-30 20_10_13.jpeg
                            //        prod = cmbProductCode.Text;

                            //            index = sName.IndexOf(prod);
                            //            sName = sName.Substring(index + prod); 
                            //            s = sName.Substring(strFiles[j].Length - 23);
                            //            t = sName.Substring(prod.Length+1,s.Length+1);
                            //            t   = t.Substring(0,t.Length-16); 
                            //            ozid = t; 
                            int fileCount = Directory.GetFiles(@args[2], "*.jpeg", SearchOption.AllDirectories).Length;
                            if (fileCount > 0)
                            {
                                strFiles = Directory.GetFiles(@args[2], "*.jpeg");
                                int index2 = 0;
                                int index1 = 0;
                                for (int j = 0; j < fileCount; j++)
                                {
                                    prod = cmbProductCode.Text; 
                                    sName = strFiles[j];
                                    
                                    index1 = sName.IndexOf(prod);
                                    
                                        sName = sName.Substring(index1 + prod.Length +1);
                                    //sName = sName.Substring(index1); 
                                    //.Substring(strFiles[j].Length - 39, 39);
                                    sName = sName.Substring(0, sName.Length - 23);

                                    ozid = sName;
                                    dtDateTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
                               
                                btImage = GetPhoto(strFiles[j]);
                                try
                                {

                                    btImage = System.IO.File.ReadAllBytes(strFiles[j]);
                                    System.Data.SqlClient.SqlCommand cmd2 = new System.Data.SqlClient.SqlCommand("insert  into Graphs([GraphName],[VIRT_OZID],[CalcID],[ImageValue], [ID]) values (@GraphName,@VIRT_OZID,@CalcID,@ImageValue, @ID)", CN);
                                    cmd2.Parameters.AddWithValue("@GraphName", sName);
                                    cmd2.Parameters.AddWithValue("@VIRT_OZID", ozid);
                                    cmd2.Parameters.AddWithValue("@CalcID", strCalcID.ToString());
                                    cmd2.Parameters.AddWithValue("@ImageValue", btImage);
                                    cmd2.Parameters.AddWithValue("@ID", strCalcID.ToString() + ozid);
                                    cmd2.ExecuteNonQuery();

                                        
                                    }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    CN.Close();
                                }

                            }

                                timer1.Enabled = true;
                                timer1_Tick(sender, e);
                                var strOutputDirMod = strOutputDir.Replace("/", "\\");
                                System.IO.DirectoryInfo di1 = new DirectoryInfo(strOutputDirMod);
                                foreach (FileInfo file1 in di1.GetFiles())
                                {
                                    file1.Delete();
                                }
                                foreach (DirectoryInfo dir1 in di1.GetDirectories())
                                {
                                    dir1.Delete(true);
                                }
                                CN.Close();

                        }

                        if (fileCount != 0)
                                    txtResult.Text = "Calculation is finished, the number of calculation is " + strCalcID;
                                else
                            txtResult.Text = "Please repeat calculations with other parameters, over insufficient data, there is no files in the output folder " + strOutputDir + strCalcID;//result;
                            txProgressBar.Visible = false;
                        }
                    else
                        MessageBox.Show("Please export the filtered viewgrid data to the excel file first! ");


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            }
                //Thread.CurrentThread.CurrentCulture = lastCulture;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                
            }
        }



        private void btnReset_Click(object sender, EventArgs e)// reset filters
        {

            //txtLastMdataPoints.Text = "";
            //txtLastNdataPoints.Text = "";
            //cmbLaufnrMin.Text = "";
            //cmbLaufnrMax.Text = "";

            //txtProcessCodeName.Text = "";
            //cmbProductCode.Items.Clear();
            //clbRefCpv.Items.Clear();
            //clbVirtOzid.Items.Clear();
            chLastN.Enabled = true;
            chLastN.Checked = false;

            chkAllRefCpv.Checked = true;

            chkAllVirtOzid.Checked = true;

            dtSortDateFrom.Text = strInitialDate;
            dtSortDateTo.Text = strEndDate;

            chkSortDate.Enabled = true;
            chkSortDate.Checked = false;
            chkExclVirtOzid.Checked = false;

            dtSortDateFrom.Enabled = true;
            dtSortDateTo.Enabled = true;
            chkSortDate.Enabled = true;
            cmbLaufnrMin.Enabled = true;
            cmbLaufnrMax.Enabled = true;
            chkLaufNr.Enabled = true;
            txtLastNdataPoints.Enabled = true;
            chLastN.Enabled = true;
            txtLastMdataPoints.Enabled = true;
            chLastM.Enabled = true;
            rbDays.Enabled = true;
            rbWeeks.Enabled = true;
            rbMonths.Enabled = true;
            rbYears.Enabled = true;


            lstCheckExclVirtOzid.Items.Clear();
            listBox1.Items.Clear();
            //clbVirtOzid.Items.Clear();
            //cmbProductCode.Items.Clear();
            //clbRefCpv.Items.Clear();
            EventArgs e1 = null;
            txtLastNdataPoints.Text = "5000";
            txtLastMdataPoints.Text = "3";
            rbYears.Enabled = true;
            //frmBICPV_Load(this, e1);
            chkSortDate.Enabled = true;
            chkSortDate.Checked = false;
            chLastM.Enabled = true;
            chLastM.Checked = false;
            for (int i =0; i< clbVirtOzid.Items.Count; i++)
            {
             
                clbVirtOzid.SetItemChecked(i, false);
            }
            chkAllRefCpv.Checked = true;
            chkAllVirtOzid.Checked = true;
            cmbProductCode.Focus();
            filterOn = false;
            lblDataGridTitle.Text = "Table of raw data entity (Original: Yes  -  Filtered: NO ) ";
            btnCalculationView.Visible = true;
            chkSortDate.Enabled = false;
            clbRefCpv.Enabled = false;
            clbVirtOzid.Enabled = false;
            dataGridViewRaw.Visible = true;
            //using (var conn = new SqlConnection(connectionString))
            //{
            //    conn.Open();





            //    var command2 = new SqlCommand("select * from Products", conn);
            //    command2.CommandType = CommandType.StoredProcedure;

            //    command2.ExecuteNonQuery();


            //    conn.Close();


            //}
            try
            {
                //frmBICPV_Load2(sender, e);
                var con = new SqlConnection(@connectionString);
                var oconn = new SqlCommand("Select * From Products where ProduktCode='" + cmbProductCode.Text + "'", con);

                con.Open();

                SqlDataAdapter sda = new SqlDataAdapter(oconn);
                //System.Data.DataTable data = new System.Data.DataTable();
                //if (dt.Columns.Count > 3)
                //{
                sda.Fill(dt);
                dataGridViewRaw.DataSource = dt;


                CultureInfo culture = new CultureInfo("de-DE");



                con.Close();
            }
            catch(Exception ex) {
                MessageBox.Show(ex.Message);
            }
            }

        public void DataTableToExcel(System.Data.DataTable dt)// making excel file after filtering
        {
            try { 
            string FileName = "Records";
            string SheetName = "Records";
            //there are data from filtered datagrid
            string folderPath = strDataDir;
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, SheetName);
                var strNamePart = DateTime.Now.ToString().Replace(".", "-");

                strExcelFileName = folderPath + "\\" + cmbProductCode.Text + strNamePart.ToString().Replace(":", "-").Replace(" ", "") + ".xlsx";

                wb.SaveAs(strExcelFileName);
                strExcelFileName = strExcelFileName.Replace("\\", "/");
                strExcelFileName = strExcelFileName.Replace("//", "/");



            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private void button1_Click_1(object sender, EventArgs e) //Export Data to Excel For Calculation
        {
            try {
            if (filterIsOn == false)
            {
                MessageBox.Show("Please filter the data first!");
            }
            else
            {
                var strQuery = "SELECT ID,PRODUKTCODE,SORT_DATE,TS_ABS,LAUFNR,CHNR_ENDPRODUKT,PROCESS_CODE,PROCESS_CODE_NAME,PARAMETER_NAME,ASSAY,VIRT_OZID,TREND_WERT,TREND_WERT_2,ISTWERT_LIMS,LCL,UCL,CL,UAL,LAL,DECIMAL_PLACES_XCL_SUBSTITUTED,DECIMAL_PLACES_AL,DATA_TYPE,SOURCE_SYSTEM,EXCURSION,REFERENCED_CPV,IS_IN_RUN_NUMBER_RANGE,LOCATION FROM [dbo].[Products]";
                var strFileName = cmbProductCode.Text + DateTime.Now.ToString().Replace(":", "-") + ".csv";

                string connString = connectionString;
                SqlDataAdapter da;
                SqlCommandBuilder builder;
                System.Data.DataTable table;
                SqlConnection conn;

                table = new System.Data.DataTable("dataGridViewRaw");
                connString = connString.ToString().Replace("//", "/");
                conn = new SqlConnection(connString);



                strSQLfiltered = strSQLfiltered.Replace("and Virt_Ozid in )", "");
                da = new SqlDataAdapter();
                if (@strSQLfiltered == "") @strSQLfiltered = strQuery;
                da.SelectCommand = new SqlCommand(@strSQLfiltered, conn);
                builder = new SqlCommandBuilder(da);
                da.Fill(table);




                DataTableToExcel(table);

            }
            try
            {
                //prepare r libraries
                string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                //C:\Dokumente und Einstellungen\49175\AppData\Local\R\win-library\4.2


                var str = userName;
                var x1 = userName.Replace("\\", "=");
                var parts = x1.Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries)
                               .Select(x => x.Trim());
                foreach (var part in parts)
                    x1 = part;
                var path = @"C:\Dokumente und Einstellungen\" + x1 + @"\AppData\Local\R\win-library\4.2";


                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(x1);
                    DirectoryCopy(@".\rlib", @"C:\Dokumente und Einstellungen\" + x1 + @"\AppData\Local\R\win-library\4.2", true);
                }
                else
                {
                    DirectoryInfo di = new DirectoryInfo(path);
                    var list = Directory.GetFiles(path, "*");
                    var testDirectories = Directory.GetDirectories(path, "*.*");

                    foreach (var directory in testDirectories)
                    {
                        Directory.Delete(directory, true);
                    }
                    //Directory.Delete(path);
                    MessageBox.Show("It is started to copy libraries to " + @"C:\Dokumente und Einstellungen\" + x1 + @"\AppData\Local\R\win-library\4.2" + " folder.");
                    DirectoryCopy(@".\rlib", @"C:\Dokumente und Einstellungen\" + x1 + @"\AppData\Local\R\win-library\4.2", true);
                }
                //DirectoryCopy(@".\rlib", @"C:\Dokumente und Einstellungen\" + userName + @"\AppData\Local\R\win-library\4.2", true);
            }
            catch (Exception e1)
            {

                MessageBox.Show(e1.Message);
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }
        

        private void btnFilterData_Click_1(object sender, EventArgs e) // filter data
        {
            try {
                //frmBICPV_Load2(sender, e);
                var con = new SqlConnection(@connectionString);
            var oconn = new SqlCommand("Select * From Products where ProduktCode='"+cmbProductCode.Text+"'", con);

            con.Open();

            SqlDataAdapter sda = new SqlDataAdapter(oconn);
            //System.Data.DataTable data = new System.Data.DataTable();
            //if (dt.Columns.Count > 3)
            //{
            sda.Fill(dt);
            dataGridViewRaw.DataSource = dt;
            

            CultureInfo culture = new CultureInfo("de-DE");



            con.Close();
            //connection.Close();

            //CopyToSQL(dataGridViewRaw);
            filterIsOn = true;
            intMaxPoints = 200000;
            strFilterExclude = "";
            var strFilterSortDate = "";
           


            int i = 0;

            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string qry = "";
                qry = "select min(CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 )) from Products where laufnr = '" + cmbLaufnrMin.Text + "'";
                SqlCommand command1 = new SqlCommand(qry, conn);
                var cmdSelectFromProduct = command1.ExecuteScalar();


                SqlDataReader dr = command1.ExecuteReader();

                SqlCommand cmd = new SqlCommand(qry, conn);


                while (dr.Read())
                {
                    strLaufNRDateMin = dr[0].ToString();

                }


                dr.Close();

                qry = "select max(CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 )) from Products where laufnr = '" + cmbLaufnrMax.Text + "'";
                command1 = new SqlCommand(qry, conn);
                cmdSelectFromProduct = command1.ExecuteScalar();


                dr = command1.ExecuteReader();

                cmd = new SqlCommand(qry, conn);


                while (dr.Read())
                {
                    strLaufNRDateMax = dr[0].ToString();

                }


                dr.Close();

                conn.Close();
            }

            strFilterSortDate = "";
            // -------------------------------------------------

            var strQuery = " (1 = 1)  and 0=0";


            if (chkAllVirtOzid.Checked != true)
            {
                var strFilterVirtOzid = " and Virt_Ozid in ('";
                //ljkh
                for (int j = 0; j < clbVirtOzid.Items.Count; ++j)
                {
                    if (clbVirtOzid.GetItemCheckState(j) == CheckState.Checked)
                    {
                        strFilterVirtOzid += (string)clbVirtOzid.Items[j] + "','";
                    }
                }
                strFilterVirtOzid = Left(strFilterVirtOzid, strFilterVirtOzid.Length - 2) + ") ";
                int index = strFilterVirtOzid.IndexOf("()");
                if (index <= 0)
                {
                    strQuery += strFilterVirtOzid;

                }
            }


            if (chkAllRefCpv.Checked != true)
            {
                var strFilterRefCpv = " and REFERENCED_CPV in ('"; ;
                for (int j = 0; j < clbRefCpv.Items.Count; ++j)
                {
                    if (clbRefCpv.GetItemCheckState(j) == CheckState.Checked)
                    {
                        strFilterRefCpv += (string)clbRefCpv.Items[j] + "','";
                    }
                }
                strFilterRefCpv = Left(strFilterRefCpv, strFilterRefCpv.Length - 2) + ") ";
                int index = strFilterRefCpv.IndexOf("()");
                if (index <= 0)
                {
                    strQuery += strFilterRefCpv;

                }
                    
            }



                


                //chkLaufNr filter
                strFilterLaufnr = "";
                string dtString1 = "";
                string dtString2 = "";
                using (var conn = new SqlConnection(connectionString))
                {
                    var strQueryTemp = "select distinct min(TS_ABS) from Products where laufnr = '" + cmbLaufnrMin.Text + "'";
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQueryTemp, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                   


                    SqlDataReader dr = command1.ExecuteReader();
                    //cmbRefCpv.Items.Add("All");
                    //cmbRefCpv.Text = "All";
                    //chkAllRefCpv.Checked = true;

                    while (dr.Read())
                    {
                        dtString1 = dr[0].ToString();

                    }

                    dr.Close();
                    conn.Close();
                }
                using (var conn = new SqlConnection(connectionString))
                {
                    var strQueryTemp = "select distinct max(TS_ABS) from Products where laufnr = '" + cmbLaufnrMax.Text + "'";
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQueryTemp, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();




                    SqlDataReader dr = command1.ExecuteReader();
                    //cmbRefCpv.Items.Add("All");
                    //cmbRefCpv.Text = "All";
                    //chkAllRefCpv.Checked = true;

                    while (dr.Read())
                    {
                        dtString2 = dr[0].ToString();

                    }

                    dr.Close();
                    conn.Close();
                }
                culture = System.Globalization.CultureInfo.CreateSpecificCulture("fr-FR");
                //dtString1 = dtString1.Replace(".", "/");
                //DateTime dt1 = DateTime.Parse(dtString1.Split('.')[0], culture);

                //Double milliseconds = Double.Parse(dtString1.Split('.')[1]);

                //dt1 = dt1.AddMilliseconds(milliseconds);



                ////System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.CreateSpecificCulture("fr-FR");
                ////dtString2 = dtString2.Replace(".", "/");
                //DateTime dt2 = DateTime.Parse(dtString2.Split('.')[0], culture);

                // milliseconds = Double.Parse(dtString2.Split('.')[1]);

                //dt2 = dt2.AddMilliseconds(milliseconds);

                //(select distinct min(TS_ABS) from Products where  laufnr = '" + cmbLaufnrMin.Text + "')
                //(select distinct max(TS_ABS) from Products where laufnr = '" + cmbLaufnrMax.Text + "')"

                if (cmbLaufnrMin.Text != "" && cmbLaufnrMax.Text != "" && chkLaufNr.Checked == true)
                    //strFilterLaufnr += " and TS_ABS between (select distinct min(TS_ABS) from Products where  laufnr = '"+ cmbLaufnrMin.Text + "') and (select distinct max(TS_ABS) from Products where laufnr = '" + cmbLaufnrMax.Text + "')";
                    strFilterLaufnr += " and TS_ABS between CONVERT(Datetime,'" + dtString1 + "', 103) and CONVERT(Datetime,'" + dtString2 + "', 103) ";
                //-----------Days
                strFilterSortDate = "";

                var @MaxDate = "";
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var strQueryMaxDate = "select max(TS_ABS) FROM [dbo].[Products]"; 
                    SqlCommand command1 = new SqlCommand(strQueryMaxDate, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                   


                    SqlDataReader dr = command1.ExecuteReader();


                 
                    while (dr.Read())
                    {
                        @MaxDate = (dr[0].ToString());
                      
                    }
                    @MaxDate = @MaxDate.Replace(".", "");
                    var day = Left(@MaxDate, 2);
                    var month = @MaxDate.Substring(2, 2);
                    var year = @MaxDate.Substring(4, 4);
                    @MaxDate = year + month + day;


                    dr.Close();
                    conn.Close();
                }




                if ((txtLastMdataPoints.Text != "") && IsInteger(txtLastMdataPoints.Text) && chLastM.Checked == true)
            {
                if ((rbDays.Checked == true) && (chLastM.Checked))
                {
                    strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(day, -" + txtLastMdataPoints.Text + ", '" + @MaxDate + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";

                    strQuery += strFilterSortDate;

                }
                if ((rbWeeks.Checked == true) && (chLastM.Checked))
                {
                    strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(week, -" + txtLastMdataPoints.Text + ",'" + @MaxDate + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";
                    //strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(week, -" + txtLastMdataPoints.Text + ", GETDATE()) AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";


                        strQuery += strFilterSortDate;

                }
                if ((rbMonths.Checked == true) && (chLastM.Checked))
                {
                    strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(month, -" + txtLastMdataPoints.Text + ", '" + @MaxDate + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";

                    strQuery += strFilterSortDate;

                }
                if ((rbYears.Checked == true) && (chLastM.Checked))
                {
                    strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT(DATETIME, ISNULL(CAST(DATEADD(year, -" + txtLastMdataPoints.Text + ", '" + @MaxDate + "') AS datetime) , '1900-01-01'), 103 ) and CONVERT(DATETIME, ISNULL( getdate() , '1900-01-01'), 103 )  ";

                    strQuery += strFilterSortDate;

                }
            }
            //------------End

            //if (strLaufNRDateMin != "" && strLaufNRDateMax != "" && chkLaufNr.Checked == true)
            //{
            //    strQuery += " and  [TS_ABS] between '" + strLaufNRDateMin + "' and '" + strLaufNRDateMax + "'";

            //}


            //if (strFilterExclude != "" && chkExclVirtOzid.Checked == true)
            //{
            //    strQuery += strFilterExclude;
            //}

            if (dtSortDateFrom.Text != "" && dtSortDateTo.Text != "" && chkSortDate.Checked == true)
            {
                strFilterSortDate += " and CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) between CONVERT( DATETIME, ISNULL(  '" + dtSortDateFrom.Text + "' , '1900-01-01'), 103 ) and CONVERT( DATETIME, ISNULL( '" + dtSortDateTo.Text + "' , '1900-01-01'), 103 )  ";
                strQuery += strFilterSortDate;
            }

            if ((txtLastNdataPoints.Text != "") && IsInteger(txtLastNdataPoints.Text) && chLastN.Checked)
            {
                intMaxPoints = Convert.ToInt32(txtLastNdataPoints.Text);
                strQuery += " and 0=0 " ;
                strQuery = "select top " + intMaxPoints + " * from Products where " + strQuery + " order by  [TS_ABS] desc";
            }
            else
            {
                //intMaxPoints = Convert.ToInt32(txtLastNdataPoints.Text);
                strQuery += strFilterLaufnr + " and 0=0";
                //strQuery = "select * from Products where " + strQuery + " order by CONVERT( DATETIME, ISNULL( [TS_ABS] , '1900-01-01'), 103 ) desc";
                strQuery = "select * from Products where " + strQuery + " and ProduktCode = '" + cmbProductCode.Text + "' order by CONVERT(DATETIME, ISNULL( [TS_ABS], '1900-01-01'), 103 ) desc";
            }
            label1.Text = strQuery ;
            strQuery = strQuery.Replace("and Virt_Ozid in )", "");
            strQuery = strQuery.Replace("and REFERENCED_CPV in )", "");
            strSQLfiltered = strQuery ;

            //for test purpose we show here the SQL query
            txtTestSQL.Text = strQuery ;

            string connection = connectionString;
            using (var conn = new SqlConnection(connectionString))
            {
                try{ 
                conn.Open();
                SqlCommand command = new SqlCommand(strQuery, conn);
                command.CommandTimeout = 600;
                var cmdSelectFromProduct = command.ExecuteScalar();

                System.Data.DataTable table = new System.Data.DataTable();

                //
                if (strHasID == "0")
                {
                    table.Columns.Add("PRODUKTCODE", typeof(string));
                    table.Columns.Add("SORT_DATE", typeof(DateTime));
                    table.Columns.Add("TS_ABS", typeof(DateTime));
                    table.Columns.Add("LAUFNR", typeof(string));
                    table.Columns.Add("CHNR_ENDPRODUKT", typeof(string));
                    table.Columns.Add("PROCESS_CODE", typeof(string));
                    table.Columns.Add("PROCESS_CODE_NAME", typeof(string));
                    table.Columns.Add("PARAMETER_NAME", typeof(string));
                    table.Columns.Add("ASSAY", typeof(string));
                    table.Columns.Add("VIRT_OZID", typeof(string));
                    table.Columns.Add("TREND_WERT", typeof(string));
                    table.Columns.Add("TREND_WERT_2", typeof(string));
                    table.Columns.Add("ISTWERT_LIMS", typeof(string));
                    table.Columns.Add("LCL", typeof(string));
                    table.Columns.Add("UCL", typeof(string));
                    table.Columns.Add("CL", typeof(string));
                    table.Columns.Add("UAL", typeof(string));
                    table.Columns.Add("LAL", typeof(string));
                    table.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", typeof(string));
                    table.Columns.Add("DECIMAL_PLACES_AL", typeof(string));
                    table.Columns.Add("DATA_TYPE", typeof(string));
                    table.Columns.Add("SOURCE_SYSTEM", typeof(string));
                    table.Columns.Add("EXCURSION", typeof(string));
                    table.Columns.Add("REFERENCED_CPV", typeof(string));
                    table.Columns.Add("IS_IN_RUN_NUMBER_RANGE", typeof(string));
                    table.Columns.Add("LOCATION", typeof(string));
                }
                if (strHasID == "1")
                {
                    table.Columns.Add("ID", typeof(string));
                    table.Columns.Add("PRODUKTCODE", typeof(string));
                    table.Columns.Add("SORT_DATE", typeof(DateTime));
                    table.Columns.Add("TS_ABS", typeof(DateTime));
                    table.Columns.Add("LAUFNR", typeof(string));
                    table.Columns.Add("CHNR_ENDPRODUKT", typeof(string));
                    table.Columns.Add("PROCESS_CODE", typeof(string));
                    table.Columns.Add("PROCESS_CODE_NAME", typeof(string));
                    table.Columns.Add("PARAMETER_NAME", typeof(string));
                    table.Columns.Add("ASSAY", typeof(string));
                    table.Columns.Add("VIRT_OZID", typeof(string));
                    table.Columns.Add("TREND_WERT", typeof(string));
                    table.Columns.Add("TREND_WERT_2", typeof(string));
                    table.Columns.Add("ISTWERT_LIMS", typeof(string));
                    table.Columns.Add("LCL", typeof(string));
                    table.Columns.Add("UCL", typeof(string));
                    table.Columns.Add("CL", typeof(string));
                    table.Columns.Add("UAL", typeof(string));
                    table.Columns.Add("LAL", typeof(string));
                    table.Columns.Add("DECIMAL_PLACES_XCL_SUBSTITUTED", typeof(string));
                    table.Columns.Add("DECIMAL_PLACES_AL", typeof(string));
                    table.Columns.Add("DATA_TYPE", typeof(string));
                    table.Columns.Add("SOURCE_SYSTEM", typeof(string));
                    table.Columns.Add("EXCURSION", typeof(string));
                    table.Columns.Add("REFERENCED_CPV", typeof(string));
                    table.Columns.Add("IS_IN_RUN_NUMBER_RANGE", typeof(string));
                    table.Columns.Add("LOCATION", typeof(string));
                }


                SqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {
                    if (strHasID == "1")
                        table.Rows.Add(dr["ID"].ToString(), dr["PRODUKTCODE"].ToString(), dr["SORT_DATE"], dr["TS_ABS"], dr["LAUFNR"].ToString(), dr["CHNR_ENDPRODUKT"].ToString(), dr["PROCESS_CODE"].ToString(), dr["PROCESS_CODE_NAME"].ToString(), dr["PARAMETER_NAME"].ToString(), dr["ASSAY"].ToString(), dr["VIRT_OZID"].ToString(), dr["TREND_WERT"].ToString(), dr["TREND_WERT_2"].ToString(), dr["ISTWERT_LIMS"].ToString(), dr["LCL"].ToString(), dr["UCL"].ToString(), dr["CL"].ToString(), dr["UAL"].ToString(), dr["LAL"].ToString(), dr["DECIMAL_PLACES_XCL_SUBSTITUTED"].ToString(), dr["DECIMAL_PLACES_AL"].ToString(), dr["DATA_TYPE"].ToString(), dr["SOURCE_SYSTEM"].ToString(), dr["EXCURSION"].ToString(), dr["REFERENCED_CPV"].ToString(), dr["IS_IN_RUN_NUMBER_RANGE"].ToString(), dr["LOCATION"].ToString());
                    if (strHasID == "0")
                        table.Rows.Add(dr["PRODUKTCODE"].ToString(), dr["SORT_DATE"], dr["TS_ABS"], dr["LAUFNR"].ToString(), dr["CHNR_ENDPRODUKT"].ToString(), dr["PROCESS_CODE"].ToString(), dr["PROCESS_CODE_NAME"].ToString(), dr["PARAMETER_NAME"].ToString(), dr["ASSAY"].ToString(), dr["VIRT_OZID"].ToString(), dr["TREND_WERT"].ToString(), dr["TREND_WERT_2"].ToString(), dr["ISTWERT_LIMS"].ToString(), dr["LCL"].ToString(), dr["UCL"].ToString(), dr["CL"].ToString(), dr["UAL"].ToString(), dr["LAL"].ToString(), dr["DECIMAL_PLACES_XCL_SUBSTITUTED"].ToString(), dr["DECIMAL_PLACES_AL"].ToString(), dr["DATA_TYPE"].ToString(), dr["SOURCE_SYSTEM"].ToString(), dr["EXCURSION"].ToString(), dr["REFERENCED_CPV"].ToString(), dr["IS_IN_RUN_NUMBER_RANGE"].ToString(), dr["LOCATION"].ToString());

                }
                        dataGridViewRaw.Columns["TS_ABS"].DefaultCellStyle.Format = "dd.MM.yyyy";
                        dataGridViewRaw.Columns["SORT_DATE"].DefaultCellStyle.Format = "dd.MM.yyyy";
                dataGridViewRaw.DataSource = table;
                        
                        dr.Close();
                conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        

        private void chkSortDate_CheckedChanged(object sender, EventArgs e)//Enable/Disable SortDate filter
        {
            if (chkSortDate.Checked)
            {
                chLastN.Checked = false;
                chLastN.Enabled = false;
                chLastM.Checked = false;
                chLastM.Enabled = false;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = false;
                rbDays.Enabled = false;
                rbWeeks.Enabled = false;
                rbMonths.Enabled = false;
                rbYears.Enabled = false;
                dtSortDateFrom.Enabled = true;
                dtSortDateTo.Enabled = true;
                cmbLaufnrMin.Enabled = false;
                cmbLaufnrMax.Enabled = false;
                txtLastNdataPoints.Enabled = false;
                txtLastMdataPoints.Enabled = false;

            }
            else
            {
                chLastN.Checked = false;
                chLastN.Enabled = true;
                chLastM.Checked = false;
                chLastM.Enabled = true;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = true;
                rbDays.Enabled = true;
                rbWeeks.Enabled = true;
                rbMonths.Enabled = true;
                rbYears.Enabled = true;
                dtSortDateFrom.Enabled = true;
                dtSortDateTo.Enabled = true;
                cmbLaufnrMin.Enabled = true;
                cmbLaufnrMax.Enabled = true;
                txtLastNdataPoints.Enabled = true;
                txtLastMdataPoints.Enabled = true;
            }

        }

        private void chkLaufNr_CheckedChanged(object sender, EventArgs e)//Enable/Disable LaufNr filter
        {
            if (chkLaufNr.Checked)
            {
                chLastN.Checked = false;
                chLastN.Enabled = false;
                chLastM.Checked = false;
                chLastM.Enabled = false;
                chkSortDate.Checked = false;
                chkSortDate.Enabled = false;
                rbDays.Enabled = false;
                rbWeeks.Enabled = false;
                rbMonths.Enabled = false;
                rbYears.Enabled = false;
                dtSortDateFrom.Enabled = false;
                dtSortDateTo.Enabled = false;

                cmbLaufnrMin.Enabled = true;
                cmbLaufnrMax.Enabled = true;
                txtLastNdataPoints.Enabled = false;
                txtLastMdataPoints.Enabled = false;

            }
            else
            {
                chLastN.Checked = false;
                chLastN.Enabled = true;
                chLastM.Checked = false;
                chLastM.Enabled = true;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = true;
                rbDays.Enabled = true;
                rbWeeks.Enabled = true;
                rbMonths.Enabled = true;
                rbYears.Enabled = true;
                chkSortDate.Checked = false;
                chkSortDate.Enabled = true;
                dtSortDateFrom.Enabled = true;
                dtSortDateTo.Enabled = true;
                cmbLaufnrMin.Enabled = true;
                cmbLaufnrMax.Enabled = true;
                txtLastNdataPoints.Enabled = true;
                txtLastMdataPoints.Enabled = true;
            }
        }

        private void chLastN_CheckedChanged(object sender, EventArgs e)//Enable/Disable LastN filter
        {
            if (chLastN.Checked)
            {
                chkSortDate.Checked = false;
                chkSortDate.Enabled = false;
                chLastM.Checked = false;
                chLastM.Enabled = false;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = false;
                rbDays.Enabled = false;
                rbWeeks.Enabled = false;
                rbMonths.Enabled = false;
                rbYears.Enabled = false;
                dtSortDateFrom.Enabled = false;
                dtSortDateTo.Enabled = false;
                //txtLastNdataPoints.Enabled = false;
                txtLastMdataPoints.Enabled = false;

                dtSortDateFrom.Enabled = false;
                dtSortDateTo.Enabled = false;
                cmbLaufnrMin.Enabled = false;
                cmbLaufnrMax.Enabled = false;
                txtLastNdataPoints.Enabled = true;
                txtLastMdataPoints.Enabled = false;

            }
            else
            {
                chkSortDate.Checked = false;
                chkSortDate.Enabled = true;
                chLastM.Checked = false;
                chLastM.Enabled = true;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = true;
                rbDays.Enabled = true;
                rbWeeks.Enabled = true;
                rbMonths.Enabled = true;
                rbYears.Enabled = true;
                dtSortDateFrom.Enabled = true;
                dtSortDateTo.Enabled = true;
                cmbLaufnrMin.Enabled = true;
                cmbLaufnrMax.Enabled = true;
                txtLastNdataPoints.Enabled = true;
                txtLastMdataPoints.Enabled = true;
            }
        }

        private void chLastM_CheckedChanged(object sender, EventArgs e)//Enable/Disable LastM filter
        {
            if (chLastM.Checked)
            {
                chkSortDate.Checked = false;
                chkSortDate.Enabled = false;
                chLastN.Checked = false;
                chLastN.Enabled = false;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = false;
                rbDays.Enabled = true;
                rbWeeks.Enabled = true;
                rbMonths.Enabled = true;
                rbYears.Enabled = true;
                dtSortDateFrom.Enabled = false;
                dtSortDateTo.Enabled = false;
                cmbLaufnrMin.Enabled = false;
                cmbLaufnrMax.Enabled = false;
                txtLastNdataPoints.Enabled = false;
                txtLastMdataPoints.Enabled = true;

            }
            else
            {
                chkSortDate.Checked = false;
                chkSortDate.Enabled = true;
                chLastN.Checked = false;
                chLastN.Enabled = true;
                chkLaufNr.Checked = false;
                chkLaufNr.Enabled = true;
                rbDays.Enabled = false;
                rbWeeks.Enabled = false;
                rbMonths.Enabled = false;
                rbYears.Enabled = false;
                dtSortDateFrom.Enabled = true;
                dtSortDateTo.Enabled = true;
                cmbLaufnrMin.Enabled = true;
                cmbLaufnrMax.Enabled = true;
                txtLastNdataPoints.Enabled = true;
                txtLastMdataPoints.Enabled = true;
            }
        }

        private void chkAllRefCpv_CheckedChanged(object sender, EventArgs e)//Enable/Disable AllRefCpv filter
        {
            try { 
            string strQuery2 = "select distinct REFERENCED_CPV from Products order by REFERENCED_CPV";
            if (chkAllRefCpv.Checked)
            {
                clbRefCpv.Enabled = false;
                clbRefCpv.Items.Clear();
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery2, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("REFERENCED_CPV", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();
                    //cmbRefCpv.Items.Add("All");
                    //cmbRefCpv.Text = "All";
                    chkAllRefCpv.Checked = true;

                    while (dr.Read())
                    {
                        clbRefCpv.Items.Add(dr["REFERENCED_CPV"].ToString());
                        
                    }

                    dr.Close();
                    conn.Close();
                }

            }
            else
            {
                clbRefCpv.Enabled = true;

            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private void chkAllVirtOzid_CheckedChanged(object sender, EventArgs e)//Enable/Disable AllVirtOzid  filter
        {
            try { 
            string strQuery = "select distinct VIRT_OZID from Products order by VIRT_OZID";
            if (chkAllVirtOzid.Checked)
            {
                clbVirtOzid.Enabled = false;
                clbVirtOzid.Items.Clear();
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("VIRT_OZID", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();



                    while (dr.Read())
                    {
                        clbVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                        lstCheckExclVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                    }

                    dr.Close();
                    conn.Close();
                }
            }
            else
            {
                clbVirtOzid.Enabled = true;
                
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private void btnFilterData_Click(object sender, EventArgs e) // filter data
        {   if ((dtSortDateFrom.Value > dtSortDateTo.Value) && (chkSortDate.Checked == true))
            {
                MessageBox.Show("The first Date should be less or equal to the second Date!");

            }
            else 
            { 
                try {
                     filterOn = true;
                    
                    if (clbRefCpv.Items.Count == 0)
                    {
                        chkAllRefCpv.Checked = true;
                    }
                    if (clbVirtOzid.Items.Count == 0)
                    {
                        chkAllVirtOzid.Checked = true;
                    }
                    lblDataGridTitle.Text = "Table of raw data entity (Original: NO  -  Filtered: Yes ) ";
                    btnFilterData_Click_1(sender, e);
                    //clbRefCpv.Items.Clear();
                    
                    //for (int i = 0; i < lstCheckExclVirtOzid.Items.Count;++i)
                    //{
                    //    clbRefCpv.Items.Add(lstCheckExclVirtOzid.Items[i]);
                        
                        
                    //    if (chkAllRefCpv.Checked == false)
                    //    {
                    //        clbRefCpv.SetItemChecked(i, true);
                    //    }
                    //}
                    
                    MessageBox.Show("Action : <Filter Data> is finished!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
            LastWindowState = FormWindowState.Minimized;
        }
        private void Form1_Resize(object sender, EventArgs e) //form resize
        {

            // When window state changes
            if (WindowState != LastWindowState)
            {
                LastWindowState = WindowState;


                if (WindowState == FormWindowState.Maximized)
                {
                    MessageBox.Show("Maximized");
                    // Maximized!
                }
                if (WindowState == FormWindowState.Normal)
                {
                    MessageBox.Show("Restored");
                    // Restored!
                }
            }

        }
        
        private void Form1_Load(object sender, EventArgs e) // initial form load
        {
            

            CultureInfo gb = CultureInfo.CreateSpecificCulture("en-GB");

            ToolTip toolTip1 = new ToolTip();

                // Set up the delays for the ToolTip.
                toolTip1.AutoPopDelay = 5000;
                toolTip1.InitialDelay = 1000;
                toolTip1.ReshowDelay = 500;
                // Force the ToolTip text to be displayed whether or not the form is active.
                toolTip1.ShowAlways = true;

                // Set up the ToolTip text for the Button and Checkbox.

                toolTip1.SetToolTip(chkAllRefCpv, "Enable/Disable RefCpv items, if it is enabled then we use all of them");

            ToolTip toolTip2 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip2.AutoPopDelay = 5000;
            toolTip2.InitialDelay = 1000;
            toolTip2.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip2.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip2.SetToolTip(chkAllVirtOzid, "Enable/Disable Virt_Ozid items, if it is enabled then we use all of them");

            ToolTip toolTip3 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip3.AutoPopDelay = 5000;
            toolTip3.InitialDelay = 1000;
            toolTip3.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip3.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip3.SetToolTip(chkSortDate, "Enable/Disable filter Sort Date");


            ToolTip toolTip4 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip4.AutoPopDelay = 5000;
            toolTip4.InitialDelay = 1000;
            toolTip4.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip4.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip4.SetToolTip(chLastN, "Enable/Disable filter Last N points");

            ToolTip toolTip5 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip5.AutoPopDelay = 5000;
            toolTip5.InitialDelay = 1000;
            toolTip5.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip5.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip5.SetToolTip(chLastM, "Enable/Disable filter Last M points");

            ToolTip toolTip6 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip6.AutoPopDelay = 5000;
            toolTip6.InitialDelay = 1000;
            toolTip6.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip6.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip6.SetToolTip(chkLaufNr, "Enable/Disable filter LaufNr");

            


            

                string INIfolderPath;

            //Reading data from app.ini file
            INIfolderPath = System.IO.Directory.GetCurrentDirectory();
            INIfolderPath = INIfolderPath + "\\app.ini";

            string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
            connectionString = lines[0];
            strRscript = lines[1];
            strRpath = lines[2];
            strDataDir = lines[3];
            strOutputDir = lines[4];
            
            delProdukts();
           
            //frmBICPV_Load3(sender, e);
            using (var conn1 = new SqlConnection(@connectionString))
            {
                conn1.Open();
                SqlCommand command1 = new SqlCommand("delete from DataGrid ", conn1);

                var cmdSelectFromProduct = command1.ExecuteScalar();
               
                conn1.Close();

            }
            btnReset.Enabled = false;

        }

        public void delProdukts() // clearing the table Products
        {


            using (var conn1 = new SqlConnection(@connectionString))
            {
                SqlConnection con = new SqlConnection(@connectionString);
                SqlCommand oconn = new SqlCommand("delete From Products", con);
                con.Open();
                SqlCommand command = new SqlCommand("delete From Products", con);
                command.CommandTimeout = 600;
                var r = command.ExecuteNonQuery();
                conn1.Close();



            }

        }
        private void tabMain_Click(object sender, EventArgs e) // Clicking the main tab
        {
            try
            {

                if (tabMain.SelectedTab.Text == "Calculation View")
                {

                //hiding fields of previous tab "Main"
                lblProductCode.Visible = false;
                cmbProductCode.Visible = false;
                lblRferencedCPV.Visible = false;
                clbRefCpv.Visible = false;
                lblVirtOzid.Visible = false;
                clbVirtOzid.Visible = false;
                chkAllVirtOzid.Visible = false;
                btnFilterData.Visible = false;
                btnReset.Visible = false;
                lblSortDate.Visible = false;
                lblLaufNRfrom.Visible = false;
                lblLastNdataPoints.Visible = false;
                lblLastMdataPoints.Visible = false;
                dtSortDateFrom.Visible = false;
                dtSortDateTo.Visible = false;
                chkSortDate.Visible = false;
                label2.Visible = false;
                cmbLaufnrMin.Visible = false;
                cmbLaufnrMax.Visible = false;
                chkLaufNr.Visible = false;
                lblActiveLN.Visible = false;
                txtLastNdataPoints.Visible = false;
                chLastN.Visible = false;
                //dataGridViewRaw.Visible = false;

                rbDays.Visible = false;
                rbWeeks.Visible = false;
                rbMonths.Visible = false;
                rbYears.Visible = false;
                btnOpenFile.Visible = false;
                button1.Visible = false;
                btnCalculationSearch.Visible = false;
                btnCalculationView.Visible = false;
                btnCalculation.Visible = false;
                //-----------------------------
                //hiding fields of previous tab "Calculation Search"
                panel3.Visible = false;
                groupFilterSelection.Visible = false;
                panelSelection.Visible = false;
                panelButtons.Visible = false;
                //-----------------------------

                tabMain.SelectTab("ViewCalculation");
                frmSave_Load(sender, e);
            }
            
            if (tabMain.SelectedTab.Text == "Compare Calculations")              
            {
              
                tabMain.SelectTab("CompareCalculation");
                string strQuery = "select distinct VIRT_OZID from CalculationRaw order by VIRT_OZID";
                string strQuery2 = "select distinct CalcID from CalculationRaw order by CalcID";
                string strQuery3 = "select distinct PRODUCTCODE from CalculationRaw order by PRODUCTCODE";
                chkVirtOzid3.Items.Clear();
                chkVirtOzid4.Items.Clear();
                cmbProdID3.Items.Clear();
                cmbProdID4.Items.Clear();
                cmbCalcID3.Items.Clear();
                cmbCalcID4.Items.Clear();
                SQLRunFillChechedListBox(strQuery, chkVirtOzid3);
                SQLRunFillChechedListBox(strQuery, chkVirtOzid4);
                SQLRunFill(strQuery3, cmbProdID3);
                SQLRunFill(strQuery3, cmbProdID4);
                SQLRunFill(strQuery2, cmbCalcID3);
                SQLRunFill(strQuery2, cmbCalcID4);
               
            }

            if (tabMain.SelectedTab.Text == "Calculation Search")
            {
                panel3.Visible = true;
                groupFilterSelection.Visible = true;
                panelSelection.Visible = true;
                panelButtons.Visible = true;
                tabMain.SelectTab("SearchCalculation");
                frmHistResults_Load(sender, e);
            }



            if (tabMain.SelectedTab.Text == "Main")

            {
                lblProductCode.Visible = true;
                cmbProductCode.Visible = true;
                lblRferencedCPV.Visible = true;
                clbRefCpv.Visible = true;
                lblVirtOzid.Visible = true;
                clbVirtOzid.Visible = true;
                chkAllVirtOzid.Visible = true;
                btnFilterData.Visible = true;
                btnReset.Visible = true;
                lblSortDate.Visible = true;
                lblLaufNRfrom.Visible = true;
                lblLastNdataPoints.Visible = true;
                lblLastMdataPoints.Visible = true;
                dtSortDateFrom.Visible = true;
                dtSortDateTo.Visible = true;
                chkSortDate.Visible = true;
                label2.Visible = true;
                cmbLaufnrMin.Visible = true;
                cmbLaufnrMax.Visible = true;
                chkLaufNr.Visible = true;
                lblActiveLN.Visible = true;
                txtLastNdataPoints.Visible = true;
                chLastN.Visible = true;
                //dataGridViewRaw.Visible = false;

                rbDays.Visible = true;
                rbWeeks.Visible = true;
                rbMonths.Visible = true;
                rbYears.Visible = true;
                btnOpenFile.Visible = true;
                button1.Visible = true;
                btnCalculationSearch.Visible = true;
                btnCalculationView.Visible = true;
                btnCalculation.Visible = true;
                tabMain.SelectTab("Main");
                    if (dataGridViewRaw.Rows.Count < 2)
                        Form1_Load(sender, e);
                    //btnOpenFile_Click(sender, e);
                    else
                        ;
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
        private void btnReset_Click_1(object sender, EventArgs e) // Reset the filters
        {
            try
            {
                rbYears.Checked = true;

                lblDataGridTitle.Text = "Table of raw data entity(Original: Yes - Filtered: NO)";

                btnReset_Click(sender, e);
                chkSortDate.Enabled = true;
                chkLaufNr.Enabled = true;
                chLastN.Enabled = true;
                chLastM.Enabled = true;
                chkSortDate.Checked = false;
                chkLaufNr.Checked = false;
                chLastN.Checked = false;
                chLastM.Checked = false;
                

                MessageBox.Show("Action : <Reset Filter> is finished!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            
        }

        private void button1_Click(object sender, EventArgs e) //Export Data to Excel For Calculation
        {
            try
            {
                button1_Click_1(sender,e);
                MessageBox.Show("Action : <Export Data to Excel For Calculation> is finished!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            
        }

        private void btnCalculation_Click_1(object sender, EventArgs e) // Start calculation process
        {
            try
            {
                stopCalc = true;
            btnCalculation_Click(sender, e);
                MessageBox.Show("The calculation is finished!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private void btnOpenFile_Click(object sender, EventArgs e) // upload input excel file
        {
            filterOn = false;
            
            lblDataGridTitle.Text = "Table of raw data entity (Original: Yes  -  Filtered: NO ) ";
            btnCalculationView.Visible = true;
            chkSortDate.Enabled = false;
            clbRefCpv.Enabled = false;
            clbVirtOzid.Enabled = false;
            dataGridViewRaw.Visible = true;
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();

                



                var command2 = new SqlCommand("ProdUpdate", conn);
                command2.CommandType = CommandType.StoredProcedure;

                command2.ExecuteNonQuery();


                conn.Close();


            }

            //for (int i = 0; i <= arrDataGridView.Count(); i++)
            //{
            //    arrDataGridView[i].DataSource = dt;
            //    arrDataGridView[i].Visible = false;
            //}
            //dataGridViewRaw.Visible = false;

            arrDataGridView[intCounterDataGridViews] = new DataGridView();

            arrDataGridView[intCounterDataGridViews].Left = dataGridViewRaw.Left;
            arrDataGridView[intCounterDataGridViews].Top = dataGridViewRaw.Top;
            arrDataGridView[intCounterDataGridViews].Visible = true;
            arrDataGridView[intCounterDataGridViews].Enabled = true;
            //arrDataGridView[intCounterDataGridViews].Columns.Add("zzz","777");
            //arrDataGridView[intCounterDataGridViews].Rows.Add("zzz");




            try
            {





                string INIfolderPath;

                //Reading data from app.ini file
                INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                INIfolderPath = INIfolderPath + "\\app.ini";

                string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                connectionString = lines[0];
                strRscript = lines[1];
                strRpath = lines[2];
                strDataDir = lines[3];
                strOutputDir = lines[4];

                SqlConnection con = new SqlConnection(@connectionString);
                SqlCommand oconn = new SqlCommand("delete From Products", con);
                con.Open();
                SqlCommand command = new SqlCommand("delete From Products", con);
                command.CommandTimeout = 600;
                oconn.CommandTimeout = 600;

                //var r = command.ExecuteNonQuery();


                //this.MaximizeBox = false;
                rbYears.Checked = true;
                dtSortDateFrom.Text = "01.01.2000";
                chkExclVirtOzid.Checked = true;
                chkSortDate.Enabled = true;
                chkSortDate.Checked = false;




                clbRefCpv.Items.Clear();
                clbVirtOzid.Items.Clear();


                using (StreamWriter writer = new StreamWriter(INIfolderPath, true))
                {
                    openFileDialog1.Filter = "Excel Files | *.xlsx";

                    //Select input excel data file
                    DialogResult result = openFileDialog1.ShowDialog();
                    strFileName = openFileDialog1.FileName;
                    if (openFileDialog1.FileName.ToString() != "openFileDialog1")
                    {
                        clbRefCpv.Items.Clear();
                        clbVirtOzid.Items.Clear();
                        btnOpenFile.BackColor = Color.LightBlue;
                        
                        DataTable dtRanges = new DataTable();
                        //dtTable = new DataTable();
                        //dt = (DataTable)dataGridViewRaw.DataSource;
                        IExcelReader excelReader = new ExcelReader();
                        IDataProcessor dataProcessor = new DataProcessor();
                        ExcelService excelService = new ExcelService(excelReader, dataProcessor);

                        string filePath = @"C:\path\to\your\excel.xlsx";
                        dt = excelService.ReadAndProcessExcel(openFileDialog1.FileName.ToString()); 
                        con = new SqlConnection(@connectionString);
                        oconn = new SqlCommand("Select * From Products", con);
                        oconn.CommandTimeout = 180;
                        con.Open();

                        SqlDataAdapter sda = new SqlDataAdapter(oconn);
                        //System.Data.DataTable data = new System.Data.DataTable();
                        //if (dt.Columns.Count > 3)
                        //{
                        sda.Fill(dt);
                        if (dt.Columns[0].ColumnName == "PRODUKTCODE") { strHasID = "0"; }
                        else if (dt.Columns[0].ColumnName == "ID" || dt.Columns[0].ColumnName == "Column1") { strHasID = "1"; }
                        else { strHasID = "2"; }
                        //dataGridViewRaw.Columns[1].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm:ss";
                        //dataGridViewRaw.Columns[2].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm:ss";



                        
                      
                        dataGridViewRaw.DataSource = dt;
                        //dataGridViewRaw.DataSource = dt;
                        //arrDataGridView[intCounterDataGridViews].Visible = true;
                        //}

                        CultureInfo culture = new CultureInfo("de-DE");
                        dataGridViewRaw.Columns["TS_ABS"].DefaultCellStyle.Format = "dd.MM.yyyy";
                        dataGridViewRaw.Columns["SORT_DATE"].DefaultCellStyle.Format = "dd.MM.yyyy";


                        con.Close();
                        //connection.Close();
                        if (strHasID == "0")
                            CopyToSQL2(dataGridViewRaw);
                        if (strHasID == "1")
                            CopyToSQL(dataGridViewRaw);
                    }
                    else
                    {
                        btnOpenFile.Focus();
                    }
                }

                if (strHasID == "1")
                {
                    string[] strFields = { "ID", "PRODUKTCODE", "SORT_DATE", "TS_ABS", "LAUFNR", "CHNR_ENDPRODUKT", "PROCESS_CODE", "PROCESS_CODE_NAME", "PARAMETER_NAME", "ASSAY", "VIRT_OZID", "TREND_WERT", "TREND_WERT_2", "ISTWERT_LIMS", "LCL", "UCL", "CL", "UAL", "LAL", "DECIMAL_PLACES_XCL_SUBSTITUTED", "DECIMAL_PLACES_AL", "DATA_TYPE", "SOURCE_SYSTEM", "EXCURSION", "REFERENCED_CPV", "IS_IN_RUN_NUMBER_RANGE", "LOCATION", "UserID", "ModifiedDate", "GraphID", "CalcID", "FilterID" };
                }
                if (strHasID == "0")
                {
                    string[] strFields = { "PRODUKTCODE", "SORT_DATE", "TS_ABS", "LAUFNR", "CHNR_ENDPRODUKT", "PROCESS_CODE", "PROCESS_CODE_NAME", "PARAMETER_NAME", "ASSAY", "VIRT_OZID", "TREND_WERT", "TREND_WERT_2", "ISTWERT_LIMS", "LCL", "UCL", "CL", "UAL", "LAL", "DECIMAL_PLACES_XCL_SUBSTITUTED", "DECIMAL_PLACES_AL", "DATA_TYPE", "SOURCE_SYSTEM", "EXCURSION", "REFERENCED_CPV", "IS_IN_RUN_NUMBER_RANGE", "LOCATION", "UserID", "ModifiedDate", "GraphID", "CalcID", "FilterID" };
                }
                var strTableName = "Products";

                //Add data to combo boxes:
                string connectionString1 = connectionString;

                //Setup of comboboxes for the fields VIRT_OZID ,REFERENCED_CPV , ProductCode
                string strQuery = "select distinct VIRT_OZID from Products order by VIRT_OZID";
                string strQuery2 = "select distinct REFERENCED_CPV from Products order by REFERENCED_CPV";
                string strQuery3 = "";
                if (cmbProductCode.Text == "")
                {
                  strQuery3 = "select distinct PRODUKTCODE from Products ";

                }
                else
                {
                    strQuery3 = "select distinct PRODUKTCODE from Products where PRODUKTCODE!='" + cmbProductCode.Text + "'";
                }
                
                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("VIRT_OZID", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();



                    while (dr.Read())
                    {
                        clbVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                        lstCheckExclVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                    }

                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery2, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("REFERENCED_CPV", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();
                    //cmbRefCpv.Items.Add("All");
                    //cmbRefCpv.Text = "All";
                    chkAllRefCpv.Checked = true;

                    while (dr.Read())
                    {
                        clbRefCpv.Items.Add(dr["REFERENCED_CPV"].ToString());

                    }

                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery3, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("PRODUKTCODE", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();


                    var prCode = "";
                    while (dr.Read())
                    {
                        cmbProductCode.Items.Add(dr["PRODUKTCODE"].ToString());
                        prCode =  dr["PRODUKTCODE"].ToString();
                    }
                    cmbProductCode.Text = prCode;
                    //cmbProductCode.Items.Add("All");
                    //cmbProductCode.Text = "All";
                    dr.Close();
                    conn.Close();

                }
                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand("SELECT  distinct  top 1000 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select distinct top 5000 TS_ABS FROM [dbo].[Products] order by TS_ABS )", conn);
                    command1.CommandTimeout = 600;





                    SqlDataReader dr = command1.ExecuteReader();



                    while (dr.Read())
                    {
                        cmbLaufnrMin.Items.Add(dr["LAUFNR"].ToString());

                    }

                    command1 = new SqlCommand("SELECT  distinct  top 1 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select min(TS_ABS) FROM [dbo].[Products] ) ", conn);
                    dr.Close();
                    command1.CommandTimeout = 600;
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        cmbLaufnrMin.Text = dr["LAUFNR"].ToString();
                    }
                    dr.Close();
                    command1.CommandTimeout = 600;
                    command1 = new SqlCommand("SELECT  distinct  top 1000 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select distinct top 5000 TS_ABS FROM [dbo].[Products] order by TS_ABS desc)", conn);
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        cmbLaufnrMax.Items.Add(dr["LAUFNR"].ToString());
                    }
                    command1 = new SqlCommand("SELECT  distinct  top 1 LAUFNR  FROM [dbo].[Products] where TS_ABS in (select max(TS_ABS) FROM [dbo].[Products] )", conn);
                    dr.Close();
                    command1.CommandTimeout = 600;
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        cmbLaufnrMax.Text = dr["LAUFNR"].ToString();
                    }


                    //cmbProductCode.Items.Add("All");
                    //cmbProductCode.Text = "All";
                    dr.Close();
                    conn.Close();

                }

                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand("SELECT  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS", conn);
                    command1.CommandTimeout = 600;





                    SqlDataReader dr = command1.ExecuteReader();



                    //while (dr.Read())
                    //{
                    //    dtSortDateFrom.Items.Add(dr["TS_ABS"].ToString());

                    //}

                    command1 = new SqlCommand("SELECT top 1  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS", conn);
                    dr.Close();
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        dtSortDateFrom.Text = dr["TS_ABS"].ToString();
                        strInitialDate = dr["TS_ABS"].ToString();
                    }
                    dr.Close();
                    command1 = new SqlCommand("SELECT  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS desc", conn);
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        dtSortDateTo.Text = dr["TS_ABS"].ToString();
                    }
                    command1 = new SqlCommand("SELECT top 1  [TS_ABS] FROM [dbo].[Products] order by  TS_ABS desc", conn);
                    dr.Close();
                    dr = command1.ExecuteReader();
                    while (dr.Read())
                    {
                        dtSortDateTo.Text = dr["TS_ABS"].ToString();
                        strEndDate = dr["TS_ABS"].ToString();
                    }


                    //cmbProductCode.Items.Add("All");
                    //cmbProductCode.Text = "All";
                    dr.Close();
                    conn.Close();

                }
                using (var conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    
                    SqlCommand command2 = new SqlCommand("insert into DataGrid select " + intCounterDataGridViews + ",'" + cmbProductCode.Text + "'", conn);
                    var cmdSelectFromProduct = command2.ExecuteScalar();
                    command2.CommandTimeout = 600;



                    command2 = new SqlCommand("ProdUpdate", conn);
                    command2.CommandType = CommandType.StoredProcedure;

                    command2.ExecuteNonQuery();
                  

                    conn.Close();


                }
                //arrDataGridView[intCounterDataGridViews].Columns.Add("zzz1", "777");
                //arrDataGridView[intCounterDataGridViews].Enabled = true;
                //arrDataGridView[intCounterDataGridViews].Visible = true;
                //arrDataGridView[intCounterDataGridViews].Top = dataGridViewRaw.Top;
                //arrDataGridView[intCounterDataGridViews].Left = dataGridViewRaw.Left;
                //intCounterDataGridViews++;
                if ((cmbProductCode.Text == "") && cmbProductCode.Items.Count < 2)
                {
                    cmbProductCode.Text = cmbProductCode.Items[cmbProductCode.Items.Count - 1].ToString();
                }
                MessageBox.Show("Action : <Select Data Raw file> is finished!");
            }
            catch (Exception ex)
            {
                var strEx = ex.ToString();
                if (strEx.IndexOf("The process cannot access the file") > 0)
                    MessageBox.Show("Please close the Excel file " + strFileName + " before reading the data into the application") ;
                else
                MessageBox.Show(ex.ToString());


            }
            if (clbVirtOzid.Items.Count > 0) btnReset.Enabled = true;
            if (cmbProductCode.Items.Count > 1)
            {

                for (int i = 0; i < cmbProductCode.Items.Count; i++)
                {
                    for (int j = cmbProductCode.Items.Count - 1; j >= 0; --j)
                    {
                        if ((string)cmbProductCode.Items[j] == (string)cmbProductCode.Items[i])
                        {
                            cmbProductCode.Items.Remove(j);

                        }
                    }
                }
                                   
                        

            }
            chkAllVirtOzid.Checked = true;
        }

        private void btnCalculationView_Click_1(object sender, EventArgs e) // Calculation View
        {
            try
            {
                tabMain.SelectTab("ViewCalculation");
            frmSave_Load(sender, e);
                MessageBox.Show("Action : <Calculation View> is finished!");
            }
            catch (Exception ex)
            {              
                   MessageBox.Show(ex.ToString());
            }

        }

        private void btnCalculationSearch_Click(object sender, EventArgs e) // Calculation search
        {
            try
            {
                tabMain.SelectTab("SearchCalculation");
            frmHistResults_Load(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnSave_Click(object sender, EventArgs e) // Saving data for every row 
        {

            try
            {
                string INIfolderPath;

                //Reading data from app.ini file
                INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                INIfolderPath = INIfolderPath + "\\app.ini";

                string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                connectionString = lines[0];


                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }

                if (!IsCalcIDAvailable)
                {

                    SqlCommand command = null;
                    using (SqlConnection connection = new SqlConnection(
                      connectionString))
                    {
                        connection.Open();




                        NParameterTotal = txtNParameterTotal.Text;
                        NStatistically = txtNStatistically.Text;
                        PercentStatistically = txtPercentStatistically.Text;
                        DoNotFitStatistically = txtDoNotFitStatistically.Text;
                        CalcID = cmbCalcID.Text;
                        User = txtUser.Text;
                        TimePointData = txtTimePointData.Text;
                        TimePointCalc = txtTimePointCalc.Text;
                        Note = txtNote.Text;
                        if (this.chkActive.Checked)
                        {
                            Active = 1;
                        }
                        else
                        {
                            Active = 0;
                        }



                        string qry = "insert into CalcResultView (ID, NParameterTotal, ";
                        qry += "NStatistically, PercentStatistically, DoNotFitStatistically, CalcID, [User],TimePointData,TimePointCalc, Note,Active) values(";
                        qry += " '" + @ID + "','" + @NParameterTotal + "','" + @NStatistically + "','" + @PercentStatistically + "','" + @DoNotFitStatistically + "','" + @CalcID + "','" + @User;

                        qry += "','" + @TimePointData;
                        qry += "','" + @TimePointCalc;
                        qry += "','" + @Note;
                        qry += "','" + @Active + "')";

                        command = new SqlCommand(qry, connection);

                        command.Parameters.Add("@ID",
                                                        SqlDbType.Int).Value = @ID;
                        command.Parameters.Add("@NParameterTotal",
                            SqlDbType.NVarChar, 50).Value = @NParameterTotal;
                        command.Parameters.Add("@NStatistically",
                             SqlDbType.NVarChar, 50).Value = @NStatistically;
                        command.Parameters.Add("@PercentStatistically",
                             SqlDbType.NVarChar, 50).Value = @PercentStatistically;
                        command.Parameters.Add("@DoNotFitStatistically",
                             SqlDbType.NVarChar, 50).Value = @DoNotFitStatistically;
                        command.Parameters.Add("@CalcID",
                             SqlDbType.NVarChar, 50).Value = @CalcID;
                        command.Parameters.Add("@User",
                             SqlDbType.NVarChar, 50).Value = @User;
                        command.Parameters.Add("@TimePointData",
                             SqlDbType.NVarChar, 50).Value = @TimePointData;
                        command.Parameters.Add("@TimePointCalc",
                             SqlDbType.NVarChar, 50).Value = @TimePointCalc;
                        command.Parameters.Add("@Note",
                             SqlDbType.NVarChar, 250).Value = @Note;
                        command.Parameters.Add("@Active",
                             SqlDbType.NVarChar, 50).Value = @Active;
                        command.ExecuteNonQuery();

                        if (Int16.Parse(txtNParameterTotal.Text) > 0)
                        {
                            for (int i = 0; i < Int16.Parse(txtNParameterTotal.Text); i++)
                            {
                                OZID = strMatrix[i, 1];
                                CalcID = cmbCalcID.Text;
                                TotalN = strMatrix[i, 2];
                                KPI0 = strMatrix[i, 3];
                                KPI1 = strMatrix[i, 4];
                                KPI2 = strMatrix[i, 5];
                                KPI3 = strMatrix[i, 6];



                                FitStatistically = strMatrix[i, 7];

                                RelevantForDiscussion = strMatrix[i, 8];
                                Additional_note = strMatrix[i, 9];

                                var sql = "select ID from Graphs where VIRT_OZID ='" + @OZID + "'" + " and CalcID = '" + CalcID + "'";
                                command = new SqlCommand(sql, connection);
                                SqlDataReader dr = command.ExecuteReader();

                                SqlCommand cmd = new SqlCommand(sql, connection);


                                while (dr.Read())
                                {
                                    @GraphID = dr["ID"].ToString();

                                }


                                dr.Close();
                                //conn.Close();




                                qry = "insert into VIRT_OZID_per_calculation (VIRT_OZID,CalcID,TotalN,KPI0,KPI1,KPI2,KPI3,Additional_note,FitStatistically,RelevantForDiscussion,GraphID) ";
                                qry += " values(";
                                qry += " '" + @OZID + "','" + @CalcID + "','" + @TotalN + "','" + @KPI0 + "','" + @KPI1 + "','" + @KPI2 + "','" + @KPI3;
                                qry += "','" + @Additional_note;
                                qry += "','" + @FitStatistically;
                                qry += "','" + @RelevantForDiscussion;
                                qry += "','" + @GraphID + "')";

                                command = new SqlCommand(qry, connection);

                                command.Parameters.Add("@VIRT_OZID",
                                    SqlDbType.NVarChar, 250).Value = @OZID;
                                command.Parameters.Add("@CalcID",
                                    SqlDbType.NVarChar, 50).Value = @CalcID;
                                command.Parameters.Add("@TotalN",
                                     SqlDbType.NVarChar, 50).Value = @TotalN;
                                command.Parameters.Add("@KPI0",
                                     SqlDbType.NVarChar, 50).Value = @KPI0;
                                command.Parameters.Add("@KPI1",
                                     SqlDbType.NVarChar, 50).Value = @KPI1;
                                command.Parameters.Add("@KPI2",
                                     SqlDbType.NVarChar, 50).Value = @KPI2;
                                command.Parameters.Add("@KPI3",
                                     SqlDbType.NVarChar, 50).Value = @KPI3;
                                command.Parameters.Add("@Additional_note",
                                     SqlDbType.NVarChar, 50).Value = @Additional_note;
                                command.Parameters.Add("@FitStatistically",
                                     SqlDbType.NVarChar, 50).Value = @FitStatistically;
                                command.Parameters.Add("@RelevantForDiscussion",
                                     SqlDbType.NVarChar, 250).Value = @RelevantForDiscussion;
                                command.Parameters.Add("@GraphID",
                                     SqlDbType.NVarChar, 50).Value = @GraphID;
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
                else
                {
                    SqlCommand command = null;
                    using (SqlConnection connection = new SqlConnection(
                      connectionString))
                    {
                        connection.Open();
                        NParameterTotal = txtNParameterTotal.Text;
                        NStatistically = txtNStatistically.Text;
                        PercentStatistically = txtPercentStatistically.Text;
                        DoNotFitStatistically = txtDoNotFitStatistically.Text;
                        CalcID = cmbCalcID.Text;
                        User = txtUser.Text;
                        TimePointData = txtTimePointData.Text;
                        TimePointCalc = txtTimePointCalc.Text;
                        Note = txtNote.Text;
                        if (this.chkActive.Checked)
                        {
                            Active = 1;
                        }
                        else
                        {
                            Active = 0;
                        }

                        //command.Parameters.Add("@ReportsTo",
                        //    SqlDbType.Int).Value = reportsTo;

                        //command.Parameters.Add("@Photo",
                        //    SqlDbType.Image, photo.Length).Value = photo;

                        var qry = "update CalcResultView ";

                        qry += " set ID='" + @ID + "',";
                        qry += " NParameterTotal='" + @NParameterTotal + "',";
                        qry += " NStatistically='" + @NStatistically + "',";
                        qry += " PercentStatistically='" + @PercentStatistically + "',";
                        qry += " DoNotFitStatistically='" + @DoNotFitStatistically + "',";
                        qry += " CalcID='" + @CalcID + "',";
                        qry += " [User]='" + @User + "',";
                        qry += " TimePointData='" + @TimePointData + "',";
                        qry += " TimePointCalc='" + @TimePointCalc + "',";
                        qry += " Note='" + @Note + "',";
                        qry += " Active  ='" + @Active + "' where " + " CalcID='" + @CalcID + "'";

                        command = new SqlCommand(qry, connection);

                        command.Parameters.Add("@ID",
                            SqlDbType.Int).Value = @ID;
                        command.Parameters.Add("@NParameterTotal",
                            SqlDbType.NVarChar, 50).Value = @NParameterTotal;
                        command.Parameters.Add("@NStatistically",
                             SqlDbType.NVarChar, 50).Value = @NStatistically;
                        command.Parameters.Add("@PercentStatistically",
                             SqlDbType.NVarChar, 50).Value = @PercentStatistically;
                        command.Parameters.Add("@DoNotFitStatistically",
                             SqlDbType.NVarChar, 50).Value = @DoNotFitStatistically;
                        command.Parameters.Add("@CalcID",
                             SqlDbType.NVarChar, 50).Value = @CalcID;
                        command.Parameters.Add("@User",
                             SqlDbType.NVarChar, 50).Value = @User;
                        command.Parameters.Add("@TimePointData",
                             SqlDbType.NVarChar, 50).Value = @TimePointData;
                        command.Parameters.Add("@TimePointCalc",
                             SqlDbType.NVarChar, 50).Value = @TimePointCalc;
                        command.Parameters.Add("@Note",
                             SqlDbType.NVarChar, 250).Value = @Note;
                        command.Parameters.Add("@Active",
                             SqlDbType.NVarChar, 50).Value = @Active;

                        command.ExecuteNonQuery();


                        ////000000000000000000000
                        if (Int16.Parse(txtNParameterTotal.Text) > 0)
                        {
                            for (int i = 0; i < Int16.Parse(txtNParameterTotal.Text); i++)
                            {

                                if (rbAll.Checked == true)
                                //&& (rbFitStat.Checked != true)) rbNotFitStat.Checked = true; 
                                {

                                }
                                OZID = strMatrix[i, 1];
                                CalcID = cmbCalcID.Text;
                                TotalN = strMatrix[i, 2];
                                KPI0 = strMatrix[i, 3];
                                KPI1 = strMatrix[i, 4];
                                KPI2 = strMatrix[i, 5];
                                KPI3 = strMatrix[i, 6];



                                FitStatistically = strMatrix[i, 7];

                                RelevantForDiscussion = strMatrix[i, 8];
                                Additional_note = strMatrix[i, 9];

                                qry = "update VIRT_OZID_per_calculation ";

                                qry += " set VIRT_OZID='" + @OZID + "',";
                                qry += " CalcID='" + cmbCalcID.Text + "',";
                                qry += " TotalN='" + @TotalN + "',";
                                qry += " KPI0='" + @KPI0 + "',";
                                qry += " KPI1='" + @KPI1 + "',";
                                qry += " KPI2='" + @KPI2 + "',";
                                qry += " KPI3='" + @KPI3 + "',";
                                qry += " FitStatistically='" + @FitStatistically + "',";
                                qry += " RelevantForDiscussion='" + @RelevantForDiscussion + "',";
                                qry += " Additional_note='" + @Additional_note + "' where VIRT_OZID='" + OZID + "' and CalcID='" + cmbCalcID.Text + "'";

                                if (strMatrix[i, 1] != null)
                                {
                                    command = new SqlCommand(qry, connection);

                                    command.Parameters.Add("@VIRT_OZID",
                                        SqlDbType.NVarChar, 250).Value = @OZID;
                                    command.Parameters.Add("@CalcID",
                                        SqlDbType.NVarChar, 50).Value = @CalcID;
                                    command.Parameters.Add("@TotalN",
                                         SqlDbType.NVarChar, 50).Value = @TotalN;
                                    command.Parameters.Add("@KPI0",
                                         SqlDbType.NVarChar, 50).Value = @KPI0;
                                    command.Parameters.Add("@KPI1",
                                         SqlDbType.NVarChar, 50).Value = @KPI1;
                                    command.Parameters.Add("@KPI2",
                                         SqlDbType.NVarChar, 50).Value = @KPI2;
                                    command.Parameters.Add("@KPI3",
                                         SqlDbType.NVarChar, 50).Value = @KPI3;
                                    command.Parameters.Add("@Additional_note",
                                         SqlDbType.NVarChar, 50).Value = @Additional_note;
                                    command.Parameters.Add("@FitStatistically",
                                         SqlDbType.NVarChar, 50).Value = @FitStatistically;
                                    command.Parameters.Add("@RelevantForDiscussion",
                                         SqlDbType.NVarChar, 250).Value = @RelevantForDiscussion;

                                    command.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }






        }

        private void frmSave_Load(object sender, EventArgs e) //Uploading data on the last calculation and build the quantity of the rows in parameters tab 
        {
            
            string INIfolderPath;
            dataGridViewRaw.Visible = true;
            // Current directory
            INIfolderPath = System.IO.Directory.GetCurrentDirectory();
            // Input data file app.ini
            INIfolderPath = INIfolderPath + "\\app.ini";

            string[] lines = System.IO.File.ReadAllLines(INIfolderPath);

            connectionString = lines[0];// Connection String            
            var strRscript = lines[1];// Second line of app.ini file          
            var strRpath = lines[2]; // Third line of app.ini file            
            var strDataDir = lines[3];// Forth line of app.ini file          
            var strOutputDir = lines[4];// Fifth line of app.ini file
            strOutPutPath = lines[4];

            // Uploading Calculation ID in the ComboBox 
            try
            {
                // Create the ToolTip and associate with the Form container.
                ToolTip toolTip1 = new ToolTip();

                // Set up the delays for the ToolTip.
                toolTip1.AutoPopDelay = 5000;
                toolTip1.InitialDelay = 1000;
                toolTip1.ReshowDelay = 500;
                // Force the ToolTip text to be displayed whether or not the form is active.
                toolTip1.ShowAlways = true;

                // Set up the ToolTip text for the Button and Checkbox.
                toolTip1.SetToolTip(this.chkActive, "Set the calculation to be in active status");

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct CalcID, CalcDate from CalculationRaw where ProductCode = '"+ cmbProductCode.Text + "' order by CalcDate";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();
                    table.Columns.Add("CalcID", typeof(string));
                    SqlDataReader dr = command1.ExecuteReader();
                    SqlCommand cmd = new SqlCommand(qry, conn);

                    //if (cmbCalcID.Items.Count ==0)

                    cmbCalcID.Items.Clear();
                    {     
                    while (dr.Read())
                    {
                        cmbCalcID.Items.Add(dr["CalcID"].ToString());

                    }
                    }
                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void cmbCalcID_SelectedIndexChanged(object sender, EventArgs e) // Action on the selection item from the 
        {
           
            {
             
                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID2.Text + "'";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);




                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {


                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr["VIRT_OZID"].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr["VIRT_OZID"].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr["VIRT_OZID"].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                //Text = string.Format("txt{0}", i),
                                Enabled = false,
                                Width = 50,
                                Text = dr["TotalN"].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr["TotalN"].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI0"].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr["KPI0"].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI1"].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr["KPI1"].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI2"].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr["KPI2"].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI3"].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr["KPI3"].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }

                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);


                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr["Additional_note"].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };

                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);

                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);
                            i++;
                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID2.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID2.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }







                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID2.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                            txtNStatistically.Text = dr["NStatistically"].ToString();
                            txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                            txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                            cmbCalcID.Text = dr["CalcID"].ToString();
                            txtUser.Text = dr["User"].ToString();
                            txtTimePointData.Text = dr["TimePointData"].ToString();
                            txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                            txtNote.Text = dr["Note"].ToString();
                            if (dr["Active"].ToString() == "0")
                                chkActive.Checked = false;
                            else
                                chkActive.Checked = true;

                        }


                        dr.Close();
                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + cmbCalcID2.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        


                        dr.Close();
                        conn.Close();
                    }




                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }




        }


        void chk_MouseMove(object sender, EventArgs e) // Tooltips on CheckBoxes  "Fit statistically" and "Relevant for disc" 
        {
            try
            {
                ToolTip toolTip1 = new ToolTip();
                System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)sender;

                //CheckBox "Fit statistically" 
                if (chk.Name.Substring(0, 3) == "ch1")
                {
                    

                    // Set up the delays for the ToolTip.
                    toolTip1.AutoPopDelay = 5000;
                    toolTip1.InitialDelay = 1000;
                    toolTip1.ReshowDelay = 500;
                    // Force the ToolTip text to be displayed whether or not the form is active.
                    toolTip1.ShowAlways = true;

                    // Set up the ToolTip text for the Button and Checkbox.
                    toolTip1.SetToolTip(chk, "This checkbox displays Fitstatistically status");

                }
                //CheckBox "Relevant for disc"
                if (chk.Name.Substring(0, 3) == "ch2")
                {

                    // Set up the delays for the ToolTip.
                    toolTip1.AutoPopDelay = 5000;
                    toolTip1.InitialDelay = 1000;
                    toolTip1.ReshowDelay = 500;
                    // Force the ToolTip text to be displayed whether or not the form is active.
                    toolTip1.ShowAlways = true;

                    // Set up the ToolTip text for the Button and Checkbox.
                    toolTip1.SetToolTip(chk, "This checkbox displays Relevant For Discussion status");

                }
                //CheckBox "Active"
                if (chk.Name.Substring(0, 3) == "ch3")
                {
                    // Set up the delays for the ToolTip.
                    toolTip1.AutoPopDelay = 5000;
                    toolTip1.InitialDelay = 1000;
                    toolTip1.ReshowDelay = 500;
                    // Force the ToolTip text to be displayed whether or not the form is active.
                    toolTip1.ShowAlways = true;

                    // Set up the ToolTip text for the Button and Checkbox.
                    toolTip1.SetToolTip(chk, "This checkbox displays Active status of the calculation");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        void chk_Click(object sender, EventArgs e) // Actions on CheckBoxes  "Fit statistically" and "Relevant for disc"       
        {
            try { 
            System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)sender;

            //CheckBox "Fit statistically" 
            if (chk.Name.Substring(0, 3) == "ch1")
            {
                var num = Int16.Parse(chk.Name.Substring(3, chk.Name.Length - 3));
                if (chk.Checked == true)
                {
                    strMatrix[num, 7] = "True";
                }
                else
                {
                    strMatrix[num, 7] = "False";
                }

            }
            //CheckBox "Relevant for disc"
            if (chk.Name.Substring(0, 3) == "ch2")
            {
                var num = Int16.Parse(chk.Name.Substring(3, chk.Name.Length - 3));
                if (chk.Checked == true)
                {
                    strMatrix[num, 8] = "True";
                }
                else
                {
                    strMatrix[num, 8] = "False";
                }

            }
            //CheckBox "Active"
            if (chk.Name.Substring(0, 3) == "ch3")
            {
                var num = Int16.Parse(chk.Name.Substring(3, chk.Name.Length - 3));
                if (chk.Checked == true)
                {
                    strMatrix[num, 9] = "True";
                }
                else
                {
                    strMatrix[num, 9] = "False";
                }

            }
        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
}

        void tb_LostFocus(object sender, EventArgs e) // Handling Note fields
        {
            System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
            // Define the number of the edited row
            var num = Int16.Parse(tb.Name.Substring(3, tb.Name.Length - 3));
            // Saving data to global matrix for saving data to DataBase         
            strMatrix[num, 9] = tb.Text;


        }





        public void ba34_Click2(object sender, EventArgs e) // pressing Chart buttons
        {
            pictureBox3.Image = null;
            pictureBox4.Image = null;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;

            blnPressed = false;
            try
            {


                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();
               

                System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
               
                var c = btn.Name.Substring(1, btn.Name.Length - 1);
                int num = Int16.Parse(c);
                if (num < 20)
                    strOZID = strArr2[num - 10];
                if ((num < 200) && (num > 19))
                    strOZID = strArr2[num - 100];

                //=======================================
                try
                {


                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID3.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count>0)

                        
                        {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count-1]["ImageValue"];
                     
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);

                        pictureBox3.Visible = true;
                        pictureBox3.Image = Image.FromStream(newImageStream, true);
                    }



                }

                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================
                try
                {

                   
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID4.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count-1]["ImageValue"];
                       
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);

                        pictureBox4.Visible = true;
                        pictureBox4.Image = Image.FromStream(newImageStream, true);
                    }



                }

                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }


            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }

        public void ba34_Click(object sender, EventArgs e) // pressing Chart buttons
        {
            pictureBox3.Image = null;
            pictureBox4.Image = null;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;

            blnPressed = false;
            try
            {
                

                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();
               

                System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                //MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                var c = btn.Name.Substring(1, btn.Name.Length - 1);
                int num = Int16.Parse(c);
                if (num < 20)
                    //MessageBox.Show(strArr2[num-10]);
                    strOZID = strArr2[num - 10];
                if ((num < 200) && (num > 19))


                    strOZID = strArr2[num - 100];




                //=======================================
                try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID3.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        //var stream = new MemoryStream(data);
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                        btnZoomOut3.Visible = true;
                        btnPrint3.Visible = true;
                        pictureBox3.Visible = true;
                        pictureBox3.Image = Image.FromStream(newImageStream, true);
                    }

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                   cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID4.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    da = new SqlDataAdapter(cmd2);
                     ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        //var stream = new MemoryStream(data);
                        //pictureBox4.Image = Image.FromStream(stream, true);
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                        btnZoomOut4.Visible = true;
                        btnPrint4.Visible = true;
                        pictureBox4.Visible = true;
                        pictureBox4.Image = Image.FromStream(newImageStream, true);
                    }
                    ba34_Click2(sender, e);
                }

                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================



            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }
        public void ba1_Click(object sender, EventArgs e) // pressing Chart buttons
        {
            try
            {
                btnZoomIn1.Visible = true;
                pictureBox2.Visible = true;
                btnPrint.Visible = true;
                label82.Visible = true;
                label83.Visible = true;
                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();
                //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";

                blnPressed = false;
                System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                //MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                var c = btn.Name.Substring(1, btn.Name.Length - 1);
                int num = Int16.Parse(c);
                if (num < 20)
                    //MessageBox.Show(strArr2[num-10]);
                    strOZID = strArr2[num - 10];
                if ((num < 200) && (num > 19))


                    strOZID = strArr2[num - 100];

               


                //=======================================
                 try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID2.Text + "' and VIRT_OZID='" + strOZID + "'", CN);
                    

                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        var stream = new MemoryStream(data);
                        pictureBox2.Image = Image.FromStream(stream);
                        btnPrint.Visible = true;
                    }


                    
                }
                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================



            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }

        private void ba_Zoom(object sender, EventArgs e) // pressing Chart buttons
        {
           
            try
            {
                
                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();

                var cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                var da = new SqlDataAdapter(cmd2);
                var ds = new DataSet();
                da.Fill(ds, "Graphs");
                var count = ds.Tables["Graphs"].Rows.Count;

                if (count > 0)
                {
                    var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                   
                    System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                 
                    PictureBox pb = new PictureBox();
                pb.Image = Image.FromStream(newImageStream, true);
                    pb.Location = new System.Drawing.Point(3, 3);
                pb.Size = new Size(1100, 900);
              
                pb.SizeMode = PictureBoxSizeMode.StretchImage;
               
                pb.Refresh();
                Form frm = new Form();
                frm.Size = new Size(1200, 1000);
                frm.Controls.Add(pb);
                frm.ShowDialog();
                    CN.Close();
                }
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }
       
        private void ba_Click(object sender, EventArgs e) // pressing Chart buttons
        {
            //picGraph.Visible = true;
            //picGraph.Image = null;
            //try
            //{
            btnZoom1.Visible = true;

            //    SqlConnection CN = new SqlConnection(connectionString);
            //    CN.Open();
            //    //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";


            //    System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            //    //MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
            //    var c = btn.Name.Substring(1, btn.Name.Length - 1);
            //    int num = Int16.Parse(c);
            //    if (num < 20)
            //        //MessageBox.Show(strArr2[num-10]);
            //        strOZID = strArr2[num - 10];
            //    if ((num < 200) && (num > 19))


            //        strOZID = strArr2[num - 100];

            //    //SqlCommand cmd2 = new SqlCommand("select top 1 GraphName from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" +  + "'", CN);

            //    SqlCommand cmd2 = new SqlCommand("select top 1 GraphName from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" + strOZID + "'", CN);
            //    //SqlDataReader dr2 = cmd2.ExecuteReader();
            //    picGraph.Visible = true;
            //    CN = new SqlConnection(connectionString);
            //    CN.Open();




            //    var da = new SqlDataAdapter(cmd2);
            //    var ds = new DataSet();
            //    da.Fill(ds, "Graphs");
            //    int count = ds.Tables["Graphs"].Rows.Count;

            //    if (count > 0)
            //    {
            //        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
            //        //var stream = new MemoryStream(data);
            //        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
            //        pictureBox3.Image = Image.FromStream(newImageStream, true);
            //    }

            //    picGraph.Visible = true;
            //    CN = new SqlConnection(connectionString);
            //    CN.Open();

            //    cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


            //    da = new SqlDataAdapter(cmd2);
            //    ds = new DataSet();
            //    da.Fill(ds, "Graphs");
            //    count = ds.Tables["Graphs"].Rows.Count;

            //    if (count > 0)
            //    {
            //        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
            //        //var stream = new MemoryStream(data);
            //        //pictureBox4.Image = Image.FromStream(stream, true);
            //        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
            //        picGraph.Image = Image.FromStream(newImageStream, true);
            //    }
            //}
            picGraph.Image = null;
            picGraph.Visible = false;
            picGraph.Image = null;
            try
            {


                SqlConnection CN = new SqlConnection(connectionString);
                CN.Open();
                //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";


                System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                //MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                var c = btn.Name.Substring(1, btn.Name.Length - 1);
                int num = Int16.Parse(c);
                if (num < 20)
                    //MessageBox.Show(strArr2[num-10]);
                    strOZID = strArr2[num - 10];
                if ((num < 200) && (num > 19))


                    strOZID = strArr2[num - 100];




                //=======================================
                try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    //if (count > 0)
                    //{
                    //    var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                    //    //var stream = new MemoryStream(data);
                    //    System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                    //    pictureBox3.Image = Image.FromStream(newImageStream, true);
                    //}

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    da = new SqlDataAdapter(cmd2);
                    ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        //var stream = new MemoryStream(data);
                        //pictureBox4.Image = Image.FromStream(stream, true);
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                        picGraph.Image = Image.FromStream(newImageStream, true);
                        btnPrint2.Visible = true;
                    }
                }
                catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
                }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }

        private void rbFitStat_CheckedChanged(object sender, EventArgs e) // filter table view fit statistically

        {
            try
            {
                strMatrix = new string[142, 11];
                if ((rbNotFitStat.Checked != true) && (rbAll.Checked != true))
                { rbFitStat.Checked = true; }

                picGraph.Visible = false;
                IsCalcIDAvailable = false;
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            if (!IsCalcIDAvailable)
            {
                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID.Text + "') and CalcID = '" + cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr[1].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr[1].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[1].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[2].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr[2].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[3].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr[3].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[4].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr[4].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[5].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr[5].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[6].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr[6].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            if ((dr[7] == "true") || (dr[7] == "True"))
                            {
                                ch1.Checked = true;


                            }


                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                //Text = string.Format("txt{0}", i),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };
                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);
                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }

            }
            else
            {

                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "' and FitStatistically = 'true'";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);




                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr["VIRT_OZID"].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr["VIRT_OZID"].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr["VIRT_OZID"].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                //Text = string.Format("txt{0}", i),
                                Enabled = false,
                                Width = 50,
                                Text = dr["TotalN"].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr["TotalN"].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI0"].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr["KPI0"].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI1"].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr["KPI1"].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI2"].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr["KPI2"].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI3"].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr["KPI3"].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }

                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);

                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr["Additional_note"].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };

                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);

                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }







                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                            txtNStatistically.Text = dr["NStatistically"].ToString();
                            txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                            txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                            cmbCalcID.Text = dr["CalcID"].ToString();
                            txtUser.Text = dr["User"].ToString();
                            txtTimePointData.Text = dr["TimePointData"].ToString();
                            txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                            txtNote.Text = dr["Note"].ToString();
                            if (dr["Active"].ToString() == "0")
                                chkActive.Checked = false;
                            else
                                chkActive.Checked = true;

                        }


                        dr.Close();
                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        //while (dr.Read())
                        //{
                        //    txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                        //    txtNStatistically.Text = dr["NStatistically"].ToString();
                        //    txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                        //    txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                        //    cmbCalcID.Text = dr["CalcID"].ToString();
                        //    txtUser.Text = dr["User"].ToString();
                        //    txtTimePointData.Text = dr["TimePointData"].ToString();
                        //    txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                        //    txtNote.Text = dr["Note"].ToString();
                        //    if (dr["Active"].ToString() == "0")
                        //        chkActive.Checked = false;
                        //    else
                        //        chkActive.Checked = true;

                        //}


                        dr.Close();
                        conn.Close();
                    }




                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void rbNotFitStat_CheckedChanged(object sender, EventArgs e) // filter table view not fit statistically
        {
            try
            {
                strMatrix = new string[142, 11];
                if ((rbAll.Checked != true) && (rbFitStat.Checked != true))
                { rbNotFitStat.Checked = true; }

                picGraph.Visible = false;
                IsCalcIDAvailable = false;
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            if (!IsCalcIDAvailable)
            {
                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID.Text + "') and CalcID = '" + cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr[1].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr[1].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[1].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[2].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr[2].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[3].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr[3].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[4].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr[4].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[5].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr[5].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[6].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr[6].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                //Text = string.Format("txt{0}", i),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);
                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }

            }
            else
            {

                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "' and FitStatistically = 'false'";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);




                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr["VIRT_OZID"].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr["VIRT_OZID"].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr["VIRT_OZID"].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                //Text = string.Format("txt{0}", i),
                                Enabled = false,
                                Width = 50,
                                Text = dr["TotalN"].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr["TotalN"].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI0"].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr["KPI0"].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI1"].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr["KPI1"].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI2"].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr["KPI2"].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI3"].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr["KPI3"].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            if (dr["FitStatistically"].ToString() == "true")
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }

                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);

                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr["Additional_note"].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };

                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);

                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }







                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                            txtNStatistically.Text = dr["NStatistically"].ToString();
                            txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                            txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                            cmbCalcID.Text = dr["CalcID"].ToString();
                            txtUser.Text = dr["User"].ToString();
                            txtTimePointData.Text = dr["TimePointData"].ToString();
                            txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                            txtNote.Text = dr["Note"].ToString();
                            if (dr["Active"].ToString() == "0")
                                chkActive.Checked = false;
                            else
                                chkActive.Checked = true;

                        }


                        dr.Close();
                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        //while (dr.Read())
                        //{
                        //    txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                        //    txtNStatistically.Text = dr["NStatistically"].ToString();
                        //    txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                        //    txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                        //    cmbCalcID.Text = dr["CalcID"].ToString();
                        //    txtUser.Text = dr["User"].ToString();
                        //    txtTimePointData.Text = dr["TimePointData"].ToString();
                        //    txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                        //    txtNote.Text = dr["Note"].ToString();
                        //    if (dr["Active"].ToString() == "0")
                        //        chkActive.Checked = false;
                        //    else
                        //        chkActive.Checked = true;

                        //}


                        dr.Close();
                        conn.Close();
                    }




                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void rbAll_CheckedChanged(object sender, EventArgs e) // filter table view All
        {
            try
            {
                strMatrix = new string[142, 11];
                if ((rbNotFitStat.Checked != true) && (rbFitStat.Checked != true))
                { rbAll.Checked = true; }

                picGraph.Visible = false;
                IsCalcIDAvailable = false;
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            if (!IsCalcIDAvailable)
            {
                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID.Text + "') and CalcID = '" + cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr[1].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr[1].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[1].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[2].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr[2].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[3].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr[3].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[4].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr[4].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[5].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr[5].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[6].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr[6].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                //Text = string.Format("txt{0}", i),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);
                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }

            }
            else
            {

                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "' ";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);




                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr["VIRT_OZID"].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr["VIRT_OZID"].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr["VIRT_OZID"].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                //Text = string.Format("txt{0}", i),
                                Enabled = false,
                                Width = 50,
                                Text = dr["TotalN"].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr["TotalN"].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI0"].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr["KPI0"].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI1"].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr["KPI1"].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI2"].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr["KPI2"].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI3"].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr["KPI3"].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }

                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };

                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);

                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr["Additional_note"].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };

                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);

                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }







                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                            txtNStatistically.Text = dr["NStatistically"].ToString();
                            txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                            txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                            cmbCalcID.Text = dr["CalcID"].ToString();
                            txtUser.Text = dr["User"].ToString();
                            txtTimePointData.Text = dr["TimePointData"].ToString();
                            txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                            txtNote.Text = dr["Note"].ToString();
                            if (dr["Active"].ToString() == "0")
                                chkActive.Checked = false;
                            else
                                chkActive.Checked = true;

                        }


                        dr.Close();
                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);



                        dr.Close();
                        conn.Close();
                    }




                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void cmbCalcID_SelectedIndexChanged_1(object sender, EventArgs e) // Selection of calculation ID
        {
            ToolTip toolTip61 = new ToolTip();


            //CheckBox "Fit statistically" 



            // Set up the delays for the ToolTip.
            toolTip61.AutoPopDelay = 5000;
            toolTip61.InitialDelay = 1000;
            toolTip61.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip61.ShowAlways = true;
            //txtNStatistically.Enabled = true;
            // Set up the ToolTip text for the textbox
            toolTip61.SetToolTip(label8, "The calculation of KPI is initial and does not change");
            
            btnZoom1.Visible = false;
            btnPrint2.Visible = false;
            try
            {
                rbAll.Checked = true;
                picGraph.Visible = false;
                IsCalcIDAvailable = false;
                rbAll.Enabled = true;
                rbFitStat.Enabled = true;
                rbNotFitStat.Enabled = true;
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            if (!IsCalcIDAvailable)
            {
                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";
                        
                        qry = "select * FROM[dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        SqlDataReader dr = command1.ExecuteReader();
                        int j = 0;
                        while (dr.Read())
                        {
                            j++;
                        }
                        dr.Close();
                        if (j==0)
                            qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID.Text + "') and CalcID = '" + cmbCalcID.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
                        else
                            qry = "select * FROM[dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "'";


                        command1 = new SqlCommand(qry, conn);

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                      dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr[1].ToString(),
                                Top = i * 20,
                                Left = 3,
                                BackColor = Color.White,
                                
                            };
                            strMatrix[i, 1] = dr[1].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[1].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[2].ToString(),
                                Top = i * 20,
                                Left = 110,
                                BackColor = Color.White,
                            };
                            strMatrix[i, 2] = dr[2].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[3].ToString(),
                                Top = i * 20,
                                Left = 190,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 3] = dr[3].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[4].ToString(),
                                Top = i * 20,
                                Left = 270,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 4] = dr[4].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[5].ToString(),
                                Top = i * 20,
                                Left = 360,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 5] = dr[5].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[6].ToString(),
                                Top = i * 20,
                                Left = 440,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 6] = dr[6].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True") || (strMatrix[i, 3].ToString() == "100"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }
                            
                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);
                            ch1.MouseLeave += new EventHandler(chk_MouseMove);

                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            ch2.MouseLeave += new EventHandler(chk_MouseMove);
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                //Text = string.Format("{0}", "yes"),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);
                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);
                            i++;
                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }

            }
            else
            {

                try
                {
                    txtNParameterTotal.Text = "";
                    txtNStatistically.Text = "";
                    txtPercentStatistically.Text = "";
                    txtDoNotFitStatistically.Text = "";
                    //cmbCalcID.Text = "";
                    txtUser.Text = "";
                    txtTimePointData.Text = "";
                    txtTimePointCalc.Text = "";
                    txtNote.Text = "";
                    chkActive.Checked = false;
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel1.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[VIRT_OZID_per_calculation] where calcid = '" + cmbCalcID.Text + "'";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);




                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {


                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr["VIRT_OZID"].ToString(),
                                Top = i * 20,
                                Left = 3,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 1] = dr["VIRT_OZID"].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr["VIRT_OZID"].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                //Text = string.Format("txt{0}", i),
                                Enabled = false,
                                Width = 50,
                                Text = dr["TotalN"].ToString(),
                                Top = i * 20,
                                Left = 110,
                                BackColor = Color.White,
                            };
                            strMatrix[i, 2] = dr["TotalN"].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI0"].ToString(),
                                Top = i * 20,
                                Left = 190,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 3] = dr["KPI0"].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI1"].ToString(),
                                Top = i * 20,
                                Left = 270,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 4] = dr["KPI1"].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI2"].ToString(),
                                Top = i * 20,
                                Left = 360,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 5] = dr["KPI2"].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr["KPI3"].ToString(),
                                Top = i * 20,
                                Left = 440,
                                BackColor = Color.White,

                            };
                            strMatrix[i, 6] = dr["KPI3"].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.Click += new EventHandler(chk_Click);
                            ch1.MouseLeave += new EventHandler(chk_MouseMove);
                            
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True") || (strMatrix[i, 3].ToString() == "100"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }

                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };
                            
                            this.panel1.Controls.Add(b1);
                            b1.Click += new EventHandler(ba_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            ch2.MouseLeave += new EventHandler(chk_MouseMove);
                            ToolTip toolTip2 = new ToolTip();

                            // Set up the delays for the ToolTip.
                            toolTip2.AutoPopDelay = 5000;
                            toolTip2.InitialDelay = 1000;
                            toolTip2.ReshowDelay = 500;
                            // Force the ToolTip text to be displayed whether or not the form is active.
                            toolTip2.ShowAlways = true;

                            // Set up the ToolTip text for the Button and Checkbox.
                            toolTip2.SetToolTip(ch2, "This checkbox displays Relevant For Discussion status");

                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr["Additional_note"].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 900

                            };

                            strMatrix[i, 9] = tb7.Text;
                            tb7.LostFocus += new EventHandler(tb_LostFocus);

                            //Add data to matrix for inserting to table

                            panel1.Controls.Add(tb1); panel1.Controls.Add(tb2); panel1.Controls.Add(tb3); panel1.Controls.Add(tb4); panel1.Controls.Add(tb5); panel1.Controls.Add(tb6);
                            panel1.Controls.Add(ch1); panel1.Controls.Add(ch2); panel1.Controls.Add(tb7); panel1.Controls.Add(ba);
                            i++;
                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        // command1 = new SqlCommand(qry, conn);
                        //dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtTimePointCalc.Text = (string)dr[0];
                        }

                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();
                        //command1 = new SqlCommand(qry, conn);
                        //SqlDataReader dr1 = command1.ExecuteReader();

                        while (dr.Read())
                        {
                            txtNStatistically.Text = dr[0].ToString();
                        }

                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtPercentStatistically.Text = dr[0].ToString();
                        }
                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        while (dr.Read())
                        {
                            txtDoNotFitStatistically.Text = dr[0].ToString();
                        }
                        //------------------------------------------------





                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }







                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                            txtNStatistically.Text = dr["NStatistically"].ToString();
                            txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                            txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                            cmbCalcID.Text = dr["CalcID"].ToString();
                            txtUser.Text = dr["User"].ToString();
                            txtTimePointData.Text = dr["TimePointData"].ToString();
                            txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                            txtNote.Text = dr["Note"].ToString();
                            if (dr["Active"].ToString() == "False")
                                chkActive.Checked = false;
                            else
                                chkActive.Checked = true;

                        }


                        dr.Close();
                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from VIRT_OZID_per_calculation where CalcID = '" + cmbCalcID.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        //while (dr.Read())
                        //{
                        //    txtNParameterTotal.Text = dr["NParameterTotal"].ToString();
                        //    txtNStatistically.Text = dr["NStatistically"].ToString();
                        //    txtPercentStatistically.Text = dr["PercentStatistically"].ToString();
                        //    txtDoNotFitStatistically.Text = dr["DoNotFitStatistically"].ToString();
                        //    cmbCalcID.Text = dr["CalcID"].ToString();
                        //    txtUser.Text = dr["User"].ToString();
                        //    txtTimePointData.Text = dr["TimePointData"].ToString();
                        //    txtTimePointCalc.Text = dr["TimePointCalc"].ToString();
                        //    txtNote.Text = dr["Note"].ToString();
                        //    if (dr["Active"].ToString() == "0")
                        //        chkActive.Checked = false;
                        //    else
                        //        chkActive.Checked = true;

                        //}


                        dr.Close();
                        conn.Close();
                    }


                    MessageBox.Show("Loading done!");

                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
           
        }
        private void txtEntryDate_Validated(object sender, EventArgs e) // Validation of TimePoint date 
        {
            if (!string.IsNullOrEmpty(txtTimePointData.Text))
            {
                DateTime entryDate;
                if (DateTime.TryParseExact(txtTimePointData.Text, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out entryDate))
                {
                    string inputString = txtTimePointData.Text;
                    DateTime dDate;

                    if (DateTime.TryParse(inputString, out dDate))
                    {
                        var dateFormats = new[] { "dd.MM.yyyy" };
                        
                        //string readAddMeeting = Console.ReadLine();
                        DateTime validDateChecked;
                        DateTime validDateChecked2;
                        bool validDate = DateTime.TryParseExact(
                            inputString,
                            dateFormats,
                            DateTimeFormatInfo.InvariantInfo,
                            DateTimeStyles.None,
                            out validDateChecked);
                        //if (validDate)
                        //    validDateChecked.ToShortDateString());
                        //else
                        //    Console.WriteLine("Not a valid date: {0}", readAddMeeting);
                        string dt = txtTimePointCalc.Text;
                        validDate = DateTime.TryParseExact(
                           dt.Substring(0,10),
                           dateFormats,
                           DateTimeFormatInfo.InvariantInfo,
                           DateTimeStyles.None,
                           out validDateChecked2);
                        int result = DateTime.Compare(validDateChecked, validDateChecked2);


                        if (result <= 0)
                            dateIsGood = true;
                        else { dateIsGood = false; MessageBox.Show("The <timepoint data> can not be greater than <timepoint calculation>");}
                            

                    }
                    else
                    {
                        dateIsGood = false;
                    }

                    
                }
                else
                {
                    MessageBox.Show("Invalid date format date must be formatted to dd.mm.yyyy");
                    
                    
                    dateIsGood = false;
                    txtTimePointData.Focus();
                }
            }
            else
            {
                MessageBox.Show("Please provide entry date in the format of dd.mm.yyyy");
            
                dateIsGood = false;
                txtTimePointData.Focus();
            }
        }
       
        private void btnSave_Click_1(object sender, EventArgs e) // Saving data in Calculation View
        {
            try {
                txtEntryDate_Validated(sender, e);
            if (dateIsGood) { 

                btnSave_Click(sender, e);
                MessageBox.Show("The data was saved!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void rbAll_CheckedChanged_1(object sender, EventArgs e) // "All" checkbox change
        {
            try { 
            rbAll_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void rbFitStat_CheckedChanged_1(object sender, EventArgs e) // filter table view fit statistically
        {
            try { 
            rbFitStat_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void rbNotFitStat_CheckedChanged_1(object sender, EventArgs e) // filter table view not fit statistically
        {
            try { 
            rbNotFitStat_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

      
        public void SQLRunFill(string sqlQuery, ComboBox cmbInput) // Filling in Combobox with SQL
        {
            try { 
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand command1 = new SqlCommand(sqlQuery, conn);
                command1.CommandTimeout = 600;
                var cmdSelectFromProduct = command1.ExecuteScalar();

                System.Data.DataTable table = new System.Data.DataTable();





                SqlDataReader dr = command1.ExecuteReader();

                cmbInput.Text = "All";


                while (dr.Read())
                {
                    cmbInput.Items.Add(dr[0].ToString());

                }

                dr.Close();
                conn.Close();
            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        public void SQLRunFillChechedListBox(string sqlQuery, CheckedListBox cmbInput) // Filling in CheckedListBox with SQL
        {
            try {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand command1 = new SqlCommand(sqlQuery, conn);
                command1.CommandTimeout = 600;
                var cmdSelectFromProduct = command1.ExecuteScalar();

                System.Data.DataTable table = new System.Data.DataTable();





                SqlDataReader dr = command1.ExecuteReader();

                cmbInput.Text = "All";


                while (dr.Read())
                {
                    cmbInput.Items.Add(dr[0].ToString());

                }

                dr.Close();
                conn.Close();
            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public void SQLRunFillCheckedListBox(string sqlQuery, CheckedListBox lsInput) // Filling in CheckedListBox with SQL
        {
            try { 
            lsInput.Items.Clear();
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand command1 = new SqlCommand(sqlQuery, conn);
                var cmdSelectFromProduct = command1.ExecuteScalar();

                System.Data.DataTable table = new System.Data.DataTable();





                SqlDataReader dr = command1.ExecuteReader();
                //lsInput.Items.Add("All");
                //lsInput.Text = "All";

                int i = 0;

                while (dr.Read())
                {
                    lsInput.Items.Add(dr[0].ToString());
                    lsInput.SetItemChecked(i, true);
                    i++;
                }

                dr.Close();
                conn.Close();
            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public void SQLRunFillListBox(string sqlQuery, System.Windows.Forms.ListBox lsInput) // Filling in ListBox with SQL
        {
            try { 
            lsInput.Items.Clear();
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand command1 = new SqlCommand(sqlQuery, conn);
                var cmdSelectFromProduct = command1.ExecuteScalar();

                System.Data.DataTable table = new System.Data.DataTable();





                SqlDataReader dr = command1.ExecuteReader();
                lsInput.Items.Add("All");
                lsInput.Text = "All";


                while (dr.Read())
                {
                    lsInput.Items.Add(dr[0].ToString());

                }

                dr.Close();
                conn.Close();
            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        private void frmHistResults_Load(object sender, EventArgs e) // Loading of results for Virt_Ozids in Calculation Search form
        {
            try
            {

                chDeActivate.Checked = false;
                chDeActivate.Enabled = false;
                chActivate.Checked = false;
                chActivate.Enabled = false;

                dtCalcDateTime.Enabled = false;
                clbVirtOzid.Enabled = false;
                rbAll1.Enabled = false;
                rbActive1.Enabled = false;
                rbNotActive1.Enabled = false;

                groupFilterSelection.Enabled = false;
                panelSelection.Enabled = false;
                panelButtons.Enabled = false;

                string INIfolderPath;

                //Reading data from app.ini file
                INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                INIfolderPath = INIfolderPath + "\\app.ini";

                string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                connectionString = lines[0];
                strRscript = lines[1];
                strRpath = lines[2];
                strDataDir = lines[3];
                strOutputDir = lines[4];
                strOutPutPath = lines[4];
                string strQuery = "select distinct VIRT_OZID from CalculationRaw order by VIRT_OZID";
                string strQuery2 = "select distinct CalcID from CalculationRaw order by CalcID";
                string strQuery3 = "select distinct PRODUCTCODE from CalculationRaw order by PRODUCTCODE";

                chkVirtOzid2.Items.Clear();
                SQLRunFillChechedListBox(strQuery, chkVirtOzid2);
                cmbProdID2.Items.Clear();
                SQLRunFill(strQuery3, cmbProdID2);
                cmbCalcID2.Items.Clear();
                SQLRunFill(strQuery2, cmbCalcID2);
                lbOzid.Items.Clear();
                SQLRunFillListBox(strQuery, lbOzid);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }



        private void rbFitStat_CheckedChanged1(object sender, EventArgs e) // Pressing FitStatistically radio button
        {
            strFilter += " and FitStatistically ='True'";

        }

        private void rbNotFitStat_CheckedChanged1(object sender, EventArgs e)
        {
            strFilter += " and FitStatistically ='False'";

        }
       
        
        void tb_LostFocus1(object sender, EventArgs e) //Lost focus functional
        {
            try
            {
                System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
                var num = Int16.Parse(tb.Name.Substring(3, tb.Name.Length - 3));
                strMatrix[num, 10] = tb.Text;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

       

       

       

        private void btnGetHistoric_Click(object sender, EventArgs e) // Get historic data in calculation search
        {
            
            btnZoomIn1.Visible = false;
            btnPrint.Visible = false;
            pictureBox2.Visible = false;
            label82.Visible = false;
            label83.Visible = false;
            strFilter = "";
            
            if (rbActive1.Checked == true)
            {
                strFilter += " and Active = 'True'";
            }
            if (rbNotActive1.Checked == true)
            {
                strFilter += " and Active = 'False'";
            }
            if (rbFitstatF.Checked == true)
            {
                strFilter += " and FitStatistically ='True'";
            }

            if (rbNotFitStatF.Checked == true)
            {
                strFilter += " and FitStatistically ='False'";
            }

            if (rbKPI0.Checked == true)
            {
                strFilter += " and CAST(KPI0 AS real) > 0";

            }
            if (rbKPI1.Checked == true)
            {
                strFilter += " and CAST(KPI1 AS real) > 0";

            }
            if (rbKPI2.Checked == true)
            {
                strFilter += " and CAST(KPI2 AS real) > 0";

            }
            if (rbKPI3.Checked == true)
            {
                strFilter += " and CAST(KPI3 AS real) > 0";

            }


            try
            {
                strQuery = " (1 = 1)     and 0=0";
                var strFilterVirtOzid = " and Virt_Ozid in ('";

                for (int j = 0; j < chkVirtOzid2.Items.Count; ++j)
                {
                    if (chkVirtOzid2.GetItemCheckState(j) == CheckState.Checked)
                    {
                        strFilterVirtOzid += (string)chkVirtOzid2.Items[j] + "','";
                    }
                }
                strFilterVirtOzid = Left(strFilterVirtOzid, strFilterVirtOzid.Length - 2) + ") ";
                if (strFilterVirtOzid == " and Virt_Ozid in ) ")
                    strFilterVirtOzid = "";
                    int index = strFilterVirtOzid.IndexOf("()");
                if (index <= 0)
                {
                    strQuery += strFilterVirtOzid;

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            try
            {



                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcRow where CalcID = '" + cmbCalcID2.Text + "' and " + strQuery;
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }






            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            if (!IsCalcIDAvailable)
            {
                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel6.Controls.Clear();



                        string qry = "";

                        qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID2.Text + "') and CalcID = '" + cmbCalcID2.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";


                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Text = dr[1].ToString(),
                                Top = i * 20,
                                Left = 3,


                            };
                            strMatrix[i, 1] = dr[1].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[1].ToString();

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[2].ToString(),
                                Top = i * 20,
                                Left = 110
                            };
                            strMatrix[i, 2] = dr[2].ToString();
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[3].ToString(),
                                Top = i * 20,
                                Left = 190

                            };
                            strMatrix[i, 3] = dr[3].ToString();
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[4].ToString(),
                                Top = i * 20,
                                Left = 270

                            };
                            strMatrix[i, 4] = dr[4].ToString();
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[5].ToString(),
                                Top = i * 20,
                                Left = 360

                            };
                            strMatrix[i, 5] = dr[5].ToString();
                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = dr[6].ToString(),
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 6] = dr[6].ToString();
                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 580

                            };
                            ch1.MouseLeave += new EventHandler(chk_MouseMove);
                            if (dr[5].ToString() == "100")
                            {
                                ch1.Checked = true;

                            }


                            else
                            {

                            }


                            ch1.Click += new EventHandler(chk_Click);
                            ch1.MouseLeave += new EventHandler(chk_MouseMove);
                            strMatrix[i, 7] = ch1.Checked.ToString();
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 680

                            };
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }
                            this.panel6.Controls.Add(b1);
                            b1.Click += new EventHandler(ba1_Click);


                            //strMatrix[i, 10] 


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 770

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            ch2.MouseLeave += new EventHandler(chk_MouseMove);
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr[12].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 1050

                            };
                            tb7.LostFocus += new EventHandler(tb_LostFocus1);
                            strMatrix[i, 10] = tb7.Text;
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }

                           
                            //Add data to matrix for inserting to table

                            panel6.Controls.Add(tb1); panel6.Controls.Add(tb2); panel6.Controls.Add(tb3); panel6.Controls.Add(tb4); panel6.Controls.Add(tb5); panel6.Controls.Add(tb6);
                            panel6.Controls.Add(ch1); panel6.Controls.Add(ch2); panel6.Controls.Add(tb7); panel6.Controls.Add(ba);

                        }



                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID2.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                        //txtNParameterTotal.Text = y.ToString();
                        //----------------------------
                        //------------------------------
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID2.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();



                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();






                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }

            }
            else
            {

                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel6.Controls.Clear();

                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";

                        //qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID2.Text + "') and CalcID = '" + cmbCalcID2.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";

                        qry = "SELECT [VIRT_OZID],dbo.KPIcount0(calcid,[VIRT_OZID]) as 'c0',dbo.KPIcount1(calcid,[VIRT_OZID]) as 'c1',dbo.KPIcount2(calcid,[VIRT_OZID]) as 'c2',dbo.KPIcount3(calcid,[VIRT_OZID]) as 'c3',[KPI0],[KPI1],[KPI2],[KPI3],[FitStatistically],[RelevantForDiscussion],[GraphID],[Additional_note], Active FROM [VIRT_OZID_per_calculation] where CalcID = '" + cmbCalcID2.Text + "' and " + strQuery + strFilter;
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();

                        System.Data.DataTable table = new System.Data.DataTable();


                        table.Columns.Add("VIRT_OZID", typeof(string));


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        int y = 0;

                        int i = 0;

                        while (dr.Read())
                        {
                            i++;

                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb1" + i.ToString(),
                                Enabled = false,
                                Width = 50,
                                Text = i.ToString(),
                                Top = i * 20,
                                Left = 3


                            };


                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 100,
                                Text = dr[0].ToString(),
                                Top = i * 20,
                                Left = 100
                            };
                            strMatrix[i, 1] = dr[0].ToString();
                            strOZID = tb1.Text;
                            strArr2[i] = dr[0].ToString();




                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 100,
                                Text = dr[1].ToString() + "," + dr[2].ToString() + "," + dr[3].ToString() + "," + dr[4].ToString(),
                                Top = i * 20,
                                Left = 205

                            };
                            strMatrix[i, 2] = dr[2].ToString();


                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 110,
                                Text = Left(dr[5].ToString(), 4) + ", " + Left(dr[6].ToString(), 4) + ", " + Left(dr[7].ToString(), 4) + ", " + Left(dr[8].ToString(), 4),
                                Top = i * 20,
                                Left = 350

                            };
                            strMatrix[i, 3] = Left(dr[5].ToString(), 4) + ", " + Left(dr[6].ToString(), 4) + ", " + Left(dr[7].ToString(), 4) + ", " + Left(dr[8].ToString(), 4);





                            var ch1 = new System.Windows.Forms.CheckBox()
                            {
                                Name = "ch1" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 520

                            };
                            
                            strMatrix[i, 6] = dr[6].ToString();

                            ch1.Click += new EventHandler(chk_Click);
                            ch1.MouseLeave += new EventHandler(chk_MouseMove);
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 600

                            };
                            if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                            { ch1.Checked = true; }
                            else
                            { ch1.Checked = false; }
                            if (Left(dr[5].ToString(), 4) == "100")
                            {
                                ch1.Checked = true;

                            }


                            else
                            {

                            }
                            this.panel6.Controls.Add(b1);
                            b1.Click += new EventHandler(ba1_Click);
                            strMatrix[i, 7] = ch1.Checked.ToString();

                        


                            var ch2 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch2" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 740

                            };
                            ch2.Click += new EventHandler(chk_Click);
                            ch2.MouseLeave += new EventHandler(chk_MouseMove);
                            var ch3 = new System.Windows.Forms.CheckBox()
                            {

                                Name = "ch3" + i.ToString(),
                                Text = string.Format("{0}", "yes"),
                                Top = i * 20,
                                Left = 920

                            };
                            ch3.Click += new EventHandler(chk_Click);
                            ch3.MouseLeave += new EventHandler(chk_MouseMove);
                            if (dr["active"].ToString() == "True")
                            { ch3.Checked = true; }
                            else
                            { ch3.Checked = false; }
                            strMatrix[i, 9] = ch3.Checked.ToString();
                            ch3.Click += new EventHandler(chk_Click);

                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                Name = "tb7" + i.ToString(),
                                Text = dr[12].ToString(),
                                Width = 150,
                                Top = i * 20,
                                Left = 1050

                            };
                            tb7.LostFocus += new EventHandler(tb_LostFocus1);
                            strMatrix[i, 10] = tb7.Text;
                            if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                            { ch2.Checked = true; }
                            else
                            { ch2.Checked = false; }
                            strMatrix[i, 8] = ch2.Checked.ToString();
                            tb7.LostFocus += new EventHandler(tb_LostFocus);
                            //Add data to matrix for inserting to table

                            panel6.Controls.Add(tb1); panel6.Controls.Add(tb2); panel6.Controls.Add(tb3); panel6.Controls.Add(tb4);
                           
                            panel6.Controls.Add(ch1); panel6.Controls.Add(ch2); panel6.Controls.Add(ch3); panel6.Controls.Add(tb7); panel6.Controls.Add(ba);

                        }

                        strRowsCount = i;

                        dr.Close();

                        string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID2.Text + "'";


                        //RunSQL(conn, query, dr);

                        dr = cmd.ExecuteReader();

                        cmd = new SqlCommand(query, conn);

                        //---------NParameterTotal----------------
                        y = 0;
                        while (dr.Read())
                        {
                            y++;

                        }
                   
                        //Calculation Date
                        qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        //N fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal = '0'";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();


                        //N fit statistically percent
                        qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID2.Text + "')";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();



                        //N does not fit statistically
                        qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal != '0'  ";
                        dr.Close();
                        command1 = new SqlCommand(qry, conn);
                        dr = command1.ExecuteReader();







                        dr.Close();
                        conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }
            }
        }

        private void cmbProductCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cmbCalcID.Items.Clear();
                if (cmbProdID2.Text == "All")
                {


                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID2.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + cmbProdID2.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID2.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();

                    }
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rbActive_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and Active ='True'";

        }

        private void rbNotActive_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and Active ='False'";

        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and Active ='True'";

        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and Active ='False'";

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and CAST(KPI0 AS real) > 0";

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and CAST(KPI1 AS real) > 0";

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and CAST(KPI2 AS real) > 0";

        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            strFilter += " and CAST(KPI3 AS real) > 0";

        }

        private void clbVirtOzid_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbOzid.Items.Clear();
            lbOzid.Items.Add(clbVirtOzid.Text);
            //SQLRunFillListBox("", lbOzid);
            if (cmbCalcID2.Text == "All")
            {
                MessageBox.Show("Please select Calculation first!");


            }
            else
            {
                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct * from CalcRow where CalcID = '" + cmbCalcID2.Text + "' and Virt_Ozid = '" + clbVirtOzid.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();


                        SqlDataReader dr = command1.ExecuteReader();

                        SqlCommand cmd = new SqlCommand(qry, conn);


                        while (dr.Read())
                        {
                            IsCalcIDAvailable = true;

                        }


                        dr.Close();
                        conn.Close();
                    }






                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
                if (!IsCalcIDAvailable)
                {
                    try
                    {

                        using (var conn = new SqlConnection(connectionString))
                        {
                            conn.Open();

                            panel6.Controls.Clear();



                            string qry = "";

                            qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID2.Text + "') and CalcID = '" + cmbCalcID2.Text + "' and Virt_Ozid = '" + clbVirtOzid.Text + "'" + " and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";


                            SqlCommand command1 = new SqlCommand(qry, conn);
                            var cmdSelectFromProduct = command1.ExecuteScalar();

                            System.Data.DataTable table = new System.Data.DataTable();


                            table.Columns.Add("VIRT_OZID", typeof(string));


                            SqlDataReader dr = command1.ExecuteReader();

                            SqlCommand cmd = new SqlCommand(qry, conn);


                            int y = 0;

                            int i = 0;

                            while (dr.Read())
                            {
                                i++;

                                var tb1 = new System.Windows.Forms.TextBox()
                                {
                                    Name = "tb1" + i.ToString(),
                                    Enabled = false,
                                    Text = dr[1].ToString(),
                                    Top = i * 20,
                                    Left = 3,


                                };
                                strMatrix[i, 1] = dr[1].ToString();
                                strOZID = tb1.Text;
                                strArr2[i] = dr[1].ToString();

                                var tb2 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 50,
                                    Text = dr[2].ToString(),
                                    Top = i * 20,
                                    Left = 110
                                };
                                strMatrix[i, 2] = dr[2].ToString();
                                var tb3 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 50,
                                    Text = dr[3].ToString(),
                                    Top = i * 20,
                                    Left = 190

                                };
                                strMatrix[i, 3] = dr[3].ToString();
                                var tb4 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 50,
                                    Text = dr[4].ToString(),
                                    Top = i * 20,
                                    Left = 270

                                };
                                strMatrix[i, 4] = dr[4].ToString();
                                var tb5 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 50,
                                    Text = dr[5].ToString(),
                                    Top = i * 20,
                                    Left = 360

                                };
                                strMatrix[i, 5] = dr[5].ToString();
                                var tb6 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 50,
                                    Text = dr[6].ToString(),
                                    Top = i * 20,
                                    Left = 440

                                };
                                strMatrix[i, 6] = dr[6].ToString();
                                var ch1 = new System.Windows.Forms.CheckBox()
                                {
                                    Name = "ch1" + i.ToString(),
                                    Text = string.Format("{0}", "yes"),
                                    Top = i * 20,
                                    Left = 580

                                };
                                ch1.Click += new EventHandler(chk_Click);
                                strMatrix[i, 7] = ch1.Checked.ToString();
                                var b1 = new System.Windows.Forms.Button()
                                {
                                    Name = "b1" + i.ToString(),


                                    Text = string.Format("{0}", "Chart"),
                                    Top = i * 20,
                                    Left = 680

                                };
                                if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                                { ch1.Checked = true; }
                                else
                                { ch1.Checked = false; }
                                this.panel6.Controls.Add(b1);
                                b1.Click += new EventHandler(ba_Click);


                                //strMatrix[i, 10] 


                                var ch2 = new System.Windows.Forms.CheckBox()
                                {

                                    Name = "ch2" + i.ToString(),
                                    Text = string.Format("{0}", "yes"),
                                    Top = i * 20,
                                    Left = 770

                                };
                                ch2.Click += new EventHandler(chk_Click);
                                strMatrix[i, 8] = ch2.Checked.ToString();
                                var tb7 = new System.Windows.Forms.TextBox()
                                {
                                    Name = "tb7" + i.ToString(),
                                    Text = dr[12].ToString(),
                                    Width = 150,
                                    Top = i * 20,
                                    Left = 1050

                                };
                                tb7.LostFocus += new EventHandler(tb_LostFocus);
                                strMatrix[i, 10] = tb7.Text;
                                if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                                { ch2.Checked = true; }
                                else
                                { ch2.Checked = false; }

                                tb7.LostFocus += new EventHandler(tb_LostFocus);
                                //Add data to matrix for inserting to table

                                panel6.Controls.Add(tb1); panel6.Controls.Add(tb2); panel6.Controls.Add(tb3); panel6.Controls.Add(tb4); panel6.Controls.Add(tb5); panel6.Controls.Add(tb6);
                                panel6.Controls.Add(ch1); panel6.Controls.Add(ch2); panel6.Controls.Add(tb7); panel6.Controls.Add(ba);

                            }



                            dr.Close();

                            string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID2.Text + "'";




                            dr = cmd.ExecuteReader();

                            cmd = new SqlCommand(query, conn);

                            //---------NParameterTotal----------------
                            y = 0;
                            while (dr.Read())
                            {
                                y++;

                            }

                            //Calculation Date
                            qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();


                            //N fit statistically
                            qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal = '0'";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();


                            //N fit statistically percent
                            qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID2.Text + "')";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();


                            //N does not fit statistically
                            qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal != '0'  ";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();




                            dr.Close();
                            conn.Close();
                        }
                    }
                    catch (Exception e1)
                    {
                        // Extract some information from this exception, and then
                        // throw it to the parent method.
                        if (e1.Source != null)
                            MessageBox.Show("IOException source: {0}", e1.Message);
                        //throw;
                    }

                }
                else
                {

                    try
                    {

                        using (var conn = new SqlConnection(connectionString))
                        {
                            conn.Open();

                            panel6.Controls.Clear();

                            //SqlDataReader dr1 = command1.ExecuteReader();
                            string qry = "";

                            //qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID2.Text + "') and CalcID = '" + cmbCalcID2.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID2.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";

                            qry = "SELECT [VIRT_OZID],dbo.KPIcount0(calcid,[VIRT_OZID]) as 'c0',dbo.KPIcount1(calcid,[VIRT_OZID]) as 'c1',dbo.KPIcount2(calcid,[VIRT_OZID]) as 'c2',dbo.KPIcount3(calcid,[VIRT_OZID]) as 'c3',[KPI0],[KPI1],[KPI2],[KPI3],[FitStatistically],[RelevantForDiscussion],[GraphID],[Additional_note], Active FROM[VIRT_OZID_per_calculation] where CalcID = '" + cmbCalcID2.Text + "' and Virt_Ozid = '" + clbVirtOzid.Text + "'";
                            SqlCommand command1 = new SqlCommand(qry, conn);
                            var cmdSelectFromProduct = command1.ExecuteScalar();

                            System.Data.DataTable table = new System.Data.DataTable();


                            table.Columns.Add("VIRT_OZID", typeof(string));


                            SqlDataReader dr = command1.ExecuteReader();

                            SqlCommand cmd = new SqlCommand(qry, conn);


                            int y = 0;

                            int i = 0;

                            while (dr.Read())
                            {
                                i++;

                                var tb1 = new System.Windows.Forms.TextBox()
                                {
                                    Name = "tb1" + i.ToString(),
                                    Enabled = false,
                                    Width = 50,
                                    Text = i.ToString(),
                                    Top = i * 20,
                                    Left = 3


                                };


                                var tb2 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 100,
                                    Text = dr[0].ToString(),
                                    Top = i * 20,
                                    Left = 100
                                };
                                strMatrix[i, 1] = dr[0].ToString();
                                strOZID = tb1.Text;
                                strArr2[i] = dr[0].ToString();




                                var tb3 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 100,
                                    Text = dr[1].ToString() + "," + dr[2].ToString() + "," + dr[3].ToString() + "," + dr[4].ToString(),
                                    Top = i * 20,
                                    Left = 205

                                };
                                strMatrix[i, 2] = dr[2].ToString();


                                var tb4 = new System.Windows.Forms.TextBox()
                                {
                                    Enabled = false,
                                    Width = 110,
                                    Text = Left(dr[5].ToString(), 4) + ", " + Left(dr[6].ToString(), 4) + ", " + Left(dr[7].ToString(), 4) + ", " + Left(dr[8].ToString(), 4),
                                    Top = i * 20,
                                    Left = 350

                                };
                                strMatrix[i, 3] = Left(dr[5].ToString(), 4) + ", " + Left(dr[6].ToString(), 4) + ", " + Left(dr[7].ToString(), 4) + ", " + Left(dr[8].ToString(), 4);


                              


                                var ch1 = new System.Windows.Forms.CheckBox()
                                {
                                    Name = "ch1" + i.ToString(),
                                    Text = string.Format("{0}", "yes"),
                                    Top = i * 20,
                                    Left = 520

                                };
                                strMatrix[i, 6] = dr[6].ToString();

                                ch1.Click += new EventHandler(chk_Click);

                                var b1 = new System.Windows.Forms.Button()
                                {
                                    Name = "b1" + i.ToString(),


                                    Text = string.Format("{0}", "Chart"),
                                    Top = i * 20,
                                    Left = 600

                                };
                                if ((dr["FitStatistically"].ToString() == "true") || (dr["FitStatistically"].ToString() == "True"))
                                { ch1.Checked = true; }
                                else
                                { ch1.Checked = false; }
                                this.panel6.Controls.Add(b1);
                                b1.Click += new EventHandler(ba_Click);
                                strMatrix[i, 7] = ch1.Checked.ToString();

                             


                                var ch2 = new System.Windows.Forms.CheckBox()
                                {

                                    Name = "ch2" + i.ToString(),
                                    Text = string.Format("{0}", "yes"),
                                    Top = i * 20,
                                    Left = 740

                                };
                                ch2.Click += new EventHandler(chk_Click);

                                var ch3 = new System.Windows.Forms.CheckBox()
                                {

                                    Name = "ch3" + i.ToString(),
                                    Text = string.Format("{0}", "yes"),
                                    Top = i * 20,
                                    Left = 920

                                };
                                ch3.Click += new EventHandler(chk_Click);
                                if (dr["active"].ToString() == "True")
                                { ch3.Checked = true; }
                                else
                                { ch3.Checked = false; }
                                strMatrix[i, 9] = ch3.Checked.ToString();

                                ch3.Click += new EventHandler(chk_Click);

                                var tb7 = new System.Windows.Forms.TextBox()
                                {
                                    Name = "tb7" + i.ToString(),
                                    Text = dr[12].ToString(),
                                    Width = 150,
                                    Top = i * 20,
                                    Left = 1050

                                };
                                tb7.LostFocus += new EventHandler(tb_LostFocus);
                                strMatrix[i, 10] = tb7.Text;
                                if ((dr["RelevantForDiscussion"].ToString() == "true") || (dr["RelevantForDiscussion"].ToString() == "True"))
                                { ch2.Checked = true; }
                                else
                                { ch2.Checked = false; }
                                strMatrix[i, 8] = ch2.Checked.ToString();
                                tb7.LostFocus += new EventHandler(tb_LostFocus);
                                //Add data to matrix for inserting to table

                                panel6.Controls.Add(tb1); panel6.Controls.Add(tb2); panel6.Controls.Add(tb3); panel6.Controls.Add(tb4);
                                //panel6.Controls.Add(tb5); panel6.Controls.Add(tb6);
                                panel6.Controls.Add(ch1); panel6.Controls.Add(ch2); panel6.Controls.Add(ch3); panel6.Controls.Add(tb7); panel6.Controls.Add(ba);

                            }



                            dr.Close();

                            string query = "select DISTINCT count(VIRT_OZID)  FROM [dbo].[CalcRow] where CalcID = '" + cmbCalcID2.Text + "'";


                        

                            dr = cmd.ExecuteReader();

                            cmd = new SqlCommand(query, conn);

                            //---------NParameterTotal----------------
                            y = 0;
                            while (dr.Read())
                            {
                                y++;

                            }
                          
                            //Calculation Date
                            qry = "select distinct CalcDate  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();


                            //N fit statistically
                            qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal = '0'";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();


                            //N fit statistically percent
                            qry = "select dbo.PerStatisticallyFit  ('" + cmbCalcID2.Text + "')";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();



                            //N does not fit statistically
                            qry = "select distinct count(signal)  from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "' and signal != '0'  ";
                            dr.Close();
                            command1 = new SqlCommand(qry, conn);
                            dr = command1.ExecuteReader();



                            dr.Close();
                            conn.Close();
                        }
                    }
                    catch (Exception e1)
                    {
                        // Extract some information from this exception, and then
                        // throw it to the parent method.
                        if (e1.Source != null)
                            MessageBox.Show("IOException source: {0}", e1.Message);
                        //throw;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string INIfolderPath;

                //Reading data from app.ini file
                INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                INIfolderPath = INIfolderPath + "\\app.ini";

                string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                connectionString = lines[0];


                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID2.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }

                //if (!IsCalcIDAvailable)
                {

                    SqlCommand command = null;
                    using (SqlConnection connection = new SqlConnection(
                      connectionString))
                    {
                        connection.Open();







                        {
                            command = null;

                            {


                                CalcID = cmbCalcID2.Text;

                                Note = txtNote.Text;
                                if (chActivate.Checked)
                                {
                                    Active = 1;
                                }
                                if (chDeActivate.Checked)
                                {
                                    Active = 0;
                                }



                                var qry = "update CalcResultView ";



                                qry += " set Active  ='" + @Active + "'";

                                command = new SqlCommand(qry, connection);


                                command.Parameters.Add("@Active",
                                     SqlDbType.NVarChar, 50).Value = @Active;

                                command.ExecuteNonQuery();


                         
                                if (strRowsCount > 0)
                                {
                                    var sActive = "";
                                    for (int i = 1; i <= strRowsCount; i++)
                                    {

                                        if (rbAll.Checked == true)
                             
                                        {

                                        }
                                        OZID = strMatrix[i, 1];
                                        CalcID = cmbCalcID2.Text;
                                        TotalN = strMatrix[i, 2];
                                        KPI0 = strMatrix[i, 3];
                                        KPI1 = strMatrix[i, 4];
                                        KPI2 = strMatrix[i, 5];
                                        KPI3 = strMatrix[i, 6];
                                        //Active = strMatrix[i, 9];


                                        FitStatistically = strMatrix[i, 7];

                                        RelevantForDiscussion = strMatrix[i, 8];
                                        sActive = strMatrix[i, 9];
                                        Additional_note = strMatrix[i, 10];
                                        qry = "update VIRT_OZID_per_calculation ";


                                        qry += " set FitStatistically='" + @FitStatistically + "',";
                                        qry += " RelevantForDiscussion='" + @RelevantForDiscussion + "',";
                                        qry += " Additional_note='" + @Additional_note + "',";
                                        qry += " Active ='" + @sActive + "' where VIRT_OZID='" + OZID + "' and CalcID='" + cmbCalcID2.Text + "'";

                                        if (strMatrix[i, 1] != null)
                                        {
                                            command = new SqlCommand(qry, connection);

                                            command.Parameters.Add("@VIRT_OZID",
                                                SqlDbType.NVarChar, 250).Value = @OZID;
                                            command.Parameters.Add("@CalcID",
                                                SqlDbType.NVarChar, 50).Value = @CalcID;

                                            command.Parameters.Add("@Additional_note",
                                                 SqlDbType.NVarChar, 250).Value = @Additional_note;
                                            command.Parameters.Add("@FitStatistically",
                                                 SqlDbType.NVarChar, 50).Value = @FitStatistically;
                                            command.Parameters.Add("@RelevantForDiscussion",
                                                 SqlDbType.NVarChar, 250).Value = @RelevantForDiscussion;
                                            command.Parameters.Add("@Active",
                                               SqlDbType.NVarChar, 250).Value = @sActive;
                                            command.ExecuteNonQuery();
                                        }
                                    }
                                    connection.Close();
                                }
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void chActivate_Click(object sender, EventArgs e)
        {

            chDeActivate.Enabled = true;
            chDeActivate.Checked = false;

            chActivate.Checked = true;
            chActivate.Enabled = true;

        }

        private void chDeActivate_Click(object sender, EventArgs e)
        {

            


                chActivate.Enabled = true;
                chActivate.Checked = false;
                
                chDeActivate.Checked = true;
                chDeActivate.Enabled = true;
                

            
        }

        private void chDeActivate_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnGetHistoric_Click_1(object sender, EventArgs e)
        {
            try { 
            btnGetHistoric_Click(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            radioButton3_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void cmbProdID2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cmbCalcID2.Items.Clear();
                if (cmbProdID2.Text == "All")
                {


                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID2.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + cmbProdID2.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID2.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private void cmbCalcID2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                btnZoomIn1.Visible = false;
                pictureBox2.Visible = false;
                btnPrint.Visible = false;
                label82.Visible = false;
                label83.Visible = false;
                //chDeActivate.Checked = false;
                //chDeActivate.Enabled = true;
                //chActivate.Checked = false;
                //chActivate.Enabled = true;

                dtCalcDateTime.Enabled = true;
                clbVirtOzid.Enabled = true;
                rbAll1.Enabled = true;
                rbActive1.Enabled = true;
                rbNotActive1.Enabled = true;

                groupFilterSelection.Enabled = true;
                panelSelection.Enabled = true;
                panelButtons.Enabled = true;
                





                var ActiveCalc = false;
                //Define if the Calculation can be activated or deactivated
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand("select distinct active from CalcResultView where calcid='" + cmbCalcID2.Text + "'", conn);
                    command1.CommandTimeout = 600;
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();





                    SqlDataReader dr = command1.ExecuteReader();


                    var statusActiv = "0";
                    
                    while (dr.Read())
                    {
                        statusActiv = dr[0].ToString();
                    }
                    if (statusActiv == "False")
                    {
                        chDeActivate.Checked = true;
                        chDeActivate.Enabled = false;
                        chActivate.Checked = false;
                        chActivate.Enabled = true;

                    }
                    else
                    {
                        chActivate.Checked = true;
                        chActivate.Enabled = false;
                        chDeActivate.Checked = false;
                        chDeActivate.Enabled = true;

                    }
                    dr.Close();
                    conn.Close();
                }

                //---------------------------------------------------------
                string strQuery = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID2.Text + "'";
                
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID2.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        //dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());
                        //if (dr["Active"] == "0")
                        //{
                        //    chActivate.Enabled = true;
                        //    chActivate.Checked = false;
                        //    chDeActivate.Enabled = false;
                        //    chDeActivate.Checked = true;
                        //}   
                        //else
                        //{
                        //    chActivate.Enabled = false;
                        //    chActivate.Checked = true;
                        //    chDeActivate.Enabled = true;
                        //    chDeActivate.Checked = false;
                        //}

                        txtNote1.Text = dr["Note"].ToString();
                    }


                    dr.Close();
                    conn.Close();
                }




                //---------------------------------------------------------
                strQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + cmbCalcID2.Text + "' order by VIRT_OZID";
                chkVirtOzid2.Items.Clear();
                lbOzid.Items.Clear();
                SQLRunFillCheckedListBox(strQuery, chkVirtOzid2);
                SQLRunFillListBox(strQuery, lbOzid);
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalculationRaw where CalcID = '" + cmbCalcID2.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());

                    }


                    dr.Close();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            btnGetHistoric_Click(sender, e);
        }

        private void rbKPI0_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkAllRefCpv_CheckedChanged_1(object sender, EventArgs e)
        {
            try
            {
                
                //chkAllVirtOzid.Checked = false;
                chkAllRefCpv_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void chkAllVirtOzid_CheckedChanged_1(object sender, EventArgs e)
        {

            try
            {
                //chkAllRefCpv.Checked = false;
                chkAllVirtOzid_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void SearchCalculation_onload(object sender, EventArgs e)
        {
            try { 
            tabMain.SelectTab("SearchCalculation");
            frmHistResults_Load(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void chkSortDate_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            chkSortDate_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void chLastN_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            chLastN_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void chLastM_CheckedChanged_1(object sender, EventArgs e)
        {
            chLastM_CheckedChanged(sender, e);
        }

        private void chkLaufNr_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            chkLaufNr_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
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
            String cmdText = queury;
            //DataSet ds;  
            //ds = new DataSet();
            string INIfolderPath;
         
            // Current directory
            INIfolderPath = System.IO.Directory.GetCurrentDirectory();
            // Input data file app.ini
            INIfolderPath = INIfolderPath + "\\app.ini";

            string[] lines = System.IO.File.ReadAllLines(INIfolderPath);

            connectionString = lines[0];// Connection String            
            var strRscript = lines[1];// Second line of app.ini file          
            var strRpath = lines[2]; // Third line of app.ini file            
            var strDataDir = lines[3];// Forth line of app.ini file          
            var strOutputDir = lines[4];// Fifth line of app.ini file
            strOutPutPath = lines[4];

            SqlConnection oConn = new SqlConnection();
            oConn.ConnectionString = @connectionString;  // get connection string
            SqlCommand cmd = new SqlCommand(cmdText, oConn);
            
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            try
            {
                oConn.Open();
                da.Fill(ds);       // retrive data
                oConn.Close();
                //Return ds;
            }
            catch (Exception ex1)
            {


                MessageBox.Show(ex1.Message);
                
            }
            DataTable dt;
            if (ds.Tables.Count > 0)
            {

                string sel = cmdText;
                SqlDataAdapter da1 = new SqlDataAdapter(sel, oConn);
                
                da.Fill(ds, "sql");
                dt = ds.Tables["sql"];



               
                dataGridViewRaw.DataSource = dt;   // fill DataGridView
                oConn.Close();
            }
        }


        private void cmbProductCode_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try {

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "SELECT * FROM Products where ProduktCode = '" + cmbProductCode.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    

                    dr.Close();
                    conn.Close();

                    for (int i = 0; i <= intCounterDataGridViews-1; i++)
                    {
                        arrDataGridView[i].Visible = false;

                    }
                   
                    arrDataGridView[intIndex].Visible = false;
                    dataGridViewRaw.Visible = true;
                    var strQuery = "select * from products where PRODUKTCODE = '" +cmbProductCode.Text+ "'";
                    FillDGV(strQuery);

                    //dataGridViewRaw
                }

                 strQuery = "select distinct VIRT_OZID from Products where PRODUKTCODE = '" + cmbProductCode.Text + "'" + " order by VIRT_OZID";
                string strQuery2 = "select distinct REFERENCED_CPV from Products where PRODUKTCODE = '" + cmbProductCode.Text + "'" + " order by REFERENCED_CPV";
                string strQuery3 = "";
                //cmbProductCode.Items.Clear();
                if (cmbProductCode.Text == "")
                {
                    strQuery3 = "select distinct PRODUKTCODE from Products ";

                }
                else
                {
                    strQuery3 = "select distinct PRODUKTCODE from Products where PRODUKTCODE!='" + cmbProductCode.Text + "'";
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("VIRT_OZID", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();


                    clbVirtOzid.Items.Clear();
                    while (dr.Read())
                    {
                        clbVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                        lstCheckExclVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                    }

                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand(strQuery2, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();
                    command1.CommandTimeout = 600;
                    System.Data.DataTable table = new System.Data.DataTable();


                    table.Columns.Add("REFERENCED_CPV", typeof(string));


                    SqlDataReader dr = command1.ExecuteReader();
                    //cmbRefCpv.Items.Add("All");
                    //cmbRefCpv.Text = "All";
                    chkAllRefCpv.Checked = true;
                    clbRefCpv.Items.Clear();
                    while (dr.Read())
                    {
                        clbRefCpv.Items.Add(dr["REFERENCED_CPV"].ToString());

                    }

                    dr.Close();
                    conn.Close();
                }




                cmbProductCode_SelectedIndexChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void lbOzid_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try { 
            var strQuery = "SELECT [GraphName],[VIRT_OZID],[ImageValue],[CalcID],[ID] FROM [Graphs] where calcid='" + cmbCalcID2.Text + "' and VIRT_OZID ='" + lbOzid.SelectedItem.ToString() + "'";
            lbGraph.Items.Clear();
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand command1 = new SqlCommand(strQuery, conn);
                var cmdSelectFromProduct = command1.ExecuteScalar();



                SqlDataReader dr = command1.ExecuteReader();

                while (dr.Read())
                {
                    lbGraph.Items.Add(dr["GraphName"].ToString());

                }

                dr.Close();
                conn.Close();

            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void lbGraph_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try
            {
                blnPressed = true;
                btnZoomIn1.Visible = true;
                btnPrint.Visible = true;
                picGraph.Visible = true;
                pictureBox2.Visible = true;
                SqlConnection CN = new SqlConnection(connectionString); 
                CN.Open();

                SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID2.Text + "' and VIRT_OZID='" + lbOzid.SelectedItem.ToString() + "'", CN);
                

                var da = new SqlDataAdapter(cmd2);
                var ds = new DataSet();
                da.Fill(ds, "Graphs");
                int count = ds.Tables["Graphs"].Rows.Count;

                if (count > 0)
                {
                    var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                    var stream = new MemoryStream(data);
                    pictureBox2.Image = Image.FromStream(stream);
                }








                //blnPressed = false;

            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string INIfolderPath;

                //Reading data from app.ini file
                INIfolderPath = System.IO.Directory.GetCurrentDirectory();
                INIfolderPath = INIfolderPath + "\\app.ini";

                string[] lines = System.IO.File.ReadAllLines(INIfolderPath);
                connectionString = lines[0];


                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID2.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable = true;

                    }


                    dr.Close();
                    conn.Close();
                }

                if (!IsCalcIDAvailable)
                {

                    SqlCommand command = null;
                    using (SqlConnection connection = new SqlConnection(
                      connectionString))
                    {
                        connection.Open();




                        NParameterTotal = txtNParameterTotal.Text;
                        NStatistically = txtNStatistically.Text;
                        PercentStatistically = txtPercentStatistically.Text;
                        DoNotFitStatistically = txtDoNotFitStatistically.Text;
                        CalcID = cmbCalcID2.Text;
                        User = txtUser.Text;
                        TimePointData = txtTimePointData.Text;
                        TimePointCalc = txtTimePointCalc.Text;
                        Note = txtNote.Text;
                        if (this.chkActive.Checked)
                        {
                            Active = 1;
                        }
                        else
                        {
                            Active = 0;
                        }



                        string qry = "insert into CalcResultView (ID, NParameterTotal, ";
                        qry += "NStatistically, PercentStatistically, DoNotFitStatistically, CalcID, [User],TimePointData,TimePointCalc, Note,Active) values(";
                        qry += " '" + @ID + "','" + @NParameterTotal + "','" + @NStatistically + "','" + @PercentStatistically + "','" + @DoNotFitStatistically + "','" + @CalcID + "','" + @User;

                        qry += "','" + @TimePointData;
                        qry += "','" + @TimePointCalc;
                        qry += "','" + @Note;
                        qry += "','" + @Active + "')";

                        command = new SqlCommand(qry, connection);

                        command.Parameters.Add("@ID",
                                                        SqlDbType.Int).Value = @ID;
                        command.Parameters.Add("@NParameterTotal",
                            SqlDbType.NVarChar, 50).Value = @NParameterTotal;
                        command.Parameters.Add("@NStatistically",
                             SqlDbType.NVarChar, 50).Value = @NStatistically;
                        command.Parameters.Add("@PercentStatistically",
                             SqlDbType.NVarChar, 50).Value = @PercentStatistically;
                        command.Parameters.Add("@DoNotFitStatistically",
                             SqlDbType.NVarChar, 50).Value = @DoNotFitStatistically;
                        command.Parameters.Add("@CalcID",
                             SqlDbType.NVarChar, 50).Value = @CalcID;
                        command.Parameters.Add("@User",
                             SqlDbType.NVarChar, 50).Value = @User;
                        command.Parameters.Add("@TimePointData",
                             SqlDbType.NVarChar, 50).Value = @TimePointData;
                        command.Parameters.Add("@TimePointCalc",
                             SqlDbType.NVarChar, 50).Value = @TimePointCalc;
                        command.Parameters.Add("@Note",
                             SqlDbType.NVarChar, 250).Value = @Note;
                        command.Parameters.Add("@Active",
                             SqlDbType.NVarChar, 50).Value = @Active;
                        command.ExecuteNonQuery();

                        if (Int16.Parse(txtNParameterTotal.Text) > 0)
                        {
                            for (int i = 1; i <= Int16.Parse(txtNParameterTotal.Text); i++)
                            {
                                OZID = strMatrix[i, 1];
                                CalcID = cmbCalcID2.Text;
                                TotalN = strMatrix[i, 2];
                                KPI0 = strMatrix[i, 3];
                                KPI1 = strMatrix[i, 4];
                                KPI2 = strMatrix[i, 5];
                                KPI3 = strMatrix[i, 6];



                                FitStatistically = strMatrix[i, 7];

                                RelevantForDiscussion = strMatrix[i, 8];
                                Additional_note = strMatrix[i, 9];

                                var sql = "select ID from Graphs where VIRT_OZID ='" + @OZID + "'" + " and CalcID = '" + cmbCalcID2.Text + "'";
                                command = new SqlCommand(sql, connection);
                                SqlDataReader dr = command.ExecuteReader();

                                SqlCommand cmd = new SqlCommand(sql, connection);


                                while (dr.Read())
                                {
                                    @GraphID = dr["ID"].ToString();

                                }


                                dr.Close();
                                //conn.Close();




                                qry = "insert into VIRT_OZID_per_calculation (VIRT_OZID,CalcID,TotalN,KPI0,KPI1,KPI2,KPI3,Additional_note,FitStatistically,RelevantForDiscussion,GraphID) ";
                                qry += " values(";
                                qry += " '" + @OZID + "','" + @CalcID + "','" + @TotalN + "','" + @KPI0 + "','" + @KPI1 + "','" + @KPI2 + "','" + @KPI3;
                                qry += "','" + @Additional_note;
                                qry += "','" + @FitStatistically;
                                qry += "','" + @RelevantForDiscussion;
                                qry += "','" + @GraphID + "')";

                                command = new SqlCommand(qry, connection);

                                command.Parameters.Add("@VIRT_OZID",
                                    SqlDbType.NVarChar, 250).Value = @OZID;
                                command.Parameters.Add("@CalcID",
                                    SqlDbType.NVarChar, 50).Value = @CalcID;
                                command.Parameters.Add("@TotalN",
                                     SqlDbType.NVarChar, 50).Value = @TotalN;
                                command.Parameters.Add("@KPI0",
                                     SqlDbType.NVarChar, 50).Value = @KPI0;
                                command.Parameters.Add("@KPI1",
                                     SqlDbType.NVarChar, 50).Value = @KPI1;
                                command.Parameters.Add("@KPI2",
                                     SqlDbType.NVarChar, 50).Value = @KPI2;
                                command.Parameters.Add("@KPI3",
                                     SqlDbType.NVarChar, 50).Value = @KPI3;
                                command.Parameters.Add("@Additional_note",
                                     SqlDbType.NVarChar, 50).Value = @Additional_note;
                                command.Parameters.Add("@FitStatistically",
                                     SqlDbType.NVarChar, 50).Value = @FitStatistically;
                                command.Parameters.Add("@RelevantForDiscussion",
                                     SqlDbType.NVarChar, 250).Value = @RelevantForDiscussion;
                                command.Parameters.Add("@GraphID",
                                     SqlDbType.NVarChar, 50).Value = @GraphID;
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }




                else
                {
                    //SqlCommand command = null;
                    using (SqlConnection connection = new SqlConnection(
                      connectionString))
                    {
                        connection.Open();
                        
                        Note = txtNote1.Text;
                        if (this.chActivate.Checked)
                        {
                            Active = 1;
                        }
                        else
                        {
                            Active = 0;
                        }

                        //command.Parameters.Add("@ReportsTo",
                        //    SqlDbType.Int).Value = reportsTo;

                        //command.Parameters.Add("@Photo",
                        //    SqlDbType.Image, photo.Length).Value = photo;
                        @CalcID = cmbCalcID2.Text;
                        var qry = "update CalcResultView ";

                     
                        qry += " set Note='" + @Note + "',";
                        qry += " Active  ='" + @Active + "' where calcid='" + @CalcID + "'";

                        SqlCommand command1 = new SqlCommand(qry, connection);

                       
                        command1.Parameters.Add("@Note",
                             SqlDbType.NVarChar, 250).Value = @Note;
                        command1.Parameters.Add("@Active",
                             SqlDbType.NVarChar, 50).Value = @Active;
                        
                        command1.ExecuteNonQuery();






                        SqlCommand command = null;

                        //connection.Open();







                        {
                            command = null;

                            {


                                CalcID = cmbCalcID2.Text;

                                Note = txtNote.Text;
                                if (chActivate.Checked)
                                {
                                    Active = 1;
                                }
                                if (chDeActivate.Checked)
                                {
                                    Active = 0;
                                }



                                qry = "update CalcResultView ";



                                qry += " set Active  ='" + @Active + "'";

                                command = new SqlCommand(qry, connection);


                                command.Parameters.Add("@Active",
                                     SqlDbType.NVarChar, 50).Value = @Active;

                                //command.ExecuteNonQuery();


                                ////000000000000000000000
                                if (strRowsCount > 0)
                                {
                                    var sActive = "";
                                    for (int i = 1; i <= strRowsCount; i++)
                                    {

                                        if (rbAll.Checked == true)
                                        //&& (rbFitStat.Checked != true)) rbNotFitStat.Checked = true; 
                                        {

                                        }
                                        OZID = strMatrix[i, 1];
                                        CalcID = cmbCalcID.Text;
                                        TotalN = strMatrix[i, 2];
                                        KPI0 = strMatrix[i, 3];
                                        KPI1 = strMatrix[i, 4];
                                        KPI2 = strMatrix[i, 5];
                                        KPI3 = strMatrix[i, 6];
                                        sActive = strMatrix[i, 9];


                                        FitStatistically = strMatrix[i, 7];

                                        RelevantForDiscussion = strMatrix[i, 8];
                                        sActive = strMatrix[i, 9];
                                        Additional_note = strMatrix[i, 10];
                                        qry = "update VIRT_OZID_per_calculation ";


                                        qry += " set FitStatistically='" + @FitStatistically + "',";
                                        qry += " RelevantForDiscussion='" + @RelevantForDiscussion + "',";
                                        qry += " Additional_note='" + @Additional_note + "',";
                                        qry += " Active ='" + @sActive + "' where VIRT_OZID='" + OZID + "' and CalcID='" + cmbCalcID2.Text + "'";

                                        if (strMatrix[i, 1] != null)
                                        {
                                            command = new SqlCommand(qry, connection);

                                            command.Parameters.Add("@VIRT_OZID",
                                                SqlDbType.NVarChar, 250).Value = @OZID;
                                            command.Parameters.Add("@CalcID",
                                                SqlDbType.NVarChar, 50).Value = @CalcID;

                                            command.Parameters.Add("@Additional_note",
                                                 SqlDbType.NVarChar, 250).Value = @Additional_note;
                                            command.Parameters.Add("@FitStatistically",
                                                 SqlDbType.NVarChar, 50).Value = @FitStatistically;
                                            command.Parameters.Add("@RelevantForDiscussion",
                                                 SqlDbType.NVarChar, 250).Value = @RelevantForDiscussion;

                                            command.Parameters.Add("@Active",
                                               SqlDbType.NVarChar, 250).Value = @sActive;
                                            command.ExecuteNonQuery();
                                        }
                                    }
                                    connection.Close();
                                }
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("The data was saved!");
        }

        private void clbRefCpv_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

//        if (chkAllVirtOzid.Checked != true)
//            {
//                var strFilterVirtOzid = " where Virt_Ozid in ('";
//                //ljkh
//                for (int j = 0; j<clbVirtOzid.Items.Count; ++j)
//                {
//                    if (clbVirtOzid.GetItemCheckState(j) == CheckState.Checked)
//                    {
//                        strFilterVirtOzid += (string) clbVirtOzid.Items[j] + "','";
//    }
//}
//strFilterVirtOzid = Left(strFilterVirtOzid, strFilterVirtOzid.Length - 2) + ") ";
//int index = strFilterVirtOzid.IndexOf("()");
//                //if (index <= 0)
//                //{
//                //    strQuery += strFilterVirtOzid;

//                //}
//            }

        private void clbVirtOzid_LostFocus(object sender, EventArgs e)
        {
            try {
                clbRefCpv.Items.Clear();
                var qry1 = " where VIRT_OZID in ('"; 
            //if (chkAllRefCpv.Checked != true)
            {
                
                for (int j = 0; j < clbVirtOzid.Items.Count; ++j)
                {
                    if (clbVirtOzid.GetItemCheckState(j) == CheckState.Checked)
                    {
                        qry1 += (string)clbVirtOzid.Items[j] + "','";
                    }
                }
                qry1 = Left(qry1, qry1.Length - 2) + ") ";
                int index = qry1.IndexOf("()");
                //if (index <= 0)
                //{
                //    strFilterRefCpv = "";
                //}

            }
            
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string qry = "";
                qry = "SELECT distinct REFERENCED_CPV FROM [Products] " + qry1;
                



               //SELECT distinct[VIRT_OZID] FROM [Products] where[REFERENCED_CPV] in @REFERENCED_CPV and PRODUKTCODE = @ProductCode
                SqlCommand command1 = new SqlCommand(qry, conn);
                //var cmdSelectFromProduct = command1.ExecuteScalar();


                SqlDataReader dr = command1.ExecuteReader();

                //SqlCommand cmd = new SqlCommand(qry, conn);

                int i = 0;
                clbRefCpv.Items.Clear();
                while (dr.Read())
                { 
                    
                    
                    clbRefCpv.Items.Add(dr["REFERENCED_CPV"].ToString());
                    clbRefCpv.SetItemChecked(i, true);
                    i++;
                }
                dr.Close();
                conn.Close();
            }


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void clbRefCpv_LostFocus(object sender, EventArgs e)
        {
            try {
                clbVirtOzid.Items.Clear();
            var qry1 = " where REFERENCED_CPV in ('";
            //if (chkAllRefCpv.Checked != true)
            {
                    lstCheckExclVirtOzid.Items.Clear();
                for (int j = 0; j < clbRefCpv.Items.Count; ++j)
                {
                    if (clbRefCpv.GetItemCheckState(j) == CheckState.Checked)
                    {
                        qry1 += (string)clbRefCpv.Items[j] + "','";
                            lstCheckExclVirtOzid.Items.Add((string)clbRefCpv.Items[j]);
                        }
                        
                }
                qry1 = Left(qry1, qry1.Length - 2) + ") ";
                int index = qry1.IndexOf("()");
                //if (index <= 0)
                //{
                //    strFilterRefCpv = "";
                //}

            }

            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string qry = "";
                qry = "SELECT distinct VIRT_OZID FROM [Products] " + qry1;




                //SELECT distinct[VIRT_OZID] FROM [Products] where[REFERENCED_CPV] in @REFERENCED_CPV and PRODUKTCODE = @ProductCode
                SqlCommand command1 = new SqlCommand(qry, conn);
                //var cmdSelectFromProduct = command1.ExecuteScalar();


                SqlDataReader dr = command1.ExecuteReader();

                //SqlCommand cmd = new SqlCommand(qry, conn);

                int i = 0;
                clbVirtOzid.Items.Clear();
                while (dr.Read())
                {
                    clbVirtOzid.Items.Add(dr["VIRT_OZID"].ToString());
                    clbVirtOzid.SetItemChecked(i, true);
                    i++;
                }
                dr.Close();
                conn.Close();
            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }



        }
        private void timer1_Tick_1(object sender, EventArgs e)
        {

        }

        private void chActivate_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            if (chActivate.Checked)
            {
                chActivate.Enabled = false;
                chDeActivate.Enabled = true;
                chDeActivate.Checked = false;

            }
            if (chDeActivate.Checked)
            {
                chActivate.Enabled = true;
                chDeActivate.Enabled = false;
                chDeActivate.Checked = true;

            }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void chDeActivate_CheckedChanged_1(object sender, EventArgs e)
        {
            try { 
            chDeActivate_CheckedChanged(sender, e);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
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
            strFilter = "";
            btnPrint3.Visible = false;
            btnPrint4.Visible = false;
            pictureBox3.Image = null;
            pictureBox4.Image = null;
            btnZoomOut3.Visible = false;
            btnZoomOut4.Visible = false;
            
            for (int l = 0; l < 500   ;l++)
            {
                strArrToolTip[l] = "";
            }


            try
            {
                strQuery3 = " (1 = 1)     and 0=0";
                strQuery4 = " (1 = 1)     and 0=0";
                //var strFilterVirtOzid = " and Virt_Ozid in ('";
                var strFilterVirtOzid3 = " and Virt_Ozid in ('";
                var strFilterVirtOzid4 = " and Virt_Ozid in ('";
                for (int j = 0; j < chkVirtOzid3.Items.Count; ++j)
                {
                    if (chkVirtOzid3.GetItemCheckState(j) == CheckState.Checked)
                    {
                        strFilterVirtOzid3 += (string)chkVirtOzid3.Items[j] + "','";
                    }
                }
                strFilterVirtOzid3 = Left(strFilterVirtOzid3, strFilterVirtOzid3.Length - 2) + ") ";
                if (strFilterVirtOzid3 == " and Virt_Ozid in ) ")
                    strFilterVirtOzid3 = "";
                int index = strFilterVirtOzid3.IndexOf("()");
                if (index <= 0)
                {
                    strQuery3 += strFilterVirtOzid3;
                }


                for (int j = 0; j < chkVirtOzid4.Items.Count; ++j)
                {
                    if (chkVirtOzid4.GetItemCheckState(j) == CheckState.Checked)
                    {
                        strFilterVirtOzid4 += (string)chkVirtOzid4.Items[j] + "','";
                    }
                }
                strFilterVirtOzid4 = Left(strFilterVirtOzid4, strFilterVirtOzid4.Length - 2) + ") ";
                if (strFilterVirtOzid4 == " and Virt_Ozid in ) ")
                    strFilterVirtOzid4 = "";
                index = strFilterVirtOzid4.IndexOf("()");
                if (index <= 0)
                {
                    strQuery4 += strFilterVirtOzid4;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            try
            {



                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcRow where CalcID = '" + cmbCalcID3.Text + "'" ;
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable3 = true;

                    }


                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcRow where CalcID = '" + cmbCalcID4.Text + "' ";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        IsCalcIDAvailable4 = true;

                    }


                    dr.Close();
                    conn.Close();
                }




            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            //if (!IsCalcIDAvailable3)
            {
                try
                {

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        panel24.Controls.Clear();



                        //SqlDataReader dr1 = command1.ExecuteReader();
                        string qry = "";
                        
                        string[,] strArr3 = new string[500, 8];
                        string[,] strArr4 = new string[500, 8];
                        string[,] strArr5 = new string[500, 16];

                        //qry = "select DISTINCT * FROM [dbo].[CalcRow] where VIRT_OZID in (select distinct  VIRT_OZID from Graphs where CalcID = '" + cmbCalcID3.Text + "') and CalcID = '" + cmbCalcID3.Text + "' and CalcID in (select CalcID from Graphs where CalcID = '" + cmbCalcID3.Text + "') group by VIRT_OZID, calcid, [Total N],[KPI0],[KPI1],[KPI2],[KPI3],FitStatistically, RelevantForDiscussion, Note";
                        var qry3 = "  Select VIRT_OZID, dbo.KPIcount0(CalcID, VIRT_OZID), dbo.KPIcount1(CalcID, VIRT_OZID), dbo.KPIcount2(CalcID, VIRT_OZID), dbo.KPIcount3(CalcID, VIRT_OZID), FitStatistically, RelevantForDiscussion,  Additional_note from   VIRT_OZID_per_calculation where calcid = '" + cmbCalcID3.Text + "' and " + strQuery3 +""+ " order by VIRT_OZID";
                        var qry4 = "  Select VIRT_OZID, dbo.KPIcount0(CalcID, VIRT_OZID), dbo.KPIcount1(CalcID, VIRT_OZID), dbo.KPIcount2(CalcID, VIRT_OZID), dbo.KPIcount3(CalcID, VIRT_OZID), FitStatistically, RelevantForDiscussion,  Additional_note from   VIRT_OZID_per_calculation where calcid = '" + cmbCalcID4.Text + "' and " + strQuery4 + "" + " order by VIRT_OZID";

                        var qry5 = "select DISTINCT VIRT_OZID FROM [dbo].VIRT_OZID_per_calculation where calcid ='" + cmbCalcID3.Text + "' and " + strQuery3 + "" +   " union select DISTINCT VIRT_OZID FROM [dbo].VIRT_OZID_per_calculation where calcid ='" + cmbCalcID4.Text + "' and " + strQuery4 + "" ;






                        SqlCommand cmd = new SqlCommand(qry3, conn);


                        SqlDataReader dr1 = cmd.ExecuteReader();
                        int k = 0; int k1 = 0; int k2 = 0;
                       try { 
                        while (dr1.Read())
                        {
                            strArr3[k, 0] = dr1[0].ToString();
                            strArr3[k, 1] = dr1[1].ToString();
                            strArr3[k, 2] = dr1[2].ToString();
                            strArr3[k, 3] = dr1[3].ToString();
                            strArr3[k, 4] = dr1[4].ToString();
                            strArr3[k, 5] = dr1[5].ToString();
                            strArr3[k, 6] = dr1[6].ToString();
                            strArr3[k, 7] = dr1[7].ToString();
                           
                            k++;
                        }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        k1 = k;

                        dr1.Close();



                        var  cmd1 = new SqlCommand(qry4, conn);
                        var dr2 = cmd1.ExecuteReader();
                        
                        k = 0;
                        while (dr2.Read())
                        {
                            strArr4[k, 0] = dr2[0].ToString();
                            strArr4[k, 1] = dr2[1].ToString();
                            strArr4[k, 2] = dr2[2].ToString();
                            strArr4[k, 3] = dr2[3].ToString();
                            strArr4[k, 4] = dr2[4].ToString();
                            strArr4[k, 5] = dr2[5].ToString();
                            strArr4[k, 6] = dr2[6].ToString();
                            strArr4[k, 7] = dr2[7].ToString();

                            k++;
                        }
                        k2 = k;
                        int max = k1;


                        dr2.Close();


                        if (k1>= k2) 
                        {
                            max = k1;
                        }
                        else
                        {
                            max = k2;
                        }
                        
                        for (int j =0; j<=k1; j++)
                        {
                            arr1[j] = strArr3[j, 0];
                        }

                        for (int j = 0; j <= k2; j++)
                        {
                            arr2[j] = strArr4[j, 0];
                        }

                        for (int j = 0; j <= k1+k2; j++)
                        {
                            arr3[j] = strArr4[j, 0];
                        }


                      cmd = new SqlCommand(qry5, conn);


                        SqlDataReader dr3 = cmd.ExecuteReader();

                        int l = 0;
                        string[] strValue1 = new string[7];
                        string[] strValue2 = new string[7];
                        while (dr3.Read())
                        {
                            //3 scenario
                            //dr3 is in strArr3 and dr3 is in strArr4
                            if ((FindOzid(dr3[0].ToString(), arr1) == 1) && (FindOzid(dr3[0].ToString(), arr2) == 1))
                            {
                                strValue1 = FindArrValue(dr3[0].ToString(), arr1, strArr3, k1);
                                strValue2 = FindArrValue(dr3[0].ToString(), arr2, strArr4, k2);
                                strArr5[l, 0] = strValue1[0];
                                strArr5[l, 1] = strValue1[1];
                                strArr5[l, 2] = strValue1[2];
                                strArr5[l, 3] = strValue1[3];
                                strArr5[l, 4] = strValue1[4];
                                strArr5[l, 5] = strValue1[5];
                                strArr5[l, 6] = strValue1[6];
                                strArr5[l, 7] = strValue1[7];
                                strArr5[l, 8] = strValue2[0];
                                strArr5[l, 9] = strValue2[1];
                                strArr5[l, 10] = strValue2[2];
                                strArr5[l, 11] = strValue2[3];
                                strArr5[l, 12] = strValue2[4];
                                strArr5[l, 13] = strValue2[5];
                                strArr5[l, 14] = strValue2[6];
                                strArr5[l, 15] = strValue2[7];
                            }
                            //MessageBox.Show(dr3[0].ToString());
                            //if dr3[0].ToString() == strArr3[l, 0] 

                            //dr3 is in strArr3 and dr3 is not in strArr4
                            if ((FindOzid(dr3[0].ToString(), arr1) == 1) && (FindOzid(dr3[0].ToString(), arr2) == 0))
                            {
                                strValue1 = FindArrValue(dr3[0].ToString(), arr1, strArr3, k1);
                                //strValue2 = FindArrValue(dr3[0].ToString(), arr2, strArr4, k2);
                                strArr5[l, 0] = strValue1[0];
                                strArr5[l, 1] = strValue1[1];
                                strArr5[l, 2] = strValue1[2];
                                strArr5[l, 3] = strValue1[3];
                                strArr5[l, 4] = strValue1[4];
                                strArr5[l, 5] = strValue1[5];
                                strArr5[l, 6] = strValue1[6];
                                strArr5[l, 7] = strValue1[7];
                                strArr5[l, 8] = "----";
                                strArr5[l, 9] = "----";
                                strArr5[l, 10] = "----";
                                strArr5[l, 11] = "----";
                                strArr5[l, 12] = "----";
                                strArr5[l, 13] = "----";
                                strArr5[l, 14] = "----";
                                strArr5[l, 15] = "----";
                            }
                            //dr3 is not in strArr3 and dr3 is in strArr4
                            if ((FindOzid(dr3[0].ToString(), arr1) == 0) && (FindOzid(dr3[0].ToString(), arr2) == 1))
                            {
                                //strValue1 = FindArrValue(dr3[0].ToString(), arr1, strArr3, k1);
                                strValue2 = FindArrValue(dr3[0].ToString(), arr2, strArr4, k2);
                                strArr5[l, 0] = strValue2[0];
                                strArr5[l, 1] = "----";
                                strArr5[l, 2] = "----";
                                strArr5[l, 3] = "----";
                                strArr5[l, 4] = "----";
                                strArr5[l, 5] = "----";
                                strArr5[l, 6] = "----";
                                strArr5[l, 7] = "----";
                                strArr5[l, 8] = strValue2[0];
                                strArr5[l, 9] = strValue2[1];
                                strArr5[l, 10] = strValue2[2];
                                strArr5[l, 11] = strValue2[3];
                                strArr5[l, 12] = strValue2[4];
                                strArr5[l, 13] = strValue2[5];
                                strArr5[l, 14] = strValue2[6];
                                strArr5[l, 15] = strValue2[7];
                            }
                            l++;
                        }


                        


                        int y = 0;

                        int i = 0;

                        for (i=0;i < max; i++)
                        {
                           

                           if (strArr5[i, 0] != null) 
                            { 
                            var tb1 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 120,
                                Text = strArr5[i, 0],
                                Top = i * 20,
                                Left = 3
                            };
                            strMatrix[i, 1] = tb1.Text;
                            strOZID = tb1.Text;
                            strArr2[i] = tb1.Text;

                            var tb2 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 90,
                                Text = strArr5[i, 1] + ","+ strArr5[i, 2] + "," + strArr5[i, 3] + ","  + strArr5[i, 4] ,
                                Top = i * 20,
                                Left = 150 
                            };
                            strMatrix[i, 2] = tb2.Text;
                            var tb3 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 90,
                                Text = strArr5[i, 9] + "," + strArr5[i, 10] + "," + strArr5[i, 11] + "," + strArr5[i, 12],
                                Top = i * 20,
                                Left = 290

                            };
                            strMatrix[i, 3] = tb3.Text;
                            var tb4 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = strArr5[i, 5],
                                Top = i * 20,
                                Left = 440

                            };
                            strMatrix[i, 4] = tb4.Text;
                            var tb5 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = strArr5[i, 13],
                                Top = i * 20,
                                Left = 529

                            };
                            strMatrix[i, 5] = tb5.Text;
                            if (tb4.Text != tb5.Text)
                            {
                                tb5.BackColor= Color.Red;
                            }    
                            var b1 = new System.Windows.Forms.Button()
                            {
                                Name = "b1" + i.ToString(),


                                Text = string.Format("{0}", "Chart"),
                                Top = i * 20,
                                Left = 619

                            };
                            this.panel24.Controls.Add(b1);
                            b1.Click += new EventHandler(ba34_Click);

                            var tb6 = new System.Windows.Forms.TextBox()
                            {
                                Enabled = false,
                                Width = 50,
                                Text = strArr5[i, 6],
                                Top = i * 20,
                                Left = 719

                            };
                            strMatrix[i, 6] = tb6.Text;
                            
                            var tb7 = new System.Windows.Forms.TextBox()
                            {
                                //Name = "tb7" + i.ToString(),
                                Enabled = false,
                                Text = strArr5[i, 14],
                                Width = 50,
                                Top = i * 20,
                                Left = 790

                            };
                            //tb7.LostFocus += new EventHandler(tb_LostFocus1);
                            strMatrix[i, 7] = tb7.Text;

                            var tb8 = new System.Windows.Forms.TextBox()
                            {
                                //Name = "tb7" + i.ToString(),
                                Enabled = true,
                                Text = strArr5[i, 7],
                                Width = 165,
                                Top = i * 20,
                                Left = 869

                            };
                         
                            //tb7.LostFocus += new EventHandler(tb_LostFocus1);
                            strMatrix[i, 7] = tb7.Text;

                            var tb9 = new System.Windows.Forms.TextBox()
                            {
                                //Name = "tb7" + i.ToString(),
                                Enabled = true,
                                Text = strArr5[i, 15],
                                Width = 165,
                                Top = i * 20,
                                Left = 1079

                            };
                            
                            strMatrix[i, 7] = tb7.Text;
                            



                            //Add data to matrix for inserting to table

                            panel24.Controls.Add(tb1); panel24.Controls.Add(tb2); panel24.Controls.Add(tb3); panel24.Controls.Add(tb4); panel24.Controls.Add(tb5); panel24.Controls.Add(tb6);
                             panel24.Controls.Add(tb7); panel24.Controls.Add(tb8); panel24.Controls.Add(tb9); panel24.Controls.Add(ba);
                            }
                        }



                        
                        //dr2.Close();
                        //conn.Close();
                    }
                }
                catch (Exception e1)
                {
                    // Extract some information from this exception, and then
                    // throw it to the parent method.
                    if (e1.Source != null)
                        MessageBox.Show("IOException source: {0}", e1.Message);
                    //throw;
                }
               
                MessageBox.Show("Comparing is finished!");
            }
            
        }
        public int FindOzid(string Ozid, string[] arrString)
        {
            try { 
            int result = 0;
            for (int i=0; i<= arrString.Length; i++ )
            {
                if ((Ozid == arrString[i]) && (arrString[i]!=""))
                { result = 1;break; }
                if (arrString[i] is null)
                {
                    break;
                }
            }    
            if (result == 0)
            { result = 0; }
            return result;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return 0;
            }
        }
        public string[] FindArrValue(string Ozid, string[] arrString, string [,] strArrMain, int k)
        {
            string[] result = new string[8];
            bool flag = false;
            for (int i = 0; i <= arrString.Length; i++)
            {
                if (Ozid == arrString[i])
                { 
                for (int j = 0; j <= k; j++)
                {
                        if (strArrMain[j, 0] == Ozid)
                        {
                            for (int l = 0; l <= 8; ++l)
                            {
                                if (l == 8) break; 
                                    
                                    
                                result[l] = strArrMain[j, l];


                            }
                            flag = true;
                            break;
                        }
                        if (flag == true)
                        break;
                    }
                    if (flag == true)
                        break;
                }
                if (flag == true)
                    break;
            }
                
            return result;
        }



    private void CompareCalculation_Click(object sender, EventArgs e)
        {

        }

        private void cmbProdID3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cmbCalcID3.Items.Clear();
                if (cmbProdID3.Text == "All")
                {


                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID3.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + cmbProdID3.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID3.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void cmbProdID4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cmbCalcID4.Items.Clear();
                if (cmbProdID4.Text == "All")
                {


                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID4.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();
                    }
                }
                else
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string qry = "";
                        qry = "select distinct CalcID from CalculationRaw where [ProductCode] ='" + cmbProdID4.Text + "'";
                        SqlCommand command1 = new SqlCommand(qry, conn);
                        var cmdSelectFromProduct = command1.ExecuteScalar();
                        SqlDataReader dr = command1.ExecuteReader();
                        SqlCommand cmd = new SqlCommand(qry, conn);
                        while (dr.Read())
                        {
                            cmbCalcID4.Items.Add(dr[0].ToString());
                        }
                        dr.Close();
                        conn.Close();

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void cmbCalcID3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalculationRaw where CalcID = '" + cmbCalcID3.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        dtCalcDateTime3.Value = DateTime.Parse(dr[1].ToString());

                    }


                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID3.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        txtNote3.Text = (dr["Note"].ToString());

                    }


                    dr.Close();
                    conn.Close();
                }




                var ActiveCalc = false;
                //Define if the Calculation can be activated or deactivated
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand("select distinct active from CalcResultView where calcid='" + cmbCalcID3.Text + "'", conn);
                    command1.CommandTimeout = 600;
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();





                    SqlDataReader dr = command1.ExecuteReader();


                    var statusActiv = "0";

                    while (dr.Read())
                    {
                        statusActiv = dr[0].ToString();
                    }
                    if (statusActiv == "False")
                    {
                        chDeActivate.Checked = true;
                        chDeActivate.Enabled = false;
                        chActivate.Checked = false;
                        chActivate.Enabled = true;

                    }
                    else
                    {
                        chActivate.Checked = true;
                        chActivate.Enabled = false;
                        chDeActivate.Checked = false;
                        chDeActivate.Enabled = true;

                    }
                    dr.Close();
                    conn.Close();
                }

                //---------------------------------------------------------
                string strQuery = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID3.Text + "'";

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID3.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        //dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());
                        //if (dr["Active"] == "0")
                        //{
                        //    chActivate.Enabled = true;
                        //    chActivate.Checked = false;
                        //    chDeActivate.Enabled = false;
                        //    chDeActivate.Checked = true;
                        //}   
                        //else
                        //{
                        //    chActivate.Enabled = false;
                        //    chActivate.Checked = true;
                        //    chDeActivate.Enabled = true;
                        //    chDeActivate.Checked = false;
                        //}

                        txtNote1.Text = dr["Note"].ToString();
                    }


                    dr.Close();
                    conn.Close();
                }




                //---------------------------------------------------------
                strQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + cmbCalcID3.Text + "' order by VIRT_OZID";
                chkVirtOzid3.Items.Clear();
                //lbOzid.Items.Clear();
                SQLRunFillCheckedListBox(strQuery, chkVirtOzid3);
                //SQLRunFillListBox(strQuery, lbOzid);
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalculationRaw where CalcID = '" + cmbCalcID3.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    //while (dr.Read())
                    //{
                    //    dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());

                    //}


                    dr.Close();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void cmbCalcID4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalculationRaw where CalcID = '" + cmbCalcID4.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        dtCalcDateTime4.Value = DateTime.Parse(dr[1].ToString());

                    }


                    dr.Close();
                    conn.Close();
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalcResultView where CalcID = '" + cmbCalcID4.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        txtNote4.Text = (dr["Note"].ToString());

                    }


                    dr.Close();
                    conn.Close();
                }







                var ActiveCalc = false;
                //Define if the Calculation can be activated or deactivated
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand command1 = new SqlCommand("select distinct active from CalcResultView where calcid='" + cmbCalcID4.Text + "'", conn);
                    command1.CommandTimeout = 600;
                    var cmdSelectFromProduct = command1.ExecuteScalar();

                    System.Data.DataTable table = new System.Data.DataTable();





                    SqlDataReader dr = command1.ExecuteReader();


                    var statusActiv = "0";

                    while (dr.Read())
                    {
                        statusActiv = dr[0].ToString();
                    }
                    if (statusActiv == "False")
                    {
                        chDeActivate.Checked = true;
                        chDeActivate.Enabled = false;
                        chActivate.Checked = false;
                        chActivate.Enabled = true;

                    }
                    else
                    {
                        chActivate.Checked = true;
                        chActivate.Enabled = false;
                        chDeActivate.Checked = false;
                        chDeActivate.Enabled = true;

                    }
                    dr.Close();
                    conn.Close();
                }

                //---------------------------------------------------------
                string strQuery = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID4.Text + "'";

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select top 1  Note, Active from CalcResultView where calcid='" + cmbCalcID4.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    while (dr.Read())
                    {
                        //dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());
                        //if (dr["Active"] == "0")
                        //{
                        //    chActivate.Enabled = true;
                        //    chActivate.Checked = false;
                        //    chDeActivate.Enabled = false;
                        //    chDeActivate.Checked = true;
                        //}   
                        //else
                        //{
                        //    chActivate.Enabled = false;
                        //    chActivate.Checked = true;
                        //    chDeActivate.Enabled = true;
                        //    chDeActivate.Checked = false;
                        //}

                        txtNote1.Text = dr["Note"].ToString();
                    }


                    dr.Close();
                    conn.Close();
                }




                //---------------------------------------------------------
                strQuery = "select distinct VIRT_OZID from CalculationRaw where calcid='" + cmbCalcID4.Text + "' order by VIRT_OZID";
                chkVirtOzid4.Items.Clear();
                //lbOzid.Items.Clear();
                SQLRunFillCheckedListBox(strQuery, chkVirtOzid4);
                //SQLRunFillListBox(strQuery, lbOzid);
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string qry = "";
                    qry = "select distinct * from CalculationRaw where CalcID = '" + cmbCalcID4.Text + "'";
                    SqlCommand command1 = new SqlCommand(qry, conn);
                    var cmdSelectFromProduct = command1.ExecuteScalar();


                    SqlDataReader dr = command1.ExecuteReader();

                    SqlCommand cmd = new SqlCommand(qry, conn);


                    //while (dr.Read())
                    //{
                    //    dtCalcDateTime.Value = DateTime.Parse(dr[1].ToString());

                    //}


                    dr.Close();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

       
        private Matrix transform = new Matrix();
        public static float s_dScrollValue = 1.01F;

        private void m_Picturebox_Canvas_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Transform = transform;
        }
        private double m_dZoomscale = 1.0;    //THIS IS THE ZOOM SCALE TO WHICH EACH OBJECT 
        private Control ba;

        //ARE ZOOMED IN THE CANVAS  


        //scale factor value for mouse scroll zooming
        //private void ZoomScroll(System.Drawing.Point location, bool zoomIn)
        //{
        //    // make zoom-point (cursor location) our origin
        //    transform.Translate(-location.X, -location.Y);

        //    // perform zoom (at origin)
        //    if (zoomIn)
        //        transform.Scale(s_dScrollValue, s_dScrollValue);
        //    else
        //        transform.Scale(1 / s_dScrollValue, 1 / s_dScrollValue);

        //    // translate origin back to cursor
        //    transform.Translate(location.X, location.Y);

        //    picGraph.Invalidate();
        //}



        private void ZoomScroll(MouseEventArgs e)
        {
            if (e.Delta != 0)
            {
                if (e.Delta <= 0)
                {
                    //set minimum size to zoom
                    if (picGraph.Width < 50)
                        // lbl_Zoom.Text = pictureBox1.Image.Size; 
                        return;
                }
                else
                {
                    //set maximum size to zoom
                    //if (picGraph.Width > 1000)
                    {
                        picGraph.Width += Convert.ToInt32(picGraph.Width * e.Delta / 1000);
                        picGraph.Height += Convert.ToInt32(picGraph.Height * e.Delta / 1000);
                        picGraph.Refresh();
                        return;
                    }
                }    
                        
                
            }
            
            ///*picGraph*/.Invalidate();
        }

        //private void ZoomScroll(System.Drawing.Point location, bool zoomIn)
        //{
        //    // make zoom-point (cursor location) our origin
        //    transform.Translate(-location.X, -location.Y);

        //    // perform zoom (at origin)
        //    if (zoomIn)
        //        transform.Scale(s_dScrollValue, s_dScrollValue);
        //    else
        //        transform.Scale(1 / s_dScrollValue, 1 / s_dScrollValue);

        //    // translate origin back to cursor
        //    transform.Translate(location.X, location.Y);

        //    picGraph.Invalidate();
        //}




        private void btnZoomOut4_Click(object sender, MouseEventArgs ea)
        {
            {
                //Image img = Image.FromFile("C:/Users/User/Desktop/11.png");
                //  flag = 1;
                // Override OnMouseWheel event, for zooming in/out with the scroll wheel
                if (pictureBox4.Image != null)
                {
                    // If the mouse wheel is moved forward (Zoom in)
                    if (ea.Delta > 0)
                    {
                        // Check if the pictureBox dimensions are in range (15 is the minimum and maximum zoom level)
                        if ((pictureBox4.Width < (15 * this.Width)) && (pictureBox4.Height < (15 * this.Height)))
                        {
                            // Change the size of the picturebox, multiply it by the ZOOMFACTOR
                            pictureBox4.Width = (int)(pictureBox4.Width * 1.25);
                            pictureBox4.Height = (int)(pictureBox4.Height * 1.25);

                            // Formula to move the picturebox, to zoom in the point selected by the mouse cursor
                            pictureBox4.Top = (int)(ea.Y - 1.25 * (ea.Y - pictureBox4.Top));
                            pictureBox4.Left = (int)(ea.X - 1.25 * (ea.X - pictureBox4.Left));
                        }
                    }
                    else
                    {
                        // Check if the pictureBox dimensions are in range (15 is the minimum and maximum zoom level)
                        if ((pictureBox4.Width > (100)) && (pictureBox4.Height > (100)))
                        {// Change the size of the picturebox, divide it by the ZOOMFACTOR
                            pictureBox4.Width = (int)(pictureBox4.Width / 1.25);
                            pictureBox4.Height = (int)(pictureBox4.Height / 1.25);

                            // Formula to move the picturebox, to zoom in the point selected by the mouse cursor
                            pictureBox4.Top = (int)(ea.Y - 0.80 * (ea.Y - pictureBox4.Top));
                            pictureBox4.Left = (int)(ea.X - 0.80 * (ea.X - pictureBox4.Left));

                        }
                    }
                }
            }
        }

        private void btnZoomOut4_Click(object sender, EventArgs e)
        {
            try
            {


                SqlConnection CN = new SqlConnection(connectionString);
                //CN.Open();
                //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";


                //System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                ////MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                //var c = btn.Name.Substring(1, btn.Name.Length - 1);
                //int num = Int16.Parse(c);
                //if (num < 20)
                //    //MessageBox.Show(strArr2[num-10]);
                //    strOZID = strArr2[num - 10];
                //if ((num < 200) && (num > 19))


                //    strOZID = strArr2[num - 100];




                //=======================================
                try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID4.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        //var stream = new MemoryStream(data);
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                        PictureBox pb = new PictureBox();
                        pb.Image = Image.FromStream(newImageStream);
                        pb.Location = new System.Drawing.Point(3, 3);
                        pb.Size = new Size(1100, 900);
                        //pictureBox1.Image = Image.FromFile(@"C:\test\1.jpeg");
                        pb.SizeMode = PictureBoxSizeMode.StretchImage;
                        //pb.Width += Convert.ToInt32(pb.Width * 2 );
                        //pb.Height += Convert.ToInt32(pb.Height * 2 );
                        //pb.Refresh();
                        Form frm2 = new Form();
                        frm2.Size = new Size(1200, 1000);
                        frm2.Controls.Add(pb);
                        frm2.ShowDialog();
                        CN.Close();
                    }


                    //picGraph.Visible = true;
                    //CN = new SqlConnection(connectionString);
                    //CN.Open();

                    //cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID4.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    //da = new SqlDataAdapter(cmd2);
                    //ds = new DataSet();
                    //da.Fill(ds, "Graphs");
                    //count = ds.Tables["Graphs"].Rows.Count;

                    //if (count > 0)
                    //{
                    //    var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                    //    //var stream = new MemoryStream(data);
                    //    //pictureBox4.Image = Image.FromStream(stream, true);
                    //    System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                    //    pictureBox4.Image = Image.FromStream(newImageStream, true);
                    //}

                }

                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================



            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void btnZoom1_Click(object sender, EventArgs e)
        {
            try { 
            ba_Zoom(sender, e);
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void btnZoomIn1_Click(object sender, EventArgs e)
        {
            try
            {


                SqlConnection CN = new SqlConnection(connectionString);
                //CN.Open();
                //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";


                //System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                ////MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                //var c = btn.Name.Substring(1, btn.Name.Length - 1);
                //int num = Int16.Parse(c);
                //if (num < 20)
                //    //MessageBox.Show(strArr2[num-10]);
                //    strOZID = strArr2[num - 10];
                //if ((num < 200) && (num > 19))


                //    strOZID = strArr2[num - 100];


                if (blnPressed)
                    strOZID = lbOzid.SelectedItem.ToString();

                //=======================================
                try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID2.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        var stream = new MemoryStream(data);
                        PictureBox pb = new PictureBox();
                        pb.Image = Image.FromStream(stream);
                        pb.Location = new System.Drawing.Point(3, 3);
                        pb.Size = new Size(1100, 900);
                        //pictureBox1.Image = Image.FromFile(@"C:\test\1.jpeg");
                        pb.SizeMode = PictureBoxSizeMode.StretchImage;
                        //pb.Width += Convert.ToInt32(pb.Width * 2 );
                        //pb.Height += Convert.ToInt32(pb.Height * 2 );
                        //pb.Refresh();
                        Form frm = new Form();
                        frm.Size = new Size(1200, 1000);
                        frm.Controls.Add(pb);
                        frm.ShowDialog();
                        CN.Close();
                    }

                    

                }
                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================

                

            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void btnZoomOut3_Click(object sender, EventArgs e)
        {
            try
            {


                SqlConnection CN = new SqlConnection(connectionString);
                //CN.Open();
                //string qry = "insert into Graphs (ID, GraphName, VIRT_OZID, CalcID, ImageValue) values(@ID, @GraphName, @VIRT_OZID, @CalcID, @ImageValue)";


                //System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
                ////MessageBox.Show(btn.Name.Substring(1, btn.Name.Length - 1));
                //var c = btn.Name.Substring(1, btn.Name.Length - 1);
                //int num = Int16.Parse(c);
                //if (num < 20)
                //    //MessageBox.Show(strArr2[num-10]);
                //    strOZID = strArr2[num - 10];
                //if ((num < 200) && (num > 19))


                //    strOZID = strArr2[num - 100];




                //=======================================
                try
                {

                    picGraph.Visible = true;
                    CN = new SqlConnection(connectionString);
                    CN.Open();

                    SqlCommand cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID3.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    var da = new SqlDataAdapter(cmd2);
                    var ds = new DataSet();
                    da.Fill(ds, "Graphs");
                    int count = ds.Tables["Graphs"].Rows.Count;

                    if (count > 0)
                    {
                        var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                        //var stream = new MemoryStream(data);
                        System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                        PictureBox pb = new PictureBox();
                        pb.Image = Image.FromStream(newImageStream);
                        pb.Location = new System.Drawing.Point(3, 3);
                        pb.Size = new Size(1100, 900);
                        //pictureBox1.Image = Image.FromFile(@"C:\test\1.jpeg");
                        pb.SizeMode = PictureBoxSizeMode.StretchImage;
                        //pb.Width += Convert.ToInt32(pb.Width * 2 );
                        //pb.Height += Convert.ToInt32(pb.Height * 2 );
                        //pb.Refresh();
                        Form frm1 = new Form();
                        frm1.Size = new Size(1200, 1000);
                        frm1.Controls.Add(pb);
                        frm1.ShowDialog();
                        CN.Close();
                    }


                    //picGraph.Visible = true;
                    //CN = new SqlConnection(connectionString);
                    //CN.Open();

                    //cmd2 = new SqlCommand("select top 1 imagevalue from Graphs where CalcID='" + cmbCalcID4.Text + "' and VIRT_OZID='" + strOZID + "'", CN);


                    //da = new SqlDataAdapter(cmd2);
                    //ds = new DataSet();
                    //da.Fill(ds, "Graphs");
                    //count = ds.Tables["Graphs"].Rows.Count;

                    //if (count > 0)
                    //{
                    //    var data = (Byte[])ds.Tables["Graphs"].Rows[count - 1]["ImageValue"];
                    //    //var stream = new MemoryStream(data);
                    //    //pictureBox4.Image = Image.FromStream(stream, true);
                    //    System.IO.MemoryStream newImageStream = new System.IO.MemoryStream(data, 0, data.Length);
                    //    pictureBox4.Image = Image.FromStream(newImageStream, true);
                    //}

                }

                catch (Exception ex)
                {


                    MessageBox.Show(ex.ToString());

                }
                //=======================================



            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            
            // Allow the user to choose the page range he or she would
            // like to print.
            printDialog1.AllowSomePages = true;

            // Show the help button.
            printDialog1.ShowHelp = true;

            // Set the Document property to the PrintDocument for 
            // which the PrintPage Event has been handled. To display the
            // dialog, either this property or the PrinterSettings property 
            // must be set 
            printDialog1.Document = docToPrint;

            DialogResult result = printDialog1.ShowDialog();

            // If the result is OK then print the document.
            if (result == DialogResult.OK)
            {
                printDocument1.DefaultPageSettings.Landscape = true;
                printDocument1.Print();
            }

        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            
            e.Graphics.DrawImage(pictureBox2.Image, 0, 0);
        }
        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            e.Graphics.DrawImage(picGraph.Image, 0, 0);
        }
        private void btnPrint3_Click(object sender, EventArgs e)
        {
            printDocument3.PrintPage += new PrintPageEventHandler(printDocument3_PrintPage);
           
            // Allow the user to choose the page range he or she would
            // like to print.
            printDialog1.AllowSomePages = true;

            // Show the help button.
            printDialog1.ShowHelp = true;

            // Set the Document property to the PrintDocument for 
            // which the PrintPage Event has been handled. To display the
            // dialog, either this property or the PrinterSettings property 
            // must be set 
            printDialog1.Document = docToPrint;

            DialogResult result = printDialog1.ShowDialog();

            // If the result is OK then print the document.
            if (result == DialogResult.OK)
            {
                printDocument3.DefaultPageSettings.Landscape = true;
                printDocument3.Print();
            }

        }


        private void printDocument3_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try { 
            e.Graphics.DrawImage(pictureBox3.Image, 0, 0);
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void document_PrintPage(object sender,
    System.Drawing.Printing.PrintPageEventArgs e)
        {

            // Insert code to render the page here.
            // This code will be called when the control is drawn.

            // The following code will render a simple
            // message on the printed document.
            string text = "In document_PrintPage method.";
            System.Drawing.Font printFont = new System.Drawing.Font
                ("Arial", 35, System.Drawing.FontStyle.Regular);

            // Draw the content.
            e.Graphics.DrawString(text, printFont,
                System.Drawing.Brushes.Black, 10, 10);
        }
        private void btnPrint2_Click(object sender, EventArgs e)
        {
            printDocument2.PrintPage += new PrintPageEventHandler(printDocument2_PrintPage);
            
            // Allow the user to choose the page range he or she would
            // like to print.
            printDialog1.AllowSomePages = true;
            //printDialog1.DefaultPageSettings.Landscape = true;
            printDialog1.PrinterSettings.DefaultPageSettings.Landscape = true;
            // Show the help button.
            printDialog1.ShowHelp = true;

            // Set the Document property to the PrintDocument for 
            // which the PrintPage Event has been handled. To display the
            // dialog, either this property or the PrinterSettings property 
            // must be set 
            printDialog1.Document = docToPrint;

            DialogResult result = printDialog1.ShowDialog();
           
            //If the result is OK then print the document.
            if (result == DialogResult.OK)
            {
                printDocument2.DefaultPageSettings.Landscape = true;

                printDocument2.Print();
                //PrintToASpecificPrinter(printDocument2);
            }




        }
        public static void PrintToASpecificPrinter(object docToPrint)
        {
            var fileName = docToPrint;
            using (PrintDialog printDialog = new PrintDialog())
            {
                printDialog.AllowSomePages = true;
                printDialog.AllowSelection = true;
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    var StartInfo = new ProcessStartInfo();
                    StartInfo.CreateNoWindow = true;
                    StartInfo.UseShellExecute = true;
                    StartInfo.Verb = "printTo";
                    StartInfo.Arguments = "\"" + printDialog.PrinterSettings.PrinterName + "\"";
                    StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    StartInfo.FileName = fileName.ToString();

                    Process.Start(StartInfo);
                }

            }


        }
        private void printDocument4_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try { 
            e.Graphics.DrawImage(pictureBox4.Image, 0, 0);
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }
        private void btnPrint4_Click(object sender, EventArgs e)
        {
            printDocument4.PrintPage += new PrintPageEventHandler(printDocument4_PrintPage);
           
            // Allow the user to choose the page range he or she would
            // like to print.
            printDialog1.AllowSomePages = true;

            // Show the help button.
            printDialog1.ShowHelp = true;

            // Set the Document property to the PrintDocument for 
            // which the PrintPage Event has been handled. To display the
            // dialog, either this property or the PrinterSettings property 
            // must be set 
            printDialog1.Document = docToPrint;

            DialogResult result = printDialog1.ShowDialog();

            // If the result is OK then print the document.
            if (result == DialogResult.OK)
            {
                printDocument4.DefaultPageSettings.Landscape = true;
                printDocument4.Print();
            }

        }

        private void btnPrint3_Click_1(object sender, EventArgs e)
        {
            try { 
            btnPrint3_Click(sender, e);
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }
        }

        private void btnPrint4_Click_1(object sender, EventArgs e)
        {
            try { 
                btnPrint4_Click(sender, e);
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.ToString());

            }

        }

        private void dataGridViewRaw_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridViewRaw.ColumnSortModeChanged
        }

        private void Main_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i=0; i< chkVirtOzid3.Items.Count; i++)
            {
                chkVirtOzid3.SetItemChecked(i, false);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkVirtOzid4.Items.Count; i++)
            {
                chkVirtOzid4.SetItemChecked(i, false);
            }
        }

        private void Main_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            tabMain.SelectTab("CompareCalculation");
        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkVirtOzid3.Items.Count; i++)
            {
                chkVirtOzid3.SetItemChecked(i, true);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkVirtOzid4.Items.Count; i++)
            {
                chkVirtOzid4.SetItemChecked(i, true);
            }
        }
    }
}
 
