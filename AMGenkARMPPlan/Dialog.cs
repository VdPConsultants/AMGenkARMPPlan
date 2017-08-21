using System;
using System.Collections.Generic;
using System.Collections;
using System.Deployment.Application;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace AMGenkARMPPlan
{
    public partial class Dialog : Form
    {
        private System.Object xx = System.Type.Missing;
        private string xlDirectory;
        private string xlFile;
        //private string xlFilePath;

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Worksheet xlWorksheet;

        //private ArrayList taskRows = new ArrayList();
        //private ArrayList excelColumns = new ArrayList();
        //private Hashtable columnLetters = new Hashtable();

        private DateTime ARMPStrtDate;
        private DateTime ARMPFnshDate;

        public Dialog()
        {
            InitializeComponent();
            xlDirectory = Properties.Settings.Default.ARMPTasksDirectory;
            xlFile = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectory + xlFile;

            ARMPStrtDate = DateTime.Now.AddDays(7);
            SetARMPStrtFnshDate();

            btnImport.Enabled = false;
            // Show add-in and deployment versions.
            lblAppVersion.Text = lblAppVersion.Text + this.ProductVersion;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                // This application is installed with ClickOnce.
                string currentVersion =
                    ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                lblPublishedVersion.Text = lblPublishedVersion.Text + currentVersion;
            }
            else
            {
                lblPublishedVersion.Text = string.Empty;
            }
        }

        // Set the enumeration for column names in Excel.

        private void btnBrowseR_Click(object sender, EventArgs e)
        {
            string xlsFiles = "Excel files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
            this.openFileDialog1.Filter = xlsFiles;
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "Select the resources import file";
            this.openFileDialog1.InitialDirectory = xlDirectory;
            this.openFileDialog1.FileName = xlFile;

            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                try
                {
                    xlFile = this.openFileDialog1.FileName;
                    int pastLastSlash = xlFile.LastIndexOf(@"\") + 1;
                    int filenameLength = xlFile.Length - pastLastSlash;
                    xlDirectory = xlFile.Substring(0, pastLastSlash);
                    xlFile = xlFile.Substring(pastLastSlash, filenameLength);

                    string xlFilePath = xlDirectory + xlFile;
                    txtARMPResourcesFile.Text = xlFilePath;
                    btnImport.Enabled = true;
                }
                catch (System.Security.SecurityException ex)
                {
                    string err = "Sorry, you don't have the necessary permissions \n";
                    err += "to read the file or directory.\n\n";
                    err += "Error message: " + ex.Message + "\n\n";
                    err += "Details:\n\n" + ex.StackTrace;
                    MessageBox.Show(err, "Security Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnBrowseT_Click(object sender, EventArgs e)
        {
            string xlsFiles = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            this.openFileDialog1.Filter = xlsFiles;
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "Select an Excel File with Project Data";
            this.openFileDialog1.InitialDirectory = xlDirectory;
            this.openFileDialog1.FileName = xlFile;

            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                try
                {
                    xlFile = this.openFileDialog1.FileName;
                    int pastLastSlash = xlFile.LastIndexOf(@"\") + 1;
                    int filenameLength = xlFile.Length - pastLastSlash;
                    xlDirectory = xlFile.Substring(0, pastLastSlash);
                    xlFile = xlFile.Substring(pastLastSlash, filenameLength);

                    string xlFilePath = xlDirectory + xlFile;
                    txtARMPTasksFile.Text = xlFilePath;
                }
                catch (System.Security.SecurityException ex)
                {
                    string err = "Sorry, you don't have the necessary permissions \n";
                    err += "to read the file or directory.\n\n";
                    err += "Error message: " + ex.Message + "\n\n";
                    err += "Details:\n\n" + ex.StackTrace;
                    MessageBox.Show(err, "Security Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        // Import the specified Excel task data and show it in a DataGridView.
        private void btnImport_Click(object sender, EventArgs e)
        {
            int lastRowIgnoreFormulas;
            btnCancel.Enabled = true;
            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            Globals.ThisAddIn.SetARMPStrtFnshDate(ARMPStrtDate, ARMPFnshDate);

            xlWorkbook = xlApp.Workbooks.Open(txtARMPTasksFile.Text,
                                              xx, xx, xx, xx, xx, xx, xx,
                                              xx, xx, xx, xx, xx, xx, xx);

            //TODO: Hardcoded
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

            Excel.Range tasksRange = xlWorksheet.UsedRange;
            Object[,] tasks = (Object[,])tasksRange.Cells.Value2;
            xlWorkbook.Close();
            Globals.ThisAddIn.SetARMPWorkplaces(tasks); 

            xlWorkbook = xlApp.Workbooks.Open(txtARMPResourcesFile.Text,
                xx, xx, xx, xx, xx, xx, xx,
                xx, xx, xx, xx, xx, xx, xx);

            //TODO: Hardcoded

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Codes"];
            lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range exceptioncodesRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionCodesRowsOrig.ClmnHead, (int)ARMPExcelLayout.ARMPExceptionCodesColsOrig.ExcdType],
                                                                xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPExceptionCodesColsOrig.ExcdDayp]];
            Object[,] exceptioncodes = (Object[,])exceptioncodesRange.Cells.Value2;
            Globals.ThisAddIn.CreateARMPExceptionCodes(exceptioncodes);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Basisdata"];
            lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range resourcesRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsOrig.RsrcYear, (int)ARMPExcelLayout.ARMPResourcesColsOrig.WorkPlce],
                                                           xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPResourcesColsOrig.RsrcAmei]];
            Object[,] resources = (Object[,])resourcesRange.Cells.Value2;
            Globals.ThisAddIn.CreateARMPResources(resources);

            DateTime ARMPLastMnth = new DateTime(ARMPStrtDate.Year, ARMPStrtDate.Month, DateTime.DaysInMonth(ARMPStrtDate.Year, ARMPStrtDate.Month));
            Object[,] exceptions_1 = null;
            Object[,] exceptions_2 = null;
            Object[,] exceptions_3 = null;

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[ARMPStrtDate.ToString("MMMM")];
            lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.WorkPlce],
                                                            xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.RsrcName]];
            exceptions_1 = (Object[,])exceptionsRange.Cells.Value2;

            if (ARMPFnshDate.CompareTo(ARMPLastMnth) <= 0)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[ARMPStrtDate.ToString("MMMM")];
                lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPStrtDate.Day],
                                                    xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPFnshDate.Day]];
                exceptions_2 = (Object[,])exceptionsRange.Cells.Value2;
            }
            else
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[ARMPStrtDate.ToString("MMMM")];
                lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPStrtDate.Day - 1],
                                                    xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPLastMnth.Day - 1]];
                exceptions_2 = (Object[,])exceptionsRange.Cells.Value2;
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[ARMPFnshDate.ToString("MMMM")];
                lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt],
                                                    xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPFnshDate.Day - 1]];
                exceptions_3 = (Object[,])exceptionsRange.Cells.Value2;
            }
            Object[,] exceptions = ResourcesMerge(exceptions_1, exceptions_2, exceptions_3);
            Globals.ThisAddIn.CreateARMPExceptions(exceptions);

            Globals.ThisAddIn.CreateARMPTasks(tasks);

            Globals.ThisAddIn.FormatARMPPlanning();
            HideImportDialogBox();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            HideImportDialogBox();
        }

        // Store the user settings and close the Import Tasks dialog box.
        private void HideImportDialogBox()
        {
            Properties.Settings.Default.ARMPTasksDirectory = xlDirectory;
            Properties.Settings.Default.ARMPTasksFile = xlFile;

            if (!(null == xlApp))
            {
                bool saveChanges = false;
                xlWorkbook.Close(saveChanges, xx, xx);
                xlApp.Quit();
            }
            this.Hide();
        }

        private void mcARMPweek_DateChanged(object sender, DateRangeEventArgs e)
        {
            ARMPStrtDate = mcARMPweek.SelectionStart;
            SetARMPStrtFnshDate();
        }

        private void SetARMPStrtFnshDate()
        {
            DayOfWeek firstDayOfWeek = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;

            // TODO: JvdP - last week/first week of a year
            while (ARMPStrtDate.DayOfWeek != firstDayOfWeek)
            {
                ARMPStrtDate = ARMPStrtDate.AddDays(-1);
            }
            ARMPFnshDate = ARMPStrtDate.AddDays(4);
        }

        private Object[,] ResourcesMerge(Object[,] r1, Object[,] r2, Object[,] r3)
        {
            int r1Rows = (r1 == null) ? 0 : r1.GetLength(0);
            int r1Cols = (r1 == null) ? 0 : r1.GetLength(1);
            int r2Cols = (r2 == null) ? 0 : r2.GetLength(1);
            int r3Cols = (r3 == null) ? 0 : r3.GetLength(1);

            int rRows = 1;
            int rCols = 1;

            if (r1 == null)
            {
                return null;
            }

            Object[,] r = NewObjectArray(r1Rows, r1Cols + r2Cols + r3Cols);
            for (int i = 1; i <= r1.GetLength(0); i++)
            {
                rCols = 1;
                for (int j = 1; j <= r1.GetLength(1); j++)
                {
                    r[rRows, rCols] = r1[i, j];
                    rCols++;
                }
                rRows++;
            }
            if (r2 == null)
            {
                return r;
            }
            rRows = 1;
            for (int i = 1; i <= r2.GetLength(0); i++)
            {
                rCols = r1Cols + 1;
                for (int j = 1; j <= r2.GetLength(1); j++)
                {
                    r[rRows, rCols] = r2[i, j];
                    rCols++;
                }
                rRows++;
            }
            if (r3 == null)
            {
                return r;
            }
            rRows = 1;
            for (int i = 1; i <= r3.GetLength(0); i++)
            {
                rCols = r1Cols + r2Cols + 1;
                for (int j = 1; j <= r3.GetLength(1); j++)
                {
                    r[rRows, rCols] = r3[i, j];
                    rCols++;
                }
                rRows++;
            }
            return r;
        }

        /// <summary>
        /// Makes the equivalent of a local Excel range that can be populated 
        ///  without leaving .net
        /// </summary>
        /// <param name="iRows">number of rows in the table</param>
        /// <param name="iCols">number of columns in the table</param>
        /// <returns>a 1's based, 2 dimensional object array which can put back to Excel in one DCOM call.</returns>
        public static object[,] NewObjectArray(int iRows, int iCols)
        {
            int[] aiLowerBounds = new int[] { 1, 1 };
            int[] aiLengths = new int[] { iRows, iCols };

            return (object[,])Array.CreateInstance(typeof(object), aiLengths, aiLowerBounds);
        }
    }
}
