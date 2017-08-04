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

namespace AMGenkARMPPlan
{
    public partial class Dialog : Form
    {
        private System.Object xx = System.Type.Missing;
        private string xlDirectory;
        private string xlFile;
        private string xlFilePath;

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Worksheet xlWorksheet;

        private ArrayList taskRows = new ArrayList();
        private ArrayList excelColumns = new ArrayList();
        private Hashtable columnLetters = new Hashtable();

        public Dialog()
        {
            InitializeComponent();
            xlDirectory = Properties.Settings.Default.ARMPTasksDirectory;
            xlFile = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectory + xlFile;

            btnImport.Enabled = false;
            if (xlFile != string.Empty)
            {
                GetWorksheet(xlDirectory + xlFile);
            }
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
                    GetWorksheet(xlFilePath);
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
            btnCancel.Enabled = true;
            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            xlWorkbook = xlApp.Workbooks.Open(txtARMPResourcesFile.Text,
                xx, xx, xx, xx, xx, xx, xx,
                xx, xx, xx, xx, xx, xx, xx);

            //TODO: Hardcoded
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Codes"];

            Excel.Range exceptioncodesRange = xlWorksheet.get_Range("B4", "E140");

            Object[,] exceptioncodes = (Object[,])exceptioncodesRange.Cells.Value2;

            Globals.ThisAddIn.CreateARMPExceptionCodes(exceptioncodes);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Basisdata"];

            Excel.Range resourcesRange = xlWorksheet.get_Range("B3", "C11");

            Object[,] resources = (Object[,])resourcesRange.Cells.Value2;

            Globals.ThisAddIn.CreateARMPResources(resources);

            string[] months = { "Augustus"};
            foreach (string month in months)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[month];

                Excel.Range exceptionsRange = xlWorksheet.get_Range("J3", "N11");

                Object[,] exceptions = (Object[,])exceptionsRange.Cells.Value2;

                Globals.ThisAddIn.CreateARMPExceptions(exceptions);
            }
            xlWorkbook.Close();

            xlWorkbook = xlApp.Workbooks.Open(txtARMPTasksFile.Text,
                                              xx, xx, xx, xx, xx, xx, xx,
                                              xx, xx, xx, xx, xx, xx, xx);

            //TODO: Hardcoded
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

            Excel.Range tasksRange = xlWorksheet.get_Range("A1", "W21");

            Object[,] tasks = (Object[,])tasksRange.Cells.Value2;

            Globals.ThisAddIn.CreateARMPTasks(tasks);
            xlWorkbook.Close();
        }

        // Start Excel and open the worksheet with task data.
        private void GetWorksheet(string xlFilePath)
        {
        }

        // Get the range of Excel task data to import.
        private Array ImportTasksFromExcel(string file)
        {
            //TODO: hardcoded 
            // Set the worksheet range of cells to import.
            Excel.Range taskRange = xlWorksheet.get_Range("A2", "Y500");

            Array taskCells = (Array)taskRange.Cells.Value2;
            return taskCells;
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

        // Create the tasks in Project from the imported Excel task data.
        private void btnCreateTasks_Click(object sender, EventArgs e)
        {

            //Globals.ThisAddIn.CreateTasks(customFieldColumn, excelColumns, taskRows);

            HideImportDialogBox();
        }
    }
}
