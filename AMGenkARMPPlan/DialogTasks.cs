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
    public partial class DialogTasks : Form
    {
        private System.Object xx = System.Type.Missing;
        private string xlDirectory;
        private string xlFile;
        //private string xlFilePath;

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Worksheet xlWorksheet;

        public DialogTasks()
        {
            InitializeComponent();
            xlDirectory = Properties.Settings.Default.ARMPTasksDirectory;
            xlFile = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectory + xlFile;

            btnImport.Enabled = false;
            // Show add-in and deployment versions.
            // lblAppVersion.Text = lblAppVersion.Text + this.ProductVersion;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                // This application is installed with ClickOnce.
                string currentVersion =
                    ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                // lblPublishedVersion.Text = lblPublishedVersion.Text + currentVersion;
            }
            else
            {
                // lblPublishedVersion.Text = string.Empty;
            }
        }

        private void btnBrowseT_Click(object sender, EventArgs e)
        {
            string xlsFiles = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";

            xlDirectory = Properties.Settings.Default.ARMPTasksDirectory;
            xlFile = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectory + xlFile;

            this.openFileDialog1.Filter = xlsFiles;
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "Select an Excel File with A-Tasks Data";
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


        // Import the specified Excel task data and show it in a DataGridView.
        private void btnImport_Click(object sender, EventArgs e)
        {
            int lastRowIgnoreFormulas;

            btnCancel.Enabled = true;
            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            xlWorkbook = xlApp.Workbooks.Open(txtARMPTasksFile.Text, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);

            //TODO: Hardcoded
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

            lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range tasksRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsOrig.TaskRows, (int)ARMPExcelLayout.ARMPTasksColsOrig.WorkPlce],
                                                       xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPExcelLayout.ARMPTasksColsOrig.WorkReal]];
            Object[,] tasks = (Object[,])tasksRange.Cells.Value2;
            xlWorkbook.Close();
            Globals.ThisAddIn.CreateUpdateARMPTasks(tasks);
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
                //xlWorkbook.Close(saveChanges, xx, xx);
                xlApp.Quit();
            }
            this.Hide();
        }

    }
}

