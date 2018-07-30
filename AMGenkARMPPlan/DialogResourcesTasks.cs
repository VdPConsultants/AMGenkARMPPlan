using System;
using System.Deployment.Application;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AMGenkARMPPlan
{
    public partial class DialogResourcesTasks : Form
    {
        private System.Object xx = System.Type.Missing;
        private string xlDirectoryT;
        private string xlFileT;

        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Worksheet xlWorksheet;

        private DateTime ARMPStrtDate;
        private DateTime ARMPFnshDate;

        public DialogResourcesTasks()
        {
            InitializeComponent();
            xlDirectoryT = Properties.Settings.Default.ARMPTasksDirectory;
            xlFileT = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectoryT + xlFileT;

            ARMPStrtDate = DateTime.Now;
            SetARMPStrtFnshDate();

            //btnImport.Enabled = false;

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
                lblPublishedVersion.Text = "In debug mode";
            }
        }

        // Set the enumeration for column names in Excel.

        private void btnBrowseT_Click(object sender, EventArgs e)
        {
            string xlsFiles = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";

            xlDirectoryT = Properties.Settings.Default.ARMPTasksDirectory;
            xlFileT = Properties.Settings.Default.ARMPTasksFile;
            txtARMPTasksFile.Text = xlDirectoryT + xlFileT;

            this.openFileDialog1.Filter = xlsFiles;
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "Select an Excel File with Project Data";
            this.openFileDialog1.InitialDirectory = xlDirectoryT;
            this.openFileDialog1.FileName = xlFileT;

            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                try
                {
                    xlFileT = this.openFileDialog1.FileName;
                    int pastLastSlash = xlFileT.LastIndexOf(@"\") + 1;
                    int filenameLength = xlFileT.Length - pastLastSlash;
                    xlDirectoryT = xlFileT.Substring(0, pastLastSlash);
                    xlFileT = xlFileT.Substring(pastLastSlash, filenameLength);

                    string xlFilePath = xlDirectoryT + xlFileT;
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

            Object[,] tasks;

            Globals.ThisAddIn.ARMPStrtDate = ARMPStrtDate;
            Globals.ThisAddIn.ARMPFnshDate = ARMPFnshDate;

            if (cbClipboard.Checked)
            {
                btnCancel.Enabled = true;
                tasks = Globals.ThisAddIn.ImportTasksFromClipboard();
            }
            else
            {
                btnCancel.Enabled = true;
                xlApp = new Excel.Application();
                // Don't interrupt with alert dialogs.
                xlApp.DisplayAlerts = false;

                xlWorkbook = xlApp.Workbooks.Open(txtARMPTasksFile.Text, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);

                //TODO: Hardcoded
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

                lastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                Excel.Range tasksRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsOrig.TaskRows, (int)ARMPPlanExcelLayout.ARMPTasksColsOrig.WorkPlce],
                                                           xlWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPTasksColsOrig.WorkReal]];
                tasks = (Object[,])tasksRange.Cells.Value2;
                xlWorkbook.Close();

                xlApp.Quit();
            }

            Globals.ThisAddIn.CreateARMPPlanWorksheet();

            Globals.ThisAddIn.SetARMPWorkplaces(tasks);

            Globals.ThisAddIn.CreateARMPExceptionCodes(AMGenkResources.GetAMGenkExceptionCodes(Properties.Settings.Default.AMGenkResourcesDirectory, Globals.ThisAddIn.ARMPStrtDate, Globals.ThisAddIn.ARMPFnshDate));
            Globals.ThisAddIn.CreateARMPResources(AMGenkResources.GetAMGenkResources(Properties.Settings.Default.AMGenkResourcesDirectory, Globals.ThisAddIn.ARMPStrtDate, Globals.ThisAddIn.ARMPFnshDate));
            Globals.ThisAddIn.CreateUpdateARMPExceptions(AMGenkResources.GetAMGenkExceptions(Properties.Settings.Default.AMGenkResourcesDirectory, Globals.ThisAddIn.ARMPStrtDate, Globals.ThisAddIn.ARMPFnshDate));

            Globals.ThisAddIn.PrepareARMPTasks();
            Globals.ThisAddIn.CreateUpdateARMPTasks(tasks);
            Globals.ThisAddIn.FormatARMPPlanning();
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            Globals.ThisAddIn.CreateUpdateQRCodesSheet();

            HideImportDialogBox();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            HideImportDialogBox();
        }

        // Store the user settings and close the Import Tasks dialog box.
        private void HideImportDialogBox()
        {
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

        private void cbClipboard_CheckedChanged(object sender, EventArgs e)
        {
            if (cbClipboard.Checked == true)
            {
                txtARMPTasksFile.Enabled = false;
                btnBrowseT.Enabled = false;
                btnImport.Enabled = true;
            }
            else
            {
                txtARMPTasksFile.Enabled = true;
                btnBrowseT.Enabled = true;
                btnImport.Enabled = false;
            }
        }


    }
}


