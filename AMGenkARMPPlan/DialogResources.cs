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
    public partial class DialogResources : Form
    {
        public DialogResources()
        {
            InitializeComponent();
        }

         // Import the specified Excel task data and show it in a DataGridView.
        private void btnImport_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.CreateUpdateARMPExceptions(AMGenkResources.GetAMGenkExceptions(Properties.Settings.Default.AMGenkResourcesDirectory, Globals.ThisAddIn.ARMPStrtDate, Globals.ThisAddIn.ARMPFnshDate));
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

    }
}

