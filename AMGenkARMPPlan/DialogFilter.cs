using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AMGenkARMPPlan
{
    public partial class DialogFilter : Form
    {
        public DialogFilter()
        {
            InitializeComponent();
        }

        private void DialogFilter_Load(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox cbox;

            int iLoc_H = 10;
            int iLoc_V = 60;
            foreach (Resource ARMPResource in Globals.ThisAddIn.ARMPPlanWorksheetLayout.ARMPResources)
            {
                cbox = new CheckBox();
                cbox.Name = "Resource_" + iLoc_V.ToString();
                cbox.Tag = ARMPResource.Amei.ToString();
                cbox.Text = ARMPResource.Name.ToString();
                cbox.Checked = ARMPResource.Show;
                cbox.AutoSize = true;
                cbox.Location = new Point(iLoc_H, iLoc_V); //vertical
                this.Controls.Add(cbox);
                iLoc_V += 25;
            }
        }

        public void btnFilteren_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ARMPPlanWorksheetLayout.ARMPResourcesFiltered.Clear();
            foreach (Control Ctrl in this.Controls)
            {
                if ((Ctrl.GetType() == typeof(CheckBox)) && (Ctrl.Name.Substring(0, 8) == "Resource"))
                    if (((CheckBox)Ctrl).Checked)
                        Globals.ThisAddIn.ARMPPlanWorksheetLayout.ARMPResourcesFiltered.Add(
                            Globals.ThisAddIn.ARMPPlanWorksheetLayout.ARMPResources.Find(x => x.Name == ((CheckBox)Ctrl).Text));
            }

            Globals.ThisAddIn.FilterARMPPlanning();

            HideFilterDialogBox(); ;
        }

        private void cbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control ctrl in this.Controls)
            {
                if ((ctrl.GetType() == typeof(CheckBox)) && (ctrl.Name.Substring(0, 8) == "Resource"))
                {
                    ((CheckBox)ctrl).Checked = cbSelectAll.Checked;
                }
            }

        }

        private void btnAnuleren_Click(object sender, EventArgs e)
        {
            HideFilterDialogBox();
        }

        private void HideFilterDialogBox()
        {
            this.Hide();
        }
    }
}
