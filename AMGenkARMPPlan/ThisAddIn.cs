
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AMGenkARMPPlan
{
    public partial class ThisAddIn
    {
        private Office.CommandBar commandBar;
        private Office.CommandBarButton importButton;
        private Office.CommandBarButton atasksButton;
        private Office.CommandBarButton resourcesButton;
        private Office.CommandBarButton filterButton;
        private Office.CommandBarButton personalButton;

        private DataTable ARMPExceptionCodes = new DataTable();


        public ARMPPlanExcelLayout ARMPPlanWorksheetLayout = new ARMPPlanExcelLayout();

        public DateTime ARMPStrtDate { get { return ARMPPlanWorksheetLayout.ARMPStrtDate; } set { ARMPPlanWorksheetLayout.ARMPStrtDate = value; } }
        public DateTime ARMPFnshDate { get { return ARMPPlanWorksheetLayout.ARMPFnshDate; } set { ARMPPlanWorksheetLayout.ARMPFnshDate = value; } }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddImportToolbar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        private void AddImportToolbar()
        {
            try
            {
                commandBar = Application.CommandBars["ImportBar"];
            }
            catch (ArgumentException e)
            {
                // The toolbar named ImportBar does not exist, so create it.
            }
            if (commandBar == null)
            {
                commandBar = Application.CommandBars.Add("ImportBar", 1, false, true);
            }
            try
            {
                // Add a button and an event handler.
                importButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                importButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                importButton.Caption = "Importeer Resources en Taken";
                // TODO: JvdP Image+text CommandbarButton
                //importButton.Picture = 
                importButton.Tag = "A";
                importButton.TooltipText = "Importeer resources and taken uit Excel.";
                importButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                // Add a button and an event handler.
                atasksButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                atasksButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                atasksButton.Caption = "Update taken";
                atasksButton.Tag = "B";
                atasksButton.TooltipText = "Update taken uit Excel.";
                atasksButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                // Add a button and an event handler.
                resourcesButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                resourcesButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                resourcesButton.Caption = "Update aanwezigheden";
                resourcesButton.Tag = "C";
                resourcesButton.TooltipText = "Update aanwezigheden uit Excel.";
                resourcesButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                // Add a button and an event handler.
                filterButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                filterButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                filterButton.Caption = "Filteren van medewerkers";
                // TODO: JvdP Image+text CommandbarButton
                //importButton.Picture = 
                filterButton.Tag = "D";
                filterButton.TooltipText = "Filter medewerkers.";
                filterButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                // Add a button and an event handler.
                personalButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                personalButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                personalButton.Caption = "Persoonlijke planningen";
                // TODO: JvdP Image+text CommandbarButton
                //importButton.Picture = 
                personalButton.Tag = "E";
                personalButton.TooltipText = "Creëer persoonlijke planningen.";
                personalButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ButtonClick(Office.CommandBarButton ctrl, ref bool cancel)
        {
            try
            {
                Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);
                InitialiseARMPPlanWorksheetLayout();
            }
            catch
            { }

            switch (ctrl.Tag)
            {
                case "A":
                    DialogResourcesTasks dlgRTDialog = new DialogResourcesTasks();
                    dlgRTDialog.Show();
                    break;
                case "B":
                    DialogTasks dlgTDialog = new DialogTasks();
                    dlgTDialog.Show();
                    break;
                case "C":
                    DialogResources dlgRDialog = new DialogResources();
                    dlgRDialog.Show();
                    break;
                case "D":
                    DialogFilter dlgFDialog = new DialogFilter();
                    dlgFDialog.ShowDialog();
                    break;
                case "E":
                    CreatePersonalPlannings();
                    break;
                default:
                    break;
            }
        }
        public void CreateARMPPlanWorksheet()
        {
            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["Blad1"]);
            ARMPPlanWorksheet.Name = "PLAN";
            ARMPPlanWorksheet.Tab.Color = Color.Green;
        }

        public void SetARMPWorkplaces(Object[,] tasks)
        {
            for (int i = 1; i < tasks.GetLength(0) + 1; i++)
            {
                if (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkPlce] != null)
                {
                    if (!ARMPPlanWorksheetLayout.ARMPWorkplaces.Contains(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkPlce].ToString()))
                        ARMPPlanWorksheetLayout.ARMPWorkplaces.Add(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].ToString());
                }
            }
        }

        public void CreateARMPExceptionCodes(Object[,] exceptioncodes)
        {
            ARMPExceptionCodes.Columns.Clear();
            ARMPExceptionCodes.Rows.Clear();
            // Convert the 2 dimensional object array in a typed datatable
            ARMPExceptionCodes.Columns.Add(new DataColumn("ExcdCode", Type.GetType("System.String")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("ExcdAbbr", Type.GetType("System.String")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("ExcdTime", Type.GetType("System.TimeSpan")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("ExcdDayp", Type.GetType("System.String")));


            for (int i = (int)ARMPPlanExcelLayout.ARMPExceptionCodesRowsImpr.ExcdStrt; i < exceptioncodes.GetLength(0) + 1; i++)
            {
                if (exceptioncodes[i, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsImpr.ExcdCode] != null)
                {
                    DataRow ARMPExceptionCodesRow = ARMPExceptionCodes.NewRow();
                    ARMPExceptionCodesRow["ExcdCode"] = exceptioncodes[i, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsImpr.ExcdCode].ToString();
                    ARMPExceptionCodesRow["ExcdAbbr"] = exceptioncodes[i, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsImpr.ExcdAbbr].ToString();
                    ARMPExceptionCodesRow["ExcdTime"] = TimeSpan.Parse(DateTime.FromOADate((double)exceptioncodes[i, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsImpr.ExcdTime]).ToString("HH:mm:ss"));
                    ARMPExceptionCodesRow["ExcdDayp"] = exceptioncodes[i, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsImpr.ExcdDayp].ToString();
                    ARMPExceptionCodes.Rows.Add(ARMPExceptionCodesRow);
                }
            }
        }
        public void CreateARMPResources(Object[,] resources)
        {
            DateTime ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;
            /// TODO: Introducing break - JvdP 20170727
            /// TODO: Variable starting hours - JvdP 20170727
            // DateTime resourceStartTime = DateTime.FromOADate((double)resourcesCells[i, 2]);
            DateTime resourceStartTime = DateTime.Parse("08:00:00");
            // DateTime resourceFinishTime = DateTime.Parse("17:00:00") - (DateTime.Parse("09:00:00") - resourceStartTime);
            DateTime resourceFinishTime = DateTime.Parse("16:00:00");
            // V1.1.2.3 JvdP - Clear resources before start
            ARMPPlanWorksheetLayout.ARMPResources.Clear();

            ARMPPlanWorksheetLayout.ARMPResourcesRow = (int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcName;
            ARMPPlanWorksheetLayout.ARMPResourcesCol = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;

            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);


            for (int i = (int)ARMPPlanExcelLayout.ARMPResourcesRowsImpr.RsrcStrt; i < resources.GetLength(0) + 1; i++)
            {
                if (resources[i, (int)ARMPPlanExcelLayout.ARMPResourcesColsImpr.WorkPlce] != null)
                {
                    if (ARMPPlanWorksheetLayout.ARMPWorkplaces.Contains(resources[i, (int)ARMPPlanExcelLayout.ARMPResourcesColsImpr.WorkPlce].ToString()))
                    {
                        if (ARMPPlanWorksheetLayout.ARMPResources.FindIndex(x => x.Amei == resources[i, (int)ARMPPlanExcelLayout.ARMPResourcesColsImpr.RsrcAmei].ToString()) < 0)
                        {
                            ARMPPlanWorksheetLayout.ARMPResources.Add(new Resource()
                            {
                                Name = resources[i, (int)ARMPPlanExcelLayout.ARMPResourcesColsImpr.RsrcName].ToString(),
                                Amei = resources[i, (int)ARMPPlanExcelLayout.ARMPResourcesColsImpr.RsrcAmei].ToString(),
                                Show = true
                            });
                        }
                    }
                }
            }
            // Add two general resources
            ARMPPlanWorksheetLayout.ARMPResources.Add(new Resource()
            {
                Name = "EXTERN",
                Amei = "999998",
                Show = true
            });
            ARMPPlanWorksheetLayout.ARMPResources.Add(new AMGenkARMPPlan.Resource()
            {
                Name = "PRODUCTIE",
                Amei = "999999",
                Show = true
            });
            do
            {
                //ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPPlanWorksheetLayout.ARMPResourcesCol].Value2 = ARMPStrtDate.ToOADate();
                foreach (Resource resource in ARMPPlanWorksheetLayout.ARMPResources)
                {
                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPPlanWorksheetLayout.ARMPResourcesCol].Value2 = ARMPStrtDate.ToOADate();
                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, ARMPPlanWorksheetLayout.ARMPResourcesCol].Value2 = resource.Amei;
                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcName, ARMPPlanWorksheetLayout.ARMPResourcesCol].Value2 = resource.Name;
                    ARMPPlanWorksheetLayout.ARMPResourcesCol++;
                }
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);
        }
        public void CreateUpdateARMPExceptions(Object[,] exceptions)
        {
            DateTime ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;

            ARMPPlanWorksheetLayout.ARMPExceptionsRow = (int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd;
            ARMPPlanWorksheetLayout.ARMPExceptionsCol = (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt;

            Excel.Worksheet ARMPPlanWorksheet = (Excel.Worksheet)Application.Sheets["PLAN"];

            do
            {
                for (int i = 1; i < exceptions.GetLength(0) + 1; i++)
                {
                    int ARMPResource = ARMPPlanWorksheetLayout.ARMPResources.FindIndex(x => x.Name == exceptions[i, (int)ARMPPlanExcelLayout.ARMPExceptionsColsImpr.RsrcName]?.ToString());
                    if (ARMPResource >= 0)
                    {
                        for (int j = (int)ARMPPlanExcelLayout.ARMPExceptionsColsImpr.ExcpStrt; j <= exceptions.GetLength(1); j++)
                        {
                            ARMPPlanWorksheetLayout.ARMPExceptionsCol = (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt + ((j - (int)ARMPPlanExcelLayout.ARMPExceptionsColsImpr.ExcpStrt) * ARMPPlanWorksheetLayout.ARMPResources.Count) + ARMPResource;

                            string[] ARMPExceptionList;
                            string[] separators = { "/" };

                            string ARMPCode = "";
                            TimeSpan ARMPWorkTime = TimeSpan.Parse("08:00:00");

                            TimeSpan TijdV, TijdL, TijdD;

                            if (exceptions[i, j] == null)
                            {
                                ARMPCode = "";
                                ARMPWorkTime = TimeSpan.Parse("08:00:00");
                            }
                            else
                            {
                                ARMPExceptionList = exceptions[i, j].ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries);
                                // Convert combination of codes
                                if (ARMPExceptionList.Count() > 1)
                                {
                                    TijdV = TimeSpan.Parse("00:00:00");
                                    TijdL = TimeSpan.Parse("00:00:00");
                                    TijdD = TimeSpan.Parse("00:00:00");
                                    foreach (string ARMPException in ARMPExceptionList)
                                    {
                                        try
                                        {
                                            DataRow Exception = (from myrow in ARMPExceptionCodes.AsEnumerable()
                                                                 where myrow.Field<string>("ExcdCode") == ARMPException
                                                                 select myrow).SingleOrDefault();
                                            switch (Exception["ExcdDayp"].ToString())
                                            {
                                                case "D":
                                                    TijdD = TijdD.Add(TimeSpan.Parse("08:00:00"));
                                                    break;
                                                case "V":
                                                    TijdV = TijdV.Add((TimeSpan)Exception["ExcdTime"]);
                                                    break;
                                                case "L":
                                                    TijdL = TijdL.Add((TimeSpan)Exception["ExcdTime"]);
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                        catch
                                        {
                                            // Code not known - skip
                                        }
                                    }
                                    if (TijdD.CompareTo(TimeSpan.Parse("08:00:00")) >= 0)
                                    {
                                        ARMPCode = "Combi";
                                        ARMPWorkTime = TimeSpan.Parse("00:00:00");
                                    }
                                    else
                                    {
                                        TijdD = TijdV.Add(TijdL);
                                        if (TijdD.CompareTo(TimeSpan.Parse("08:00:00")) >= 0)
                                        {
                                            ARMPCode = "Combi";
                                            ARMPWorkTime = TimeSpan.Parse("00:00:00");
                                        }
                                        else
                                        {
                                            ARMPCode = "Combi";
                                            ARMPWorkTime = TimeSpan.Parse("00:00:00").Subtract(TijdV).Subtract(TijdL);
                                        }
                                    }
                                }
                                else if (ARMPExceptionList.Count() == 1)
                                {
                                    try
                                    {
                                        DataRow test = (from myrow in ARMPExceptionCodes.AsEnumerable()
                                                        where myrow.Field<string>("ExcdCode") == ARMPExceptionList[0].ToString()
                                                        select myrow).SingleOrDefault();
                                        ARMPCode = test["ExcdCode"].ToString();
                                        ARMPWorkTime = TimeSpan.Parse("08:00:00").Subtract((TimeSpan)test["ExcdTime"]);
                                    }
                                    catch
                                    {
                                        // Code not known - skip
                                    }
                                }
                            }
                            ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, ARMPPlanWorksheetLayout.ARMPExceptionsCol].Value2 = ARMPCode;
                            ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork, ARMPPlanWorksheetLayout.ARMPExceptionsCol].Value2 = ARMPWorkTime.TotalHours;
                            // TODO: JvdP only task cells in formula -split
                        }
                    }
                }
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);

            TimeSpan ARMPDays = ARMPPlanWorksheetLayout.ARMPFnshDate.Subtract(ARMPPlanWorksheetLayout.ARMPStrtDate);
            ARMPPlanWorksheetLayout.ARMPExceptionsCol = (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt + (ARMPPlanWorksheetLayout.ARMPResources.Count * ((int)ARMPDays.TotalDays + 1)) - 1;

            Excel.Range rngFormula;
            rngFormula = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                             ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormula.Formula = "=r[-1]c[0] - r[+1]c[0]";
            rngFormula.FormulaHidden = true;
            rngFormula.Calculate();
            rngFormula = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                             ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormula.Formula = "=SUM(r[1]c[0]:r[9999]c[0])";
            rngFormula.FormulaHidden = true;
            rngFormula.Calculate();
        }
        public void UpdateARMPExceptions(Object[,] exceptions)
        {
            DateTime ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;
        }
        public void PrepareARMPTasks()
        {
            // V1.1.2.13 JvdP - New prioroties including prio 0
            ARMPPlanWorksheetLayout.ARMPTasksRow0  = (int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt;
            ARMPPlanWorksheetLayout.ARMPTasksRow1  = ARMPPlanWorksheetLayout.ARMPTasksRow0 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow2  = ARMPPlanWorksheetLayout.ARMPTasksRow1 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow3  = ARMPPlanWorksheetLayout.ARMPTasksRow2 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow4  = ARMPPlanWorksheetLayout.ARMPTasksRow3 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow5  = ARMPPlanWorksheetLayout.ARMPTasksRow4 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow6  = ARMPPlanWorksheetLayout.ARMPTasksRow5 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow7  = ARMPPlanWorksheetLayout.ARMPTasksRow6 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow8  = ARMPPlanWorksheetLayout.ARMPTasksRow7 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow9  = ARMPPlanWorksheetLayout.ARMPTasksRow8 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow10 = ARMPPlanWorksheetLayout.ARMPTasksRow9 + 1;
            ARMPPlanWorksheetLayout.ARMPTasksRow = ARMPPlanWorksheetLayout.ARMPTasksRow10;

            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);

            foreach (ARMPPlanExcelLayout.ARMPTasksColsCnvt eCol in Enum.GetValues(typeof(ARMPPlanExcelLayout.ARMPTasksColsCnvt)))
            {
                ARMPPlanWorksheet.Cells[ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskTitl, eCol].Value2 = ARMPPlanWorksheetLayout.ARMPTasksColsHead[(int)eCol - 1];
            }

            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow,   1].Value2 = "PRIORITEIT 0";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow1,  1].Value2 = "PRIORITEIT 1";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow2,  1].Value2 = "PRIORITEIT 2";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow3,  1].Value2 = "PRIORITEIT 3";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow4,  1].Value2 = "PRIORITEIT 4";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5,  1].Value2 = "PRIORITEIT 5";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow6,  1].Value2 = "PRIORITEIT 6";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow7,  1].Value2 = "PRIORITEIT 7";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow8,  1].Value2 = "PRIORITEIT 8";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow9,  1].Value2 = "PRIORITEIT 9";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, 1].Value2 = "EINDE TAAKLIJST";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow0,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 0";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow1,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 1";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow2,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 2";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow3,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 3";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow4,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 4";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 5";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow6,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 6";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow7,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 7";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow8,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 8";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow9,  ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "PRIORITEIT 9";
            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "EINDE TAAKLIJST";
        }

        public object[,] ImportTasksFromClipboard()
        {
            int lastRowIgnoreFormulas;
            Excel.Worksheet ARMPPlanWorksheet;
            Excel.Worksheet ARMPImportWorksheet;

            string strDateForm = "dd/mm/jjjj";
            // Change the column definitions if your LayOut variant is changing
            List<Tuple<string, string>> lstSAPColumnHeaders = new List<Tuple<string, string>>();
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Werkplek", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Main Work Center", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Prioriteit", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Order", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Operatie", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Inperk. start", strDateForm));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Inperk. einde", strDateForm));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Order Basic Start Date", strDateForm));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Uitvoerdatum Gatekeeper", strDateForm));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Order Description", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Omschrijving operatie", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Toestand techn.syst.", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Ordersoort", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Gebruikersstatus", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Revisie", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Aantal", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Normale duur", "0.00"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Eenheid duur normaal", "@"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Werk", "0.00"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Eenheid werk", "0.00"));
            lstSAPColumnHeaders.Add(new Tuple<string, string>("Werkelijk werk", "0.00"));

            try
            {
                ARMPPlanWorksheet = (Excel.Worksheet)Application.Sheets["PLAN"];
                ARMPImportWorksheet = ((Excel.Worksheet)Application.Sheets.Add(Before: ARMPPlanWorksheet));
            }
            catch 
            {
                ARMPImportWorksheet = ((Excel.Worksheet)Application.Sheets.Add());
            }
            DateTime DateName = DateTime.Now;
            ARMPImportWorksheet.Name = "IN" + DateName.ToString("yyMMddhhmmss");
            ARMPImportWorksheet.Tab.Color = Color.Yellow;
            ARMPImportWorksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;

            char[] delimiters = new char[] { '\t' };
            StringReader strReader = new StringReader(Clipboard.GetText());

            int iRow = 1;
            int iCol = 1;
            foreach (Tuple<string, string> tplSAPColumnHeader in lstSAPColumnHeaders)
            {
                ARMPImportWorksheet.Cells[1, iCol] = tplSAPColumnHeader.Item1;
                ARMPImportWorksheet.Cells[1, iCol].EntireColumn.NumberFormat = tplSAPColumnHeader.Item2;
                iCol++;
            }
            iRow = 2;
            iCol = 1;
            while (true)
            {
                string strTask = strReader.ReadLine();
                if (strTask == null)
                    break;
                string[] strTaskFields = strTask.Split(delimiters);
                foreach (string value in strTaskFields)
                {
                    if (!string.IsNullOrEmpty(value))
                    {
                        if (lstSAPColumnHeaders[iCol - 1].Item2 == strDateForm)
                        {
                            DateTime dtDate = DateTime.Parse(value);
                            ARMPImportWorksheet.Cells[iRow, iCol] = dtDate;
                        }
                        else
                            ARMPImportWorksheet.Cells[iRow, iCol] = value;
                    }
                    iCol++;
                }
                iCol = 1;
                iRow++;
            }

            lastRowIgnoreFormulas = ARMPImportWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range tasksRange = ARMPImportWorksheet.Range[ARMPImportWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsOrig.TaskRows, (int)ARMPPlanExcelLayout.ARMPTasksColsOrig.WorkPlce],
                                                               ARMPImportWorksheet.Cells[lastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPTasksColsOrig.WorkReal]];
            return ((Object[,])tasksRange.Cells.Value2);
        }
        public void InitialiseARMPPlanWorksheetLayout()
        {
            int iARMPStrtDateCol = 0;
            int iARMPNextDateCol = 0;
            int iARMPLastDateCol = 0;

            int iLastRowIgnoreFormulas = 1;

            Excel.Range rngSearch;

            try
            {
                // First find start of priorities
                Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);

                ARMPStrtDate = ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt].Value;
                ARMPFnshDate = ARMPStrtDate.AddDays(4);

                iLastRowIgnoreFormulas = ARMPPlanWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                for (int i = (int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt; i <= iLastRowIgnoreFormulas; i++)
                {
                    switch ((string)ARMPPlanWorksheet.Cells[i, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].Value2)
                    {
                        case "0":
                            ARMPPlanWorksheetLayout.ARMPTasksRow0 = i;
                            break;
                        case "1":
                            ARMPPlanWorksheetLayout.ARMPTasksRow1 = i;
                            break;
                        case "2":
                            ARMPPlanWorksheetLayout.ARMPTasksRow2 = i;
                            break;
                        case "3":
                            ARMPPlanWorksheetLayout.ARMPTasksRow3 = i;
                            break;
                        case "4":
                            ARMPPlanWorksheetLayout.ARMPTasksRow4 = i;
                            break;
                        case "5":
                            ARMPPlanWorksheetLayout.ARMPTasksRow5 = i;
                            break;
                        case "6":
                            ARMPPlanWorksheetLayout.ARMPTasksRow6 = i;
                            break;
                        case "7":
                            ARMPPlanWorksheetLayout.ARMPTasksRow7 = i;
                            break;
                        case "8":
                            ARMPPlanWorksheetLayout.ARMPTasksRow8 = i;
                            break;
                        case "9":
                            ARMPPlanWorksheetLayout.ARMPTasksRow9 = i;
                            break;
                        case "EINDE TAAKLIJST":
                            ARMPPlanWorksheetLayout.ARMPTasksRow10 = i;
                            goto label_einde;
                        default:
                            break;
                    }
                }
            label_einde:
                ARMPPlanWorksheetLayout.ARMPResourcesCol = ARMPPlanWorksheet.Cells[1, 1].EntireRow.Columns.Count;
                ARMPPlanWorksheetLayout.ARMPResourcesCol = ARMPPlanWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                ARMPPlanWorksheetLayout.ARMPExceptionsCol = ARMPPlanWorksheetLayout.ARMPResourcesCol;
                ARMPPlanWorksheetLayout.ARMPTasksCol = ARMPPlanWorksheetLayout.ARMPResourcesCol;

                iARMPStrtDateCol = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;
                rngSearch = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                                ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPPlanWorksheetLayout.ARMPResourcesCol]];
                iARMPLastDateCol = rngSearch.FindPrevious().Column;

                ARMPPlanWorksheetLayout.ARMPResources.Clear();
                int iARMPRsrcCol = iARMPStrtDateCol;
                string strDate = DateTime.FromOADate((double)ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPStrtDateCol].Value2).ToString("dd/MM/yyyy");

                do
                {
                    ARMPPlanWorksheetLayout.ARMPResources.Add(new AMGenkARMPPlan.Resource()
                    {
                        Name = ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcName, iARMPRsrcCol].Value2.ToString(),
                        Amei = ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, iARMPRsrcCol].Value2.ToString(),
                        Show = true
                    });
                    iARMPRsrcCol++;
                } while (strDate == DateTime.FromOADate((double)ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPRsrcCol].Value2).ToString("dd/MM/yyyy"));
                ARMPPlanWorksheetLayout.ARMPStrtDate = DateTime.FromOADate((double)ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPStrtDateCol].Value2);
                ARMPPlanWorksheetLayout.ARMPFnshDate = DateTime.FromOADate((double)ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPLastDateCol].Value2);
            }
            catch
            { }
        }

        public void CreateUpdateARMPTasks(Object[,] tasks)
        {
            int ARMPTasksRowStrt = (int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt;
            int ARMPTasksRowFnsh = ARMPTasksRowStrt;
            int ARMPTasksRow = ARMPTasksRowFnsh;

            Double dOrdrStrtTarg = 0.0;
            Double dOrdrStrtSrce = 0.0;

            int iOrdrNmbrTarg = 0;
            int iOrdrNmbrSrce = 0;
            int iOperNmbrTarg = 0;
            int iOperNmbrSrce = 0;

            string ARMPWorkPlanForm = "=SUM(r[0]c[1]:r[0]C[" + (ARMPPlanWorksheetLayout.ARMPResourcesCol - (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            string ARMPWorkTodoForm = "=R[0]C[-2] - R[0]C[-1] - R[0]C[1]";

            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);
            for (int i = (int)ARMPPlanExcelLayout.ARMPTasksRowsImpr.TaskRows; i <= tasks.GetLength(0); i++)
            {
                if (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio] == null)
                    tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio] = 'O';
                // Priority A tasks
                switch (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].ToString())
                {
                    case "0":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow0 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow1;
                        ARMPPlanWorksheetLayout.ARMPTasksRow1++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow2++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow3++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow4++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow5++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        //  ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    case "1":
                    case "A":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow1 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow2;
                        ARMPPlanWorksheetLayout.ARMPTasksRow2++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow3++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow4++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow5++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        //  ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;

                    case "2":
                    case "B":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow2 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow3;
                        ARMPPlanWorksheetLayout.ARMPTasksRow3++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow4++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow5++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;

                    case "3":
                    case "C":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow3 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow4;
                        ARMPPlanWorksheetLayout.ARMPTasksRow4++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow5++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;

                    case "4":
                    case "D":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow4 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow5;
                        ARMPPlanWorksheetLayout.ARMPTasksRow5++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    case "5":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow5 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow6;
                        ARMPPlanWorksheetLayout.ARMPTasksRow6++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    case "6":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow6 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow7;
                        ARMPPlanWorksheetLayout.ARMPTasksRow7++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    case "7":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow7 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow8;
                        ARMPPlanWorksheetLayout.ARMPTasksRow8++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    case "8":
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow8 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow9;
                        ARMPPlanWorksheetLayout.ARMPTasksRow9++;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                    default:
                        ARMPTasksRowStrt = ARMPPlanWorksheetLayout.ARMPTasksRow9 + 1;
                        ARMPTasksRowFnsh = ARMPPlanWorksheetLayout.ARMPTasksRow10;
                        ARMPPlanWorksheetLayout.ARMPTasksRow10++;
                        // ARMPPlanWorksheetLayout.ARMPTasksRow++;
                        break;
                }

                // First order task - Order is Basic Start Date Ascending - OrderNmbrer Ascending
                ARMPTasksRow = ARMPTasksRowFnsh;
                if (ARMPTasksRowStrt != ARMPTasksRowFnsh)
                {
                    for (int j = ARMPTasksRowStrt; j < ARMPTasksRowFnsh; j++)
                    {
                        try
                        {
                            dOrdrStrtTarg = (double)ARMPPlanWorksheet.Cells[j, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].Value2;
                        }
                        catch (Exception)
                        {
                            dOrdrStrtTarg = 0.0;
                        }
                        try
                        {
                            dOrdrStrtSrce = (double)((tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrStrt] : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrStrt]);
                        }
                        catch (Exception)
                        {
                            dOrdrStrtSrce = 0.0;
                        }
                        try
                        {
                            iOrdrNmbrTarg = Int32.Parse(ARMPPlanWorksheet.Cells[j, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2.ToString());
                        }
                        catch
                        {
                            iOrdrNmbrTarg = 0;
                        }
                        try
                        {
                            iOrdrNmbrSrce = Int32.Parse(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrNmbr].ToString());
                        }
                        catch
                        {
                            iOrdrNmbrSrce = 0;
                        }
                        try
                        {
                            iOperNmbrTarg = Int32.Parse(ARMPPlanWorksheet.Cells[j, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].Value2.ToString());
                        }
                        catch
                        {
                            iOperNmbrTarg = 0;
                        }
                        try
                        {
                            iOperNmbrSrce = Int32.Parse(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OperNmbr].ToString());
                        }
                        catch
                        {
                            iOperNmbrSrce = 0;
                        }

                        if (dOrdrStrtTarg > dOrdrStrtSrce)
                        {
                            ARMPTasksRow = j;
                            break;
                        }
                        else if (dOrdrStrtTarg == dOrdrStrtSrce)
                        {
                            if (iOrdrNmbrTarg > iOrdrNmbrSrce)
                            {
                                ARMPTasksRow = j;
                                break;
                            }
                            else if (iOrdrNmbrTarg == iOrdrNmbrSrce)
                            {
                                if (iOperNmbrTarg > iOperNmbrSrce)
                                {
                                    ARMPTasksRow = j;
                                    break;
                                }
                            }
                        }
                    }
                }

                ARMPPlanWorksheet.Cells[ARMPTasksRow, 1].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                // V2.1.1.4 20180104 JVDP: do not reformat (original format must be kept)
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireRow.Interior.Color = Color.White;
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireRow.Font.Color = Color.Black;
                // V2.1.1.4 20180104 JVDP: END


                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkPlce] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkPlce].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.MainWork].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.MainWork] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.MainWork].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrPrio] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrPrio].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrNmbr] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrNmbr].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OperNmbr] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OperNmbr].ToString();

                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrStrt] : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrStrt];
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.GateDate] : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.RstrFnsh];

                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrDesc].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrDesc] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OrdrDesc].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperDesc].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OperDesc] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.OperDesc].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkUnit].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkUnit] == null) ? "" : tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkUnit].ToString();
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkNorm].Value2 = (tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkNorm] == null) ? 0.0 : Convert.ToDouble((tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkNorm].ToString()));

                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkHour].Value2 = Conversions.TimeUnit2Todo(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkNorm]?.ToString(),
                                                                                                                                      tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkUnit]?.ToString());
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkReal].Value2 = Conversions.TimeUnit2Todo(tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkReal]?.ToString(),
                                                                                                                                      tasks[i, (int)ARMPPlanExcelLayout.ARMPTasksColsImpr.WorkUnit]?.ToString());
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo].Formula = ARMPWorkTodoForm;
                ARMPPlanWorksheet.Cells[ARMPTasksRow, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan].Formula = ARMPWorkPlanForm;
            }
            // Summary
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol1].Formula =
                 "=SUM(r[" + (ARMPPlanWorksheetLayout.ARMPTasksRow1 - (int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1).ToString() + "]c[0]:r[" + (ARMPPlanWorksheetLayout.ARMPTasksRow5 - (int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1 - 1).ToString() + "]C[0])";
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4].Formula =
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPPlanWorksheetLayout.ARMPResourcesCol - (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow2, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4].Formula =
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPPlanWorksheetLayout.ARMPResourcesCol - (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow3, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4].Formula =
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPPlanWorksheetLayout.ARMPResourcesCol - (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";

            ARMPPlanWorksheet.Activate();
        }
        public void FilterARMPPlanning()
        {
            Excel.Worksheet ARMPPlanWorksheet = (Excel.Worksheet)Application.Sheets["PLAN"];

            for (int iResourceCol = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt; iResourceCol <= ARMPPlanWorksheetLayout.ARMPResourcesCol; iResourceCol++)
            {
                if (Globals.ThisAddIn.ARMPPlanWorksheetLayout.ARMPResourcesFiltered.Any(rs => rs.Amei == ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, iResourceCol].Value2.ToString()))
                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, iResourceCol].EntireColumn.Hidden = false;
                else
                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, iResourceCol].EntireColumn.Hidden = true;
            }
        }
        public void FormatARMPPlanning()
        {
            Application.DisplayAlerts = false;
            Excel.Worksheet ARMPPlanWorksheet = (Excel.Worksheet)Application.Sheets["PLAN"];
            Excel.Range rngFormat;
            Excel.FormatCondition rngfmcCondition;

            DateTime ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, ARMPPlanWorksheetLayout.ARMPResourcesCol - 1]];
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Merge();
            rngFormat.Borders.Weight = Excel.XlBorderWeight.xlThick;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt - 1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Font.Bold = true;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol1],
                                            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow3, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4]];
            rngFormat.NumberFormat = "0.00";
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Interior.Color = Color.LightCyan;

            int RsrcCols = (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt;
            do
            {
                rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols],
                                                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols + ARMPPlanWorksheetLayout.ARMPResources.Count - 1]];
                //rngFormat.Merge();
                rngFormat.NumberFormat = "dd/mmm";
                //rngFormat.Orientation = 90;
                //rngFormat.Borders.Weight = Excel.XlBorderWeight.xlThick;

                rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols],
                                                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, RsrcCols + ARMPPlanWorksheetLayout.ARMPResources.Count - 1]];
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                RsrcCols = RsrcCols + ARMPPlanWorksheetLayout.ARMPResources.Count;
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);

            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, 1].EntireRow.RowHeight = 0;
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcName, 1].EntireRow.RowHeight = 100;
            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, 1].EntireRow.RowHeight = 0;

            for (int i = (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce; i <= (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan; i++)
            {
                rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskTitl, i],
                                                    ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskTitl + 1, i]];
                rngFormat.Merge();
                rngFormat.Orientation = 90;
                rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;
                //rngFormat.Style.WrapText = true;
            }

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                                ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.Orientation = 90;
            rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.NumberFormat = "0.00";
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                            ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
            rngfmcCondition.Font.Color = Color.Orange;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
            rngfmcCondition.Font.Color = Color.Green;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
            rngfmcCondition.Font.Color = Color.Red;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;
            int TaskCols = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;
            do
            {
                rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, TaskCols],
                                                ARMPPlanWorksheet.Cells[(int)ARMPPlanWorksheetLayout.ARMPTasksRow10, TaskCols + ARMPPlanWorksheetLayout.ARMPResources.Count - 1]];
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                TaskCols = TaskCols + ARMPPlanWorksheetLayout.ARMPResources.Count;
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr]];
            rngFormat.NumberFormat = "0000";

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkNorm],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.NumberFormat = "0.00";

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo]];
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
            rngfmcCondition.Font.Color = Color.Orange;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
            rngfmcCondition.Font.Color = Color.Green;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
            rngfmcCondition.Font.Color = Color.Red;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow0, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow0, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.DarkRed;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow1, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow1, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Red;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow2, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow2, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Orange;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow3, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow3, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Green;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow4, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow4, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow6, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow6, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow7, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow7, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow8, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow8, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow9, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow9, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, 1],
                                            ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Font.Color = Color.White;
            rngFormat.Interior.Color = Color.Blue;


            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireColumn.ColumnWidth = 0;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.MainWork].EntireColumn.ColumnWidth = 0;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.ColumnWidth = 0;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.ColumnWidth = 12;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.Numberformat = "0";
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.ColumnWidth = 3;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.Numberformat = "0";
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.Numberformat = "dd/mm/jjjj";
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.ColumnWidth = 10;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.Numberformat = "dd/mm/jjjj";
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.ColumnWidth = 10;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrDesc].EntireColumn.ColumnWidth = 40;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperDesc].EntireColumn.ColumnWidth = 40;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkNorm].EntireColumn.ColumnWidth = 6;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkUnit].EntireColumn.ColumnWidth = 5;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkHour].EntireColumn.ColumnWidth = 6;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkReal].EntireColumn.ColumnWidth = 6;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo].EntireColumn.ColumnWidth = 6;
            ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan].EntireColumn.ColumnWidth = 6;

            rngFormat = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt], ARMPPlanWorksheet.Cells[1, ARMPPlanWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.Columns.ColumnWidth = 6;
        }

        public void CreatePersonalPlannings()
        {
            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);
            Excel.Worksheet ARMPRsrcWorksheet;
            ARMPPersExcelLayout ARMPRsrcWorksheetLayout = new ARMPPersExcelLayout();

            Excel.Range rngARMPPlan;

            Excel.Range rngARMPPlanRsrc;
            Excel.Range rngARMPPlanOrdr;

            foreach (Resource resource in ARMPPlanWorksheetLayout.ARMPResources)
            {
                try
                {
                    ARMPRsrcWorksheet = ((Excel.Worksheet)Application.Sheets[resource.Name]);
                    ARMPRsrcWorksheet.Unprotect();
                    ARMPRsrcWorksheet.Cells.Clear();
                }
                catch
                {
                    ARMPRsrcWorksheet = (Excel.Worksheet)Application.Sheets.Add(After: Application.Sheets[Application.Sheets.Count]);
                    ARMPRsrcWorksheet.Name = resource.Name;
                    ARMPRsrcWorksheet.Tab.Color = Color.Aqua;
                }
                ARMPRsrcWorksheetLayout.CopyFromPlanLayout(ARMPPlanWorksheetLayout);

                rngARMPPlan = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                                      ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow10, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt - 1]];
                // Work around to avoid run time error: commented out lines do not work !
                //rngARMPRsrc = ARMPRsrcWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                //                                      ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt - 1]];
                //rngARMPPlan.Copy(rngARMPRsrc);
                string strARMPRsrc = ((Excel.Range)ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce]).Address;
                rngARMPPlan.Copy(ARMPRsrcWorksheet.Range[strARMPRsrc]);

                rngARMPPlanRsrc = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                                          ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, ARMPPlanWorksheetLayout.ARMPResourcesCol]];


                ARMPRsrcWorksheetLayout.ARMPResourcesCol = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;
                foreach (Excel.Range rngARMPPlanRsrcCell in rngARMPPlanRsrc.Cells)
                {
                    if (rngARMPPlanRsrcCell.Value2?.ToString() == resource.Amei)
                    {
                        rngARMPPlan = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, rngARMPPlanRsrcCell.Column],
                                                              ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, rngARMPPlanRsrcCell.Column]];
                        // Work around to avoid run time error: commented out lines do not work !
                        strARMPRsrc = ((Excel.Range)ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPRsrcWorksheetLayout.ARMPResourcesCol]).Address;
                        rngARMPPlan.Copy(ARMPRsrcWorksheet.Range[strARMPRsrc]);

                        ARMPRsrcWorksheetLayout.ARMPResourcesCol++;
                    }
                }

                Application.DisplayAlerts = false;
                Excel.Range rngFormat;
                Excel.FormatCondition rngfmcCondition;

                DateTime ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                                    ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow5, ARMPRsrcWorksheetLayout.ARMPResourcesCol - 1]];
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
                rngFormat.Merge();
                rngFormat.Borders.Weight = Excel.XlBorderWeight.xlThick;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
                rngFormat.Font.Bold = true;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol1],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPSummRowsCnvt.SummRow3, (int)ARMPPlanExcelLayout.ARMPSummColsCnvt.SummCol4]];
                rngFormat.NumberFormat = "0.00";
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Interior.Color = Color.LightCyan;

                int RsrcCols = (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt;
                do
                {
                    rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols],
                                                        ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols]];
                    rngFormat.NumberFormat = "dd/mmm";
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                    rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols],
                                                        ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, RsrcCols]];
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                    rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                    RsrcCols++;
                    ARMPStrtDate = ARMPStrtDate.AddDays(1);
                } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);
                ARMPRsrcWorksheetLayout.ARMPExceptionsCol = RsrcCols - 1;

                ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, 1].EntireRow.RowHeight = 0;
                ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcName, 1].EntireRow.RowHeight = 100;
                ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, 1].EntireRow.RowHeight = 0;

                for (int i = (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce; i <= (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan; i++)
                {
                    rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskTitl, i],
                                                        ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskTitl + 1, i]];
                    rngFormat.Merge();
                    rngFormat.Orientation = 90;
                    rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;
                    //rngFormat.Style.WrapText = true;
                }

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.Orientation = 90;
                rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.NumberFormat = "0.00";
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                                    ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
                rngfmcCondition.Font.Color = Color.Orange;
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
                rngfmcCondition.Font.Color = Color.Green;
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
                rngfmcCondition.Font.Color = Color.Red;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, 1],
                                                    ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow4, ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

                ARMPStrtDate = ARMPPlanWorksheetLayout.ARMPStrtDate;
                int TaskCols = (int)ARMPPlanExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;
                do
                {
                    rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, TaskCols],
                                                        ARMPRsrcWorksheet.Cells[(int)ARMPRsrcWorksheetLayout.ARMPTasksRow4, TaskCols]];
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                    rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                    rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                    TaskCols++;
                    ARMPStrtDate = ARMPStrtDate.AddDays(1);
                } while (ARMPStrtDate.CompareTo(ARMPPlanWorksheetLayout.ARMPFnshDate) <= 0);

                //rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr],
                //                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow4, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr]];
                //rngFormat.NumberFormat = "0000";

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkNorm],
                                                    ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow9, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.NumberFormat = "0.00";

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo],
                                                    ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow10, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo]];
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
                rngfmcCondition.Font.Color = Color.Orange;
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
                rngfmcCondition.Font.Color = Color.Green;
                rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
                rngfmcCondition.Font.Color = Color.Red;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow0, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow10, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.Interior.Color = Color.White;
                rngFormat.Font.Color = Color.Black;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow0, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow1, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Interior.Color = Color.DarkRed;
                rngFormat.Font.Color = Color.White;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow1, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow1, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Interior.Color = Color.Red;
                rngFormat.Font.Color = Color.White;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow2, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow2, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Interior.Color = Color.Orange;
                rngFormat.Font.Color = Color.White;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow3, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow3, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Interior.Color = Color.Green;
                rngFormat.Font.Color = Color.White;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow4, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow4, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Interior.Color = Color.LightBlue;
                rngFormat.Font.Color = Color.White;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow5, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow5, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.LightBlue;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow6, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow6, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.LightBlue;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow7, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow7, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.LightBlue;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow8, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow8, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.LightBlue;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow9, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow9, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.LightBlue;
                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow10, 1],
                                                ARMPRsrcWorksheet.Cells[ARMPRsrcWorksheetLayout.ARMPTasksRow10, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.FormatConditions.Delete();
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Font.Color = Color.White;
                rngFormat.Interior.Color = Color.Blue;



                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireColumn.ColumnWidth = 0;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.MainWork].EntireColumn.ColumnWidth = 0;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.ColumnWidth = 0;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.ColumnWidth = 12;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.Numberformat = "0";
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.ColumnWidth = 3;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.Numberformat = "0";
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.Numberformat = "dd/mm/jjjj";
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.ColumnWidth = 10;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.Numberformat = "dd/mm/jjjj";
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.ColumnWidth = 10;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrDesc].EntireColumn.ColumnWidth = 40;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperDesc].EntireColumn.ColumnWidth = 40;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkNorm].EntireColumn.ColumnWidth = 6;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkUnit].EntireColumn.ColumnWidth = 5;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkHour].EntireColumn.ColumnWidth = 6;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkReal].EntireColumn.ColumnWidth = 6;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkTodo].EntireColumn.ColumnWidth = 6;
                ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlan].EntireColumn.ColumnWidth = 6;

                rngFormat = ARMPRsrcWorksheet.Range[ARMPRsrcWorksheet.Cells[1, (int)ARMPPlanExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt], ARMPRsrcWorksheet.Cells[1, ARMPRsrcWorksheetLayout.ARMPExceptionsCol]];
                rngFormat.Columns.ColumnWidth = 6;

                ARMPRsrcWorksheet.Protect();
            }
        }

        /*
        public void CreateUpdateQRCodesSheet()
        {
            Excel.Worksheet ARMPPlanWorksheet = ((Excel.Worksheet)Application.Sheets["PLAN"]);
            Excel.Worksheet ARMPCodeWorksheet;

            Excel.Range rngARMPPlanOrdr;

            try
            {
                ARMPCodeWorksheet = ((Excel.Worksheet)Application.Sheets["QRCODE"]);
                ARMPCodeWorksheet.Unprotect();
                ARMPCodeWorksheet.Delete();
            }
            catch
            {
            }
            ARMPCodeWorksheet = (Excel.Worksheet)Application.Sheets.Add(After: ARMPPlanWorksheet);
            ARMPCodeWorksheet.Name = "QRCODE";
            ARMPCodeWorksheet.Tab.Color = Color.LightGreen;

            rngARMPPlanOrdr = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                                      ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkPlce]];
            rngARMPPlanOrdr.Copy(Type.Missing);
            ARMPCodeWorksheet.Range["A1"].PasteSpecial();

            rngARMPPlanOrdr = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrNmbr],
                                                          ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperNmbr]];

            rngARMPPlanOrdr.Copy(Type.Missing);
            ARMPCodeWorksheet.Range["B1"].PasteSpecial();

            rngARMPPlanOrdr = ARMPPlanWorksheet.Range[ARMPPlanWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OrdrDesc],
                                                          ARMPPlanWorksheet.Cells[ARMPPlanWorksheetLayout.ARMPTasksRow5, (int)ARMPPlanExcelLayout.ARMPTasksColsCnvt.OperDesc]];
            rngARMPPlanOrdr.Copy(Type.Missing);
            ARMPCodeWorksheet.Range["D1"].PasteSpecial();

            ARMPCodeExcelLayout ARMPCodeWorksheetLayout = new ARMPCodeExcelLayout();
            ARMPCodeWorksheetLayout.CopyFromPlanLayout(ARMPPlanWorksheetLayout);

            for (int iTask = 1; iTask <= (ARMPCodeWorksheetLayout.ARMPTasksRow5); iTask++)
            {
                int iOrdrNmbr;
                string strOrdrNmbr = ARMPCodeWorksheet.Cells[iTask, 2].Value2?.ToString();

                if (!string.IsNullOrEmpty(strOrdrNmbr) && Int32.TryParse(strOrdrNmbr, out iOrdrNmbr))
                {
                    // 1. Rijhoogte en uitlijning instellen
                    ARMPCodeWorksheet.Cells[iTask, 2].EntireRow.RowHeight = 60;
                    ARMPCodeWorksheet.Cells[iTask, 2].EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // Pak de doellocatie (cel A + rijnummer)
                    Excel.Range rngCode = ARMPCodeWorksheet.Range["A" + iTask.ToString()];
                    rngCode.Clear(); // Oude data/vormen in de cel wissen

                    // Maak een uniek tijdelijk bestandspad aan op de harde schijf
                    string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"QR_{Guid.NewGuid()}.png");

                    try
                    {
                        // 2. QR-code genereren en opslaan als tijdelijk bestand
                        using (ICodeGenerator qrGenerator = new QRCodeGenerator())
                        {
                            using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(strOrdrNmbr, QRCodeGenerator.ECCLevel.Q))
                            {
                                using (QRCode qrCode = new QRCode(qrCodeData))
                                {
                                    using (Bitmap qrCodeImage = qrCode.GetGraphic(2))
                                    {
                                        qrCodeImage.Save(tempFilePath, System.Drawing.Imaging.ImageFormat.Png);
                                    }
                                }
                            }
                        }

                        // 3. De QR-afbeelding direct in Excel laden via Shapes.AddPicture
                        // We initialiseren de positie op 0, de exacte positionering en schaling doen we direct hierna
                        Excel.Shape shp = ARMPCodeWorksheet.Shapes.AddPicture(
                            tempFilePath,
                            Microsoft.Office.Core.MsoTriState.msoFalse, // LinkToFile: Niet linken naar bestand op schijf
                            Microsoft.Office.Core.MsoTriState.msoTrue,  // SaveWithDocument: Sla de afbeelding op ín het Excel-bestand
                            0, 0, -1, -1                                // Positie (0,0) en -1 voor behoud van originele grootte
                        );

                        // 4. De QR-code exact centreren binnen Cel A[iTask]
                        // Formule: Cel-startpositie + ((Totale Celruimte - Afbeeldingsgrootte) / 2)
                        shp.Left = (float)(rngCode.Left + (rngCode.Width - shp.Width) / 2);
                        shp.Top = (float)(rngCode.Top + (rngCode.Height - shp.Height) / 2);
                    }
                    catch (COMException ex)
                    {
                        // Log of verwerk eventuele Excel Interop-fouten hier
                        // Bijvoorbeeld: Console.WriteLine("Excel fout: " + ex.Message);
                    }
                    finally
                    {
                        // 5. Netjes opruimen: Verwijder het tijdelijke PNG-bestand van de harde schijf
                        if (System.IO.File.Exists(tempFilePath))
                        {
                            try { System.IO.File.Delete(tempFilePath); } catch { // Negeren als bestand nog bezet is  }
                        }
                    }
                }
            }
            ARMPCodeWorksheet.Cells[1, (int)ARMPCodeExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.ColumnWidth = 10;
            ARMPCodeWorksheet.Cells[1, (int)ARMPCodeExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.ColumnWidth = 6;
            ARMPCodeWorksheet.Cells[1, (int)ARMPCodeExcelLayout.ARMPTasksColsCnvt.OrdrDesc].EntireColumn.ColumnWidth = 40;
            ARMPCodeWorksheet.Cells[1, (int)ARMPCodeExcelLayout.ARMPTasksColsCnvt.OperDesc].EntireColumn.ColumnWidth = 40;

            ARMPCodeWorksheet.Protect();

            ARMPPlanWorksheet.Activate();
        }*/
    }
}

