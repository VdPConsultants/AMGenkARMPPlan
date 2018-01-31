using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Drawing;

namespace AMGenkARMPPlan
{
    public partial class ThisAddIn
    {
        private Office.CommandBar commandBar;
        private Office.CommandBarButton importButton;
        private Office.CommandBarButton atasksButton;
        private Office.CommandBarButton resourcesButton;

        private DataTable ARMPExceptionCodes = new DataTable();


        private ARMPExcelLayout ARMPWorksheetLayout = new ARMPExcelLayout();

        public DateTime GetARMPStrtDate()
        {
            return ARMPWorksheetLayout.ARMPStrtDate;
        }

        public DateTime ARMPStrtDate { get { return ARMPWorksheetLayout.ARMPStrtDate; } set { ARMPWorksheetLayout.ARMPStrtDate = value; } }
        public DateTime ARMPFnshDate { get { return ARMPWorksheetLayout.ARMPFnshDate; } set { ARMPWorksheetLayout.ARMPFnshDate = value; } }

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
        }
        private void ButtonClick(Office.CommandBarButton ctrl, ref bool cancel)
        {
            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            InitialiseARMPWorksheetLayout();

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
                default:
                    break;
            }
        }
        public void SetARMPWorkplaces(Object[,] tasks)
        {
            for (int i = 1; i < tasks.GetLength(0) + 1; i++)
            {
                if (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkPlce] != null)
                {
                    if (!ARMPWorksheetLayout.ARMPWorkplaces.Contains(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkPlce].ToString()))
                        ARMPWorksheetLayout.ARMPWorkplaces.Add(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].ToString());
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


            for (int i = (int)ARMPExcelLayout.ARMPExceptionCodesRowsImpr.ExcdStrt; i < exceptioncodes.GetLength(0) + 1; i++)
            {
                if (exceptioncodes[i, (int)ARMPExcelLayout.ARMPExceptionCodesColsImpr.ExcdCode] != null)
                {
                    DataRow ARMPExceptionCodesRow = ARMPExceptionCodes.NewRow();
                    ARMPExceptionCodesRow["ExcdCode"] = exceptioncodes[i, (int)ARMPExcelLayout.ARMPExceptionCodesColsImpr.ExcdCode].ToString();
                    ARMPExceptionCodesRow["ExcdAbbr"] = exceptioncodes[i, (int)ARMPExcelLayout.ARMPExceptionCodesColsImpr.ExcdAbbr].ToString();
                    ARMPExceptionCodesRow["ExcdTime"] = TimeSpan.Parse(DateTime.FromOADate((double)exceptioncodes[i, (int)ARMPExcelLayout.ARMPExceptionCodesColsImpr.ExcdTime]).ToString("HH:mm:ss"));
                    ARMPExceptionCodesRow["ExcdDayp"] = exceptioncodes[i, (int)ARMPExcelLayout.ARMPExceptionCodesColsImpr.ExcdDayp].ToString();
                    ARMPExceptionCodes.Rows.Add(ARMPExceptionCodesRow);
                }
            }
        }
        public void CreateARMPResources(Object[,] resources)
        {
            DateTime ARMPStrtDate = ARMPWorksheetLayout.ARMPStrtDate;
            // TODO: Introducing break - JvdP 20170727
            // TODO: Variable starting hours - JvdP 20170727
            // DateTime resourceStartTime = DateTime.FromOADate((double)resourcesCells[i, 2]);
            DateTime resourceStartTime = DateTime.Parse("08:00:00");
            // DateTime resourceFinishTime = DateTime.Parse("17:00:00") - (DateTime.Parse("09:00:00") - resourceStartTime);
            DateTime resourceFinishTime = DateTime.Parse("16:00:00");
            // V1.1.2.3 JvdP - Clear resources before start
            ARMPWorksheetLayout.ARMPResources.Clear();

            ARMPWorksheetLayout.ARMPResourcesRow = (int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcName;
            ARMPWorksheetLayout.ARMPResourcesCol = (int)ARMPExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);


            for (int i = (int)ARMPExcelLayout.ARMPResourcesRowsImpr.RsrcStrt; i < resources.GetLength(0) + 1; i++)
            {
                if (resources[i, (int)ARMPExcelLayout.ARMPResourcesColsImpr.WorkPlce] != null)
                {
                    if (ARMPWorksheetLayout.ARMPWorkplaces.Contains(resources[i, (int)ARMPExcelLayout.ARMPResourcesColsImpr.WorkPlce].ToString()))
                    {
                        if (ARMPWorksheetLayout.ARMPResources.FindIndex(x => x.Amei == resources[i, (int)ARMPExcelLayout.ARMPResourcesColsImpr.RsrcAmei].ToString()) < 0)
                        {
                            ARMPWorksheetLayout.ARMPResources.Add( new AMGenkARMPPlan.Resource()
                            {
                                Name = resources[i, (int)ARMPExcelLayout.ARMPResourcesColsImpr.RsrcName].ToString(),
                                Amei = resources[i, (int)ARMPExcelLayout.ARMPResourcesColsImpr.RsrcAmei].ToString()
                            });
                        }
                    }
                }
            }
            // Add two general resources
            ARMPWorksheetLayout.ARMPResources.Add(new AMGenkARMPPlan.Resource()
            {
                Name = "EXTERN",
                Amei = "000000"
            });
            ARMPWorksheetLayout.ARMPResources.Add(new AMGenkARMPPlan.Resource()
            {
                Name = "PRODUCTIE",
                Amei = "000000"
            });
            do
            {
                ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPWorksheetLayout.ARMPResourcesCol].Value2 = ARMPStrtDate.ToOADate();
                foreach (Resource resource in ARMPWorksheetLayout.ARMPResources)
                {
                    ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, ARMPWorksheetLayout.ARMPResourcesCol].Value2 = resource.Amei;
                    ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcName, ARMPWorksheetLayout.ARMPResourcesCol].Value2 = resource.Name;
                    ARMPWorksheetLayout.ARMPResourcesCol++;
                }
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPWorksheetLayout.ARMPFnshDate) <= 0);
        }
        public void CreateUpdateARMPExceptions(Object[,] exceptions)
        {
            DateTime ARMPStrtDate = ARMPWorksheetLayout.ARMPStrtDate;

            ARMPWorksheetLayout.ARMPExceptionsRow = (int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd;
            ARMPWorksheetLayout.ARMPExceptionsCol = (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt;

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            do
            {
                for (int i = 1; i < exceptions.GetLength(0) + 1; i++)
                {
                    int ARMPResource = ARMPWorksheetLayout.ARMPResources.FindIndex(x => x.Name == exceptions[i, (int)ARMPExcelLayout.ARMPExceptionsColsImpr.RsrcName]?.ToString());
                    if (ARMPResource >= 0)
                    {
                        for (int j = (int)ARMPExcelLayout.ARMPExceptionsColsImpr.ExcpStrt; j <= exceptions.GetLength(1); j++)
                        {
                            ARMPWorksheetLayout.ARMPExceptionsCol = (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt + ((j - (int)ARMPExcelLayout.ARMPExceptionsColsImpr.ExcpStrt) * ARMPWorksheetLayout.ARMPResources.Count) + ARMPResource;

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
                            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, ARMPWorksheetLayout.ARMPExceptionsCol].Value2 = ARMPCode;
                            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork, ARMPWorksheetLayout.ARMPExceptionsCol].Value2 = ARMPWorkTime.TotalHours;
                            // TODO: JvdP only task cells in formula -split
                        }
                    }
                }
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPWorksheetLayout.ARMPFnshDate) <= 0) ;

            TimeSpan ARMPDays = ARMPWorksheetLayout.ARMPFnshDate.Subtract(ARMPWorksheetLayout.ARMPStrtDate);
            ARMPWorksheetLayout.ARMPExceptionsCol = (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt + (ARMPWorksheetLayout.ARMPResources.Count * ((int)ARMPDays.TotalDays + 1)) - 1;

            Excel.Range rngFormula;
            rngFormula = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                             ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormula.Formula = "=r[-1]c[0] - r[+1]c[0]";
            rngFormula.FormulaHidden = true;
            rngFormula.Calculate();
            rngFormula = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                             ARMPWorksheet.Cells[ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormula.Formula = "=SUM(r[1]c[0]:r[9999]c[0])";
            rngFormula.FormulaHidden = true;
            rngFormula.Calculate();
        }
        public void UpdateARMPExceptions(Object[,] exceptions)
        {
            DateTime ARMPStrtDate = ARMPWorksheetLayout.ARMPStrtDate;
        }
                public void PrepareARMPTasks()
        {
            ARMPWorksheetLayout.ARMPTasksRowA = (int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt;
            ARMPWorksheetLayout.ARMPTasksRowB = ARMPWorksheetLayout.ARMPTasksRowA + 1;
            ARMPWorksheetLayout.ARMPTasksRowC = ARMPWorksheetLayout.ARMPTasksRowB + 1;
            ARMPWorksheetLayout.ARMPTasksRowO = ARMPWorksheetLayout.ARMPTasksRowC + 1;
            ARMPWorksheetLayout.ARMPTasksRowZ = ARMPWorksheetLayout.ARMPTasksRowO + 1;
            ARMPWorksheetLayout.ARMPTasksRow = ARMPWorksheetLayout.ARMPTasksRowZ;

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].Value2 = "Uitvoerende werkplek";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.MainWork].Value2 = "Verantwoordelijke werkplek";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].Value2 = "Uitvoerende werkplek";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.OrdrPrio].Value2 = "Prioriteit";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = "Order nummer";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.OperNmbr].Value2 = "Operatie nummer";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].Value2 = "Start datum";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.RstrFnsh].Value2 = "Eind datum";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.OrdrDesc].Value2 = "Order beschrijving";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.OperDesc].Value2 = "Operatie beschrijving";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkUnit].Value2 = "Werktijd eenheid";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkNorm].Value2 = "Werktijd";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkHour].Value2 = "Werktijd in uur";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkReal].Value2 = "Werktijd gewerkt";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkTodo].Value2 = "Werktijd todo";
            ARMPWorksheet.Cells[ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan].Value2 = "Werktijd gepland";

            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowA, 1].Value2 = "PRIORITEIT A";
            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowB, 1].Value2 = "PRIORITEIT B";
            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowC, 1].Value2 = "PRIORITEIT C";
            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, 1].Value2 = "PRIORITEIT O";
            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowZ, 1].Value2 = "EINDE TAAKLIJST";
        }

        public void InitialiseARMPWorksheetLayout()
        {
            int iARMPStrtDateCol = 0;
            int iARMPNextDateCol = 0;
            int iARMPLastDateCol = 0;

            int iLastRowIgnoreFormulas = 1;

            Excel.Range rngSearch;

            try
            {
                // First find start of priorities
                Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

                ARMPStrtDate = ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPResourcesColsCnvt.RsrcStrt].Value;
                ARMPFnshDate = ARMPStrtDate.AddDays(4);

                iLastRowIgnoreFormulas = ARMPWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                for (int i = (int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt; i <= iLastRowIgnoreFormulas; i++)
                {
                    switch ((string)ARMPWorksheet.Cells[i, 1].Value2)
                    {
                        case "PRIORITEIT A":
                            ARMPWorksheetLayout.ARMPTasksRowA = i;
                            break;
                        case "PRIORITEIT B":
                            ARMPWorksheetLayout.ARMPTasksRowB = i;
                            break;
                        case "PRIORITEIT C":
                            ARMPWorksheetLayout.ARMPTasksRowC = i;
                            break;
                        case "PRIORITEIT O":
                            ARMPWorksheetLayout.ARMPTasksRowO = i;
                            break;
                        case "EINDE TAAKLIJST":
                            ARMPWorksheetLayout.ARMPTasksRowZ = i;
                            goto label_einde;
                        default:
                            break;
                    }
                }
            label_einde:

                ARMPWorksheetLayout.ARMPResourcesCol = ARMPWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                ARMPWorksheetLayout.ARMPExceptionsCol = ARMPWorksheetLayout.ARMPResourcesCol;
                ARMPWorksheetLayout.ARMPTasksCol = ARMPWorksheetLayout.ARMPResourcesCol;

                iARMPStrtDateCol = (int)ARMPExcelLayout.ARMPResourcesColsCnvt.RsrcStrt;
                rngSearch = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                                ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, ARMPWorksheetLayout.ARMPResourcesCol]];
                iARMPNextDateCol = rngSearch.FindNext().Column;
                iARMPLastDateCol = rngSearch.FindPrevious().Column;

                ARMPWorksheetLayout.ARMPResources.Clear();
                for (int i = iARMPStrtDateCol; i < iARMPNextDateCol; i++)
                {
                    ARMPWorksheetLayout.ARMPResources.Add(new AMGenkARMPPlan.Resource()
                    {
                        Name = ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcName, i].Value2.ToString(),
                        Amei = ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, i].Value2.ToString()
                    });
                }
                ARMPWorksheetLayout.ARMPStrtDate = DateTime.FromOADate((double)ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPStrtDateCol].Value2);
                ARMPWorksheetLayout.ARMPFnshDate = DateTime.FromOADate((double)ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, iARMPLastDateCol].Value2);
            }
            catch
            { }
        }

        public void CreateUpdateARMPTasks(Object[,] tasks)
        {
            int ARMPTasksRowStrt = (int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt;
            int ARMPTasksRowFnsh = ARMPTasksRowStrt;
            int ARMPTasksRow = ARMPTasksRowFnsh;
            
            Double dOrdrStrtTarg = 0.0;
            Double dOrdrStrtSrce = 0.0;

            int iOrdrNmbrTarg = 0;
            int iOrdrNmbrSrce = 0;
            int iOperNmbrTarg = 0;
            int iOperNmbrSrce = 0;

            string ARMPWorkPlanForm = "=SUM(r[0]c[1]:r[0]C[" + (ARMPWorksheetLayout.ARMPResourcesCol - (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            string ARMPWorkTodoForm = "=R[0]C[-2] - R[0]C[-1] - R[0]C[1]";

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            for (int i = (int)ARMPExcelLayout.ARMPTasksRowsImpr.TaskRows; i <= tasks.GetLength(0); i++)
            {
                // Priority A tasks
                switch (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrPrio].ToString())
                {
                    case "A":
                        ARMPTasksRowStrt = ARMPWorksheetLayout.ARMPTasksRowA + 1;
                        ARMPTasksRowFnsh = ARMPWorksheetLayout.ARMPTasksRowB;
                        ARMPWorksheetLayout.ARMPTasksRowB++;
                        ARMPWorksheetLayout.ARMPTasksRowC++;
                        ARMPWorksheetLayout.ARMPTasksRowO++;
                        ARMPWorksheetLayout.ARMPTasksRowZ++;
                        //  ARMPWorksheetLayout.ARMPTasksRow++;
                        break;

                    case "B":
                    case "1":
                    case "2":
                    case "3":
                    case "4":
                    case "5":
                    case "6":
                    case "7":
                    case "8":
                        ARMPTasksRowStrt = ARMPWorksheetLayout.ARMPTasksRowB + 1;
                        ARMPTasksRowFnsh = ARMPWorksheetLayout.ARMPTasksRowC;
                        ARMPWorksheetLayout.ARMPTasksRowC++;
                        ARMPWorksheetLayout.ARMPTasksRowO++;
                        ARMPWorksheetLayout.ARMPTasksRowZ++;
                        // ARMPWorksheetLayout.ARMPTasksRow++;
                        break;

                    case "C":
                        ARMPTasksRowStrt = ARMPWorksheetLayout.ARMPTasksRowC + 1;
                        ARMPTasksRowFnsh = ARMPWorksheetLayout.ARMPTasksRowO;
                        ARMPWorksheetLayout.ARMPTasksRowO++;
                        ARMPWorksheetLayout.ARMPTasksRowZ++;
                        // ARMPWorksheetLayout.ARMPTasksRow++;
                        break;

                    default:
                        ARMPTasksRowStrt = ARMPWorksheetLayout.ARMPTasksRowO + 1;
                        ARMPTasksRowFnsh = ARMPWorksheetLayout.ARMPTasksRowZ;
                        ARMPWorksheetLayout.ARMPTasksRowZ++;
                        // ARMPWorksheetLayout.ARMPTasksRow++;
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
                            dOrdrStrtTarg = (double)ARMPWorksheet.Cells[j, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].Value2;
                        }
                        catch (Exception)
                        {
                            dOrdrStrtTarg = 0.0;
                        }
                        try
                        {
                            dOrdrStrtSrce = (double)((tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrStrt] : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrStrt]);
                        }
                        catch (Exception)
                        {
                            dOrdrStrtSrce = 0.0;
                        }
                        try
                        {
                            iOrdrNmbrTarg = Int32.Parse(ARMPWorksheet.Cells[j, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2.ToString());
                        }
                        catch
                        {
                            iOrdrNmbrTarg = 0;
                        }
                        try
                        {
                            iOrdrNmbrSrce = Int32.Parse(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrNmbr].ToString());
                        }
                        catch
                        {
                            iOrdrNmbrSrce = 0;
                        }
                        try
                        {
                            iOperNmbrTarg = Int32.Parse(ARMPWorksheet.Cells[j, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperNmbr].Value2.ToString());
                        }
                        catch
                        {
                            iOperNmbrTarg = 0;
                        }
                        try
                        {
                            iOperNmbrSrce = Int32.Parse(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OperNmbr].ToString());
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

                ARMPWorksheet.Cells[ARMPTasksRow, 1].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                // V2.1.1.4 20180104 JVDP: do not reformat (original format must be kept)
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireRow.Interior.Color = Color.White;
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].EntireRow.Font.Color = Color.Black;
                // V2.1.1.4 20180104 JVDP: END


                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkPlce] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkPlce].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.MainWork].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.MainWork] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.MainWork].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrPrio].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrPrio] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrPrio].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrNmbr] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrNmbr].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperNmbr].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OperNmbr] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OperNmbr].ToString();
                
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrStrt] : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrStrt];
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrFnsh].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrStrt] == null) ? tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.GateDate] : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.RstrFnsh];

                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrDesc].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrDesc] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OrdrDesc].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperDesc].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OperDesc] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.OperDesc].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkUnit].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkUnit] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkUnit].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkNorm].Value2 = (tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkNorm] == null) ? "" : tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkNorm].ToString();

                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkHour].Value2 = Conversions.TimeUnit2Todo(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkNorm].ToString(),
                                                                                                                                      tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkUnit].ToString());
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkReal].Value2 = Conversions.TimeUnit2Todo(tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkReal].ToString(),
                                                                                                                                      tasks[i, (int)ARMPExcelLayout.ARMPTasksColsImpr.WorkUnit].ToString());
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkTodo].Formula = ARMPWorkTodoForm;
                ARMPWorksheet.Cells[ARMPTasksRow, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan].Formula = ARMPWorkPlanForm;
            }
            // Summary
            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol1].Formula = 
                 "=SUM(r[" + (ARMPWorksheetLayout.ARMPTasksRowA - (int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow1).ToString() + "]c[0]:r[" + (ARMPWorksheetLayout.ARMPTasksRowZ - (int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow1 - 1).ToString() + "]C[0])";
            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4].Formula =
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPWorksheetLayout.ARMPResourcesCol - (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow2, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4].Formula =
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPWorksheetLayout.ARMPResourcesCol - (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";
            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow3, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4].Formula = 
                "=SUM(r[0]c[1]:r[0]C[" + (ARMPWorksheetLayout.ARMPResourcesCol - (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4 - 1).ToString() + "])";



        }
        public void FormatARMPPlanning()
        {
            Application.DisplayAlerts = false;
            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range rngFormat;
            Excel.FormatCondition rngfmcCondition;

            DateTime ARMPStrtDate = ARMPWorksheetLayout.ARMPStrtDate;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowZ, ARMPWorksheetLayout.ARMPResourcesCol - 1]];
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Merge();
            rngFormat.Borders.Weight = Excel.XlBorderWeight.xlThick;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Font.Bold = true;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow1, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol1],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPSummRowsCnvt.SummRow3, (int)ARMPExcelLayout.ARMPSummColsCnvt.SummCol4]];
            rngFormat.NumberFormat = "0.00";
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Interior.Color = Color.LightCyan;

            int RsrcCols = (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt;
            do
            {
                rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols], 
                                                ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.ExcpDate, RsrcCols + ARMPWorksheetLayout.ARMPResources.Count - 1]];
                rngFormat.Merge();
                rngFormat.NumberFormat = "dd/mm/yyyy";
                rngFormat.Borders.Weight = Excel.XlBorderWeight.xlThick;

                rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, RsrcCols],
                                                ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, RsrcCols + ARMPWorksheetLayout.ARMPResources.Count - 1]];
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                RsrcCols = RsrcCols + ARMPWorksheetLayout.ARMPResources.Count;
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPWorksheetLayout.ARMPFnshDate) <= 0);

            for (int i = (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlce; i <= (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan; i++)
            {
                rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl, i],
                                                ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskTitl + 1, i]];
                rngFormat.Merge();
                rngFormat.Orientation = 90;
                rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;
            }

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPResourcesRowsCnvt.RsrcAmei, (int)ARMPExcelLayout.ARMPResourcesColsCnvt.RsrcStrt],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcExcd, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.Orientation = 90;
            rngFormat.HorizontalAlignment = Excel.Constants.xlCenter;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork, (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.NumberFormat = "0.00";
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt],
                                            ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcTodo, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
            rngfmcCondition.Font.Color = Color.Orange;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
            rngfmcCondition.Font.Color = Color.Green;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
            rngfmcCondition.Font.Color = Color.Red;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan]];
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            ARMPStrtDate = ARMPWorksheetLayout.ARMPStrtDate;
            int TaskCols = (int)ARMPExcelLayout.ARMPTasksColsCnvt.RsrcTime;
            do
            {
                rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt, TaskCols],
                                                ARMPWorksheet.Cells[(int)ARMPWorksheetLayout.ARMPTasksRowO, TaskCols + ARMPWorksheetLayout.ARMPResources.Count - 1]];
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                rngFormat.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                TaskCols = TaskCols + ARMPWorksheetLayout.ARMPResources.Count;
                ARMPStrtDate = ARMPStrtDate.AddDays(1);
            } while (ARMPStrtDate.CompareTo(ARMPWorksheetLayout.ARMPFnshDate) <= 0);

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkNorm],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.NumberFormat = "0.00";

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)ARMPExcelLayout.ARMPTasksRowsCnvt.TaskStrt, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkTodo],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkTodo]];
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=0");
            rngfmcCondition.Font.Color = Color.Orange;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=0");
            rngfmcCondition.Font.Color = Color.Green;
            rngfmcCondition = rngFormat.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=0");
            rngfmcCondition.Font.Color = Color.Red;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowA, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowZ, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.Interior.Color = Color.White;
            rngFormat.Font.Color = Color.Black;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowA, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowA, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Red;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowB, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowB, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Orange;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowC, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowC, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.Green;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowO, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Interior.Color = Color.LightBlue;
            rngFormat.Font.Color = Color.White;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowZ, 1],
                                            ARMPWorksheet.Cells[ARMPWorksheetLayout.ARMPTasksRowZ, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.FormatConditions.Delete();
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            rngFormat.Font.Color = Color.White;
            rngFormat.Interior.Color = Color.Blue;


            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrPrio].EntireColumn.ColumnWidth = 2;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.ColumnWidth = 12;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrNmbr].EntireColumn.Numberformat = "0";
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.ColumnWidth = 3;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperNmbr].EntireColumn.Numberformat = "0";
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.Numberformat = "dd/mm/jjjj";
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrStrt].EntireColumn.ColumnWidth = 10;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.Numberformat = "dd/mm/jjjj";
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.HorizontalAlignment = Excel.Constants.xlRight;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.RstrFnsh].EntireColumn.ColumnWidth = 10;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OrdrDesc].EntireColumn.ColumnWidth = 40;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.OperDesc].EntireColumn.ColumnWidth = 40;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkNorm].EntireColumn.ColumnWidth = 6;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkUnit].EntireColumn.ColumnWidth = 5;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkHour].EntireColumn.ColumnWidth = 6;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkReal].EntireColumn.ColumnWidth = 6;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkTodo].EntireColumn.ColumnWidth = 6;
            ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPTasksColsCnvt.WorkPlan].EntireColumn.ColumnWidth = 6;

            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[1, (int)ARMPExcelLayout.ARMPExceptionsColsCnvt.ExcpStrt], ARMPWorksheet.Cells[1, ARMPWorksheetLayout.ARMPExceptionsCol]];
            rngFormat.Columns.ColumnWidth = 5;
        }
    }
}
