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

namespace AMGenkARMPPlan
{
    public partial class ThisAddIn
    {
        private Office.CommandBar commandBar;
        private Office.CommandBarButton importButton;

        private DataTable ARMPExceptionCodes = new DataTable();

        private int ARMPExceptionsFini = 0;

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
                // Add a button named ImportProject and an event handler.
                importButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                importButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                importButton.Caption = "Import Resources and Tasks";
                importButton.Tag = "import";
                importButton.TooltipText = "Import resources and tasks from Excel.";
                importButton.Click +=
                    new Office._CommandBarButtonEvents_ClickEventHandler(ImportButtonClick);

                commandBar.Visible = true;
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message, "Error adding toolbar button",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ImportButtonClick(Office.CommandBarButton ctrl, ref bool cancel)
        {
            Dialog dlgDialog = new Dialog();
            dlgDialog.Show();
        }
        public void CreateARMPExceptionCodes(Object[,] exceptioncodes)
        {
            // Convert the 2 dimensional object array in a typed datatable
            ARMPExceptionCodes.Columns.Add(new DataColumn("Code", Type.GetType("System.String")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("Afkorting", Type.GetType("System.String")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("Tijd", Type.GetType("System.TimeSpan")));
            ARMPExceptionCodes.Columns.Add(new DataColumn("Dagdeel", Type.GetType("System.String")));


            for (int i = 1; i < exceptioncodes.GetLength(0) + 1; i++)
            {
                if (exceptioncodes[i, 1] != null)
                {
                    DataRow ARMPExceptionCodesRow = ARMPExceptionCodes.NewRow();
                    ARMPExceptionCodesRow["Code"] = exceptioncodes[i, 1].ToString();
                    ARMPExceptionCodesRow["Afkorting"] = exceptioncodes[i, 2].ToString();
                    ARMPExceptionCodesRow["Tijd"] = TimeSpan.Parse(DateTime.FromOADate((double)exceptioncodes[i, 3]).ToString("HH:mm:ss"));
                    ARMPExceptionCodesRow["Dagdeel"] = exceptioncodes[i, 4].ToString();
                    ARMPExceptionCodes.Rows.Add(ARMPExceptionCodesRow);
                }
            }
        }

        public void CreateARMPResources(Object[,] resources)
        {
            // TODO: Introducing break - JvdP 20170727
            // TODO: Variable starting hours - JvdP 20170727
            // DateTime resourceStartTime = DateTime.FromOADate((double)resourcesCells[i, 2]);
            DateTime resourceStartTime = DateTime.Parse("08:00:00");
            // DateTime resourceFinishTime = DateTime.Parse("17:00:00") - (DateTime.Parse("09:00:00") - resourceStartTime);
            DateTime resourceFinishTime = DateTime.Parse("16:00:00");

            // TODO: hardcoded - JvdP 201700804
            int ARMPResourcesRow = 2;
            int ARMPResourcesCol = 24;

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            ARMPWorksheet.Cells[1, 24].Value2 = "Maandag 07 augustus 2017"; ;

            // VALUES
            // To test general Exception handler, use the upper bound: i <= taskList.Count
            for (int i = 1; i < resources.GetLength(0) + 1; i++)
            {
                ARMPWorksheet.Cells[ARMPResourcesRow, ARMPResourcesCol].Value2 = resources[i, 2].ToString();
                ARMPResourcesCol += 2;
            }
            // FORMAT
            // TODO: hardcoded - JvdP 201700804
            for (int i = 24; i < ARMPResourcesCol; i += 2)
            {
                Excel.Range rngARMPResource = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPResourcesRow, i], ARMPWorksheet.Cells[ARMPResourcesRow, i + 1]];
                rngARMPResource.MergeCells = true;
                rngARMPResource.Orientation = 90;
                rngARMPResource.HorizontalAlignment = Excel.Constants.xlCenter;
            }
        }
        public void CreateARMPExceptions(Object[,] exceptions)
        {
            int ARMPExceptionsRow = 3;
            int ARMPExceptionsCol = (int)Constants.ARMPTasksColsCnvt.ExcpStrt;

            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            // VALUES
            for (int i = 1; i < exceptions.GetLength(0) + 1; i++)
            //for (int i = 1; i <= 1; i++)
            {
                // Set the resource exceptions 
                //for (int j = 1; j < exceptions.GetLength(1); j++)
                for (int j = 1; j <= 1; j++)
                {
                    string[] ARMPExceptionList;
                    string[] separators = { "/" };

                    string ARMPCode = "";
                    TimeSpan ARMPWorkTime = TimeSpan.Parse("08:00:00");

                    TimeSpan TijdV, TijdL, TijdD;

                    DateTime ARMPExceptionDate = new DateTime(2017, 8, j);
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
                                    DataRow test = (from myrow in ARMPExceptionCodes.AsEnumerable()
                                                    where myrow.Field<string>("Code") == ARMPException
                                                    select myrow).SingleOrDefault();
                                    switch (test["Dagdeel"].ToString())
                                    {
                                        case "D":
                                            TijdD = TijdD.Add(TimeSpan.Parse("08:00:00"));
                                            break;
                                        case "V":
                                            TijdV = TijdV.Add((TimeSpan)test["Tijd"]);
                                            break;
                                        case "L":
                                            TijdL = TijdL.Add((TimeSpan)test["Tijd"]);
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
                                                where myrow.Field<string>("Code") == ARMPExceptionList[0].ToString()
                                                select myrow).SingleOrDefault();
                                ARMPCode = test["Code"].ToString();
                                ARMPWorkTime = TimeSpan.Parse("08:00:00").Subtract((TimeSpan)test["Tijd"]);
                            }
                            catch
                            {
                                // Code not known - skip
                            }
                        }
                    }

                    ARMPWorksheet.Cells[ARMPExceptionsRow, ARMPExceptionsCol].Value2 = ARMPCode;
                    //ARMPWorksheet.Cells[ARMPExceptionsRow + 1, ARMPExceptionsCol].Value2 = ARMPWorkTime.ToString();
                    ARMPWorksheet.Cells[ARMPExceptionsRow + 2, ARMPExceptionsCol].Formula = "=SUM(r[0]c[1]:r[1]C[1])";
                    ARMPWorksheet.Cells[ARMPExceptionsRow + 2, ARMPExceptionsCol].FormulaHidden = true;
                    ARMPWorksheet.Cells[ARMPExceptionsRow + 2, ARMPExceptionsCol].Calculate();

                    ARMPWorksheet.Cells[ARMPExceptionsRow + 2, ARMPExceptionsCol + 1].Value2 = ARMPWorkTime.TotalHours;

                    ARMPExceptionsCol += 2;
                }
                ARMPExceptionsFini = ARMPExceptionsCol - 2;
            }
            // FORMAT
            // TODO: hardcoded - JvdP 201700804
            Excel.Range rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[ARMPExceptionsRow + 2, 24], ARMPWorksheet.Cells[ARMPExceptionsRow + 2, ARMPExceptionsCol]];
            rngFormat.NumberFormat = "0.00" ;

            //ARMPWorksheet.Columns()
        }

        public void CreateARMPTasks(Object[,] tasks)
        {

            int ARMPTasksRow = (int)Constants.ARMPTasksRows.Titl;
            int ARMPTasksRowA = ARMPTasksRow + 1;
            int ARMPTasksRowB = ARMPTasksRow + 2;
            int ARMPTasksRowC = ARMPTasksRow + 3;
            int ARMPTasksRowO = ARMPTasksRow + 4;

            string ARMPWorkDoneForm = "=SUM(r[0]c[1]:r[0]C[" + (ARMPExceptionsFini - Constants.ARMPTasksColsCnvt.ExcpStrt).ToString() + "])";

            // VALUES
            Excel.Worksheet ARMPWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            ARMPWorksheet.Cells[ARMPTasksRowA, 1].Value2 = "Taken met prioriteit A";
            ARMPWorksheet.Cells[ARMPTasksRowB, 1].Value2 = "Taken met prioriteit B";
            ARMPWorksheet.Cells[ARMPTasksRowC, 1].Value2 = "Taken met prioriteit C";
            ARMPWorksheet.Cells[ARMPTasksRowO, 1].Value2 = "Taken met prioriteit andere";
            for (int i = 1; i < tasks.GetLength(0) + 1; i++)
            {
                // Priority A tasks
                switch (tasks[i, (int)Constants.ARMPTasksCols.OrdrPrio].ToString())
                {
                    case "A":
                        ARMPWorksheet.Cells[ARMPTasksRowA + 1, 1].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        ARMPTasksRowA++;
                        ARMPTasksRowB++;
                        ARMPTasksRowC++;
                        ARMPTasksRowO++;
                        ARMPTasksRow = ARMPTasksRowA;
                        break;

                    case "B":
                        ARMPWorksheet.Cells[ARMPTasksRowB + 1, 1].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        ARMPTasksRowB++;
                        ARMPTasksRowC++;
                        ARMPTasksRowO++;
                        ARMPTasksRow = ARMPTasksRowB;
                        break;

                    case "C":
                        ARMPWorksheet.Cells[ARMPTasksRowC + 1, 1].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        ARMPTasksRowC++;
                        ARMPTasksRowO++;
                        ARMPTasksRow = ARMPTasksRowC;
                        break;

                    default:
                        ARMPTasksRowO++;
                        ARMPTasksRow = ARMPTasksRowO;
                        break;
                }
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.WorkPlce].Value2 = tasks[i, (int)Constants.ARMPTasksCols.WorkPlce].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.OrdrPrio].Value2 = tasks[i, (int)Constants.ARMPTasksCols.OrdrPrio].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.OrdrNmbr].Value2 = tasks[i, (int)Constants.ARMPTasksCols.OrdrNmbr].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.OrdrStrt].Value2 = tasks[i, (int)Constants.ARMPTasksCols.OrdrStrt].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.OrdrDesc].Value2 = tasks[i, (int)Constants.ARMPTasksCols.OrdrDesc].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.WorkTime].Value2 = tasks[i, (int)Constants.ARMPTasksCols.WorkTime].ToString();
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.WorkUnit].Value2 = tasks[i, (int)Constants.ARMPTasksCols.WorkUnit].ToString();

                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.WorkTodo].Value2 = Conversions.TimeUnit2Todo(tasks[i, (int)Constants.ARMPTasksCols.WorkTime].ToString(),
                                                                                                                                tasks[i, (int)Constants.ARMPTasksCols.WorkUnit].ToString());
                ARMPWorksheet.Cells[ARMPTasksRow, (int)Constants.ARMPTasksColsCnvt.WorkDone].Formula = ARMPWorkDoneForm;
            }
            // FORMATS
            Excel.Range rngFormat;
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)Constants.ARMPTasksRows.Titl + 1, (int)Constants.ARMPTasksColsCnvt.WorkTime], ARMPWorksheet.Cells[ARMPTasksRowO, (int)Constants.ARMPTasksColsCnvt.WorkTime]];
            rngFormat.NumberFormat = "0.00";
            rngFormat = ARMPWorksheet.Range[ARMPWorksheet.Cells[(int)Constants.ARMPTasksRows.Titl + 1, (int)Constants.ARMPTasksColsCnvt.WorkTodo], ARMPWorksheet.Cells[ARMPTasksRowO, ARMPExceptionsFini]];
            rngFormat.NumberFormat = "0.00";
        }
    }
}
