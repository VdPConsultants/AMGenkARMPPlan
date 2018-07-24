using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AMGenkARMPPlan
{
    public static class AMGenkResources
    {
        public static Object[,] GetAMGenkExceptions(string AMGenkResourcesDirectory, DateTime StrtDate, DateTime FnshDate)
        {
            System.Object xx = System.Type.Missing;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;

            string strAMGenkResourcesFile;

            Object[,] exceptions_1 = null;
            Object[,] exceptions_2 = null;
            Object[,] exceptions_3 = null;

            int ilastRowIgnoreFormulas = 0;

            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            strAMGenkResourcesFile = AMGenkResourcesDirectory + "\\" + "Aanwezigheden " + Globals.ThisAddIn.ARMPStrtDate.ToString("yyyy") + ".xlsm";

            xlWorkbook = xlApp.Workbooks.Open(strAMGenkResourcesFile, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);

            DateTime ARMPLastMnth = new DateTime(Globals.ThisAddIn.ARMPStrtDate.Year, Globals.ThisAddIn.ARMPStrtDate.Month, DateTime.DaysInMonth(Globals.ThisAddIn.ARMPStrtDate.Year, Globals.ThisAddIn.ARMPStrtDate.Month));

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[Globals.ThisAddIn.ARMPStrtDate.ToString("MMMM")];
            ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.WorkPlce],
                                                            xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.RsrcName]];
            exceptions_1 = (Object[,])exceptionsRange.Cells.Value2;

            if (Globals.ThisAddIn.ARMPStrtDate.Month == Globals.ThisAddIn.ARMPFnshDate.Month)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[Globals.ThisAddIn.ARMPStrtDate.ToString("MMMM")];
                ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + Globals.ThisAddIn.ARMPStrtDate.Day - 1],
                                                    xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + Globals.ThisAddIn.ARMPFnshDate.Day - 1]];
                exceptions_2 = (Object[,])exceptionsRange.Cells.Value2;
            }
            else
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[Globals.ThisAddIn.ARMPStrtDate.ToString("MMMM")];
                ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + Globals.ThisAddIn.ARMPStrtDate.Day - 1],
                                                    xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + ARMPLastMnth.Day - 1]];
                exceptions_2 = (Object[,])exceptionsRange.Cells.Value2;

                if (Globals.ThisAddIn.ARMPStrtDate.Year == Globals.ThisAddIn.ARMPFnshDate.Year)
                {
                    xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[Globals.ThisAddIn.ARMPFnshDate.ToString("MMMM")];
                    ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                    exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt],
                                                        xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + Globals.ThisAddIn.ARMPFnshDate.Day - 1]];
                    exceptions_3 = (Object[,])exceptionsRange.Cells.Value2;
                }
                else
                {
                    strAMGenkResourcesFile = AMGenkResourcesDirectory + "\\" + "Aanwezigheden " + Globals.ThisAddIn.ARMPStrtDate.ToString("yyyy") + ".xlsm";

                    xlWorkbook = xlApp.Workbooks.Open(strAMGenkResourcesFile, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);
                    xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[FnshDate.ToString("MMMM")];
                    ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                    exceptionsRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionsRowsOrig.ExcpStrt, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt],
                                                        xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionsColsOrig.ExcpStrt + FnshDate.Day - 1]];
                    exceptions_3 = (Object[,])exceptionsRange.Cells.Value2;
                }
            }
            Object[,] exceptions = ObjectMerge(exceptions_1, exceptions_2, exceptions_3);
            xlWorkbook.Close(0);
            xlApp.Quit();
            return exceptions;
        }
        public static Object[,] GetAMGenkExceptionCodes(string AMGenkResourcesDirectory, DateTime StrtDate, DateTime FnshDate)
        {
            System.Object xx = System.Type.Missing;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;

            string strAMGenkResourcesFile;

            int ilastRowIgnoreFormulas = 0;

            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            strAMGenkResourcesFile = AMGenkResourcesDirectory + "\\" + "Aanwezigheden " + Globals.ThisAddIn.ARMPStrtDate.ToString("yyyy") + ".xlsm";

            xlWorkbook = xlApp.Workbooks.Open(strAMGenkResourcesFile, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Codes"];
            ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range exceptioncodesRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPExceptionCodesRowsOrig.ClmnHead, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsOrig.ExcdType],
                                                                xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPExceptionCodesColsOrig.ExcdDayp]];
            Object[,] exceptioncodes = (Object[,])exceptioncodesRange.Cells.Value2;
            xlWorkbook.Close(0);
            xlApp.Quit();
            return exceptioncodes;
        }

        public static Object[,] GetAMGenkResources(string AMGenkResourcesDirectory, DateTime StrtDate, DateTime FnshDate)
        {
            System.Object xx = System.Type.Missing;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;

            string strAMGenkResourcesFile;

            int ilastRowIgnoreFormulas = 0;

            xlApp = new Excel.Application();
            // Don't interrupt with alert dialogs.
            xlApp.DisplayAlerts = false;

            strAMGenkResourcesFile = AMGenkResourcesDirectory + "\\" + "Aanwezigheden " + Globals.ThisAddIn.ARMPStrtDate.ToString("yyyy") + ".xlsm";

            xlWorkbook = xlApp.Workbooks.Open(strAMGenkResourcesFile, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx, xx);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets["Basisdata"];

            ilastRowIgnoreFormulas = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            Excel.Range resourcesRange = xlWorksheet.Range[xlWorksheet.Cells[(int)ARMPPlanExcelLayout.ARMPResourcesRowsOrig.RsrcYear, (int)ARMPPlanExcelLayout.ARMPResourcesColsOrig.WorkPlce],
                                                           xlWorksheet.Cells[ilastRowIgnoreFormulas, (int)ARMPPlanExcelLayout.ARMPResourcesColsOrig.RsrcAmei]];
            Object[,] resources = (Object[,])resourcesRange.Cells.Value2;
            xlWorkbook.Close(0);
            xlApp.Quit();
            return resources;
        }

        static Object[,] ObjectMerge(Object[,] r1, Object[,] r2, Object[,] r3)
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
        static object[,] NewObjectArray(int iRows, int iCols)
        {
            int[] aiLowerBounds = new int[] { 1, 1 };
            int[] aiLengths = new int[] { iRows, iCols };

            return (object[,])Array.CreateInstance(typeof(object), aiLengths, aiLowerBounds);
        }
    }
}
