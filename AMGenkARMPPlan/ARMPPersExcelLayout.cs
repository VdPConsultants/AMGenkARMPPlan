using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    public class ARMPPersExcelLayout
    {
        public int ARMPResourcesRow { get; set; }
        public int ARMPResourcesCol { get; set; }

        public int ARMPExceptionsRow { get; set; }
        public int ARMPExceptionsCol { get; set; }

        //public int ARMPTasksRow { get; set; }
        public int ARMPTasksRow1 { get; set; }
        public int ARMPTasksRow2 { get; set; }
        public int ARMPTasksRow3 { get; set; }
        public int ARMPTasksRow4 { get; set; }
        public int ARMPTasksRow5 { get; set; }
        public int ARMPTasksRow6 { get; set; }
        public int ARMPTasksRow7 { get; set; }
        public int ARMPTasksRow8 { get; set; }
        public int ARMPTasksRow9 { get; set; }

        public int ARMPTasksRow { get; set; }

        public int ARMPTasksCol { get; set; }

        public DateTime ARMPStrtDate { get; set; }
        public DateTime ARMPFnshDate { get; set; }

        public void CopyFromPlanLayout(ARMPPlanExcelLayout InPlanLayout)
        {
            ARMPResourcesRow = InPlanLayout.ARMPResourcesRow;
            ARMPResourcesCol = InPlanLayout.ARMPResourcesCol;

            ARMPExceptionsRow = InPlanLayout.ARMPExceptionsRow;
            ARMPExceptionsCol = InPlanLayout.ARMPExceptionsCol;

            ARMPTasksRow1 = InPlanLayout.ARMPTasksRow1;
            ARMPTasksRow2 = InPlanLayout.ARMPTasksRow2;
            ARMPTasksRow3 = InPlanLayout.ARMPTasksRow3;
            ARMPTasksRow4 = InPlanLayout.ARMPTasksRow4;
            ARMPTasksRow5 = InPlanLayout.ARMPTasksRow5;
            ARMPTasksRow = InPlanLayout.ARMPTasksRow;

            ARMPTasksCol = InPlanLayout.ARMPTasksCol;

            ARMPStrtDate = InPlanLayout.ARMPStrtDate;
            ARMPFnshDate = InPlanLayout.ARMPFnshDate;
        }
    }
}

