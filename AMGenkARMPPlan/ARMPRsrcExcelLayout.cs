using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    public class ARMPRsrcExcelLayout
    {
        public int ARMPResourcesRow { get; set; }
        public int ARMPResourcesCol { get; set; }

        public int ARMPExceptionsRow { get; set; }
        public int ARMPExceptionsCol { get; set; }

        //public int ARMPTasksRow { get; set; }
        public int ARMPTasksRowA { get; set; }
        public int ARMPTasksRowB { get; set; }
        public int ARMPTasksRowC { get; set; }
        public int ARMPTasksRowO { get; set; }
        public int ARMPTasksRowZ { get; set; }

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

            ARMPTasksRowA = InPlanLayout.ARMPTasksRowA;
            ARMPTasksRowB = InPlanLayout.ARMPTasksRowB;
            ARMPTasksRowC = InPlanLayout.ARMPTasksRowC;
            ARMPTasksRowO = InPlanLayout.ARMPTasksRowO;
            ARMPTasksRowZ = InPlanLayout.ARMPTasksRowZ;
            ARMPTasksRow = InPlanLayout.ARMPTasksRow;

            ARMPTasksCol = InPlanLayout.ARMPTasksCol;

            ARMPStrtDate = InPlanLayout.ARMPStrtDate;
            ARMPFnshDate = InPlanLayout.ARMPFnshDate;
        }
    }
}

