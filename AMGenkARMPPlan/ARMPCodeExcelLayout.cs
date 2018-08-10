using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    public class ARMPCodeExcelLayout
    {
        public enum ARMPTasksRowsCnvt
        {
            TaskStrt = 1
        }
        public enum ARMPTasksColsCnvt
        {
            WorkPlce = 1,
            OrdrNmbr = WorkPlce + 1,
            OperNmbr = OrdrNmbr + 1,
            OrdrDesc = OperNmbr + 1,
            OperDesc = OrdrDesc + 1,
        };
        //public int ARMPTasksRow { get; set; }
        public int ARMPTasksRowA { get; set; }
        public int ARMPTasksRowB { get; set; }
        public int ARMPTasksRowC { get; set; }
        public int ARMPTasksRowO { get; set; }
        public int ARMPTasksRowZ { get; set; }

        public int ARMPTasksRow { get; set; }

        public void CopyFromPlanLayout(ARMPPlanExcelLayout InPlanLayout)
        {
            ARMPTasksRowA = InPlanLayout.ARMPTasksRowA - InPlanLayout.ARMPTasksRowA + 1;
            ARMPTasksRowB = InPlanLayout.ARMPTasksRowB - InPlanLayout.ARMPTasksRowA + 1;
            ARMPTasksRowC = InPlanLayout.ARMPTasksRowC - InPlanLayout.ARMPTasksRowA + 1;
            ARMPTasksRowO = InPlanLayout.ARMPTasksRowO - InPlanLayout.ARMPTasksRowA + 1;
            ARMPTasksRowZ = InPlanLayout.ARMPTasksRowZ - InPlanLayout.ARMPTasksRowA + 1;
            ARMPTasksRow = InPlanLayout.ARMPTasksRow - InPlanLayout.ARMPTasksRowA + 1;
        }
    }
}
