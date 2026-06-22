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
        public int ARMPTasksRow1 { get; set; }
        public int ARMPTasksRow2 { get; set; }
        public int ARMPTasksRow3 { get; set; }
        public int ARMPTasksRow4 { get; set; }
        public int ARMPTasksRow5 { get; set; }

        public int ARMPTasksRow { get; set; }

        public void CopyFromPlanLayout(ARMPPlanExcelLayout InPlanLayout)
        {
            ARMPTasksRow1 = InPlanLayout.ARMPTasksRow1 - InPlanLayout.ARMPTasksRow1 + 1;
            ARMPTasksRow2 = InPlanLayout.ARMPTasksRow2 - InPlanLayout.ARMPTasksRow1 + 1;
            ARMPTasksRow3 = InPlanLayout.ARMPTasksRow3 - InPlanLayout.ARMPTasksRow1 + 1;
            ARMPTasksRow4 = InPlanLayout.ARMPTasksRow4 - InPlanLayout.ARMPTasksRow1 + 1;
            ARMPTasksRow5 = InPlanLayout.ARMPTasksRow5 - InPlanLayout.ARMPTasksRow1 + 1;
            ARMPTasksRow = InPlanLayout.ARMPTasksRow - InPlanLayout.ARMPTasksRow1 + 1;
        }
    }
}
