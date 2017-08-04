using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    class Constants
    {
        public enum ARMPTasksRows
        {
            Titl = 10,
            ARow = 11
        }
        public enum ARMPTasksCols
        {
            WorkPlce = 1,
            OrdrPrio = 2,
            OrdrNmbr = 3,
            OrdrStrt = 4,
            OrdrDesc = 7,
            WorkTime = 11,
            WorkUnit = 12
        }
        public enum ARMPTasksColsCnvt
        {
            WorkPlce = 1,
            OrdrPrio = 2,
            OrdrNmbr = 3,
            OrdrStrt = 4,
            OrdrDesc = 5,
            WorkTime = 6,
            WorkUnit = 7,
            WorkTodo = 9,
            WorkDone = 10,
            ExcpStrt = 11
        }
    }
}
