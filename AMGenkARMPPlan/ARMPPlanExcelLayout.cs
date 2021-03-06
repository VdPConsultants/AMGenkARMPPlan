﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    public class ARMPPlanExcelLayout
    {
        public enum ARMPExceptionCodesRowsOrig
        {
            ClmnHead = 1,
            ExcdStrt = ClmnHead + 1
        }
        public enum ARMPExceptionCodesRowsImpr
        {
            ClmnHead = 1,
            ExcdStrt = ClmnHead + 1
        }
        public enum ARMPExceptionCodesColsOrig
        {
            ExcdType = 1,
            ExcdCode = ExcdType + 1,
            ExcdAbbr = ExcdCode + 1,
            ExcdTime = ExcdAbbr + 1,
            ExcdDayp = ExcdTime + 1
        }
        public enum ARMPExceptionCodesColsImpr
        {
            ExcdType = 1,
            ExcdCode = ExcdType + 1,
            ExcdAbbr = ExcdCode + 1,
            ExcdTime = ExcdAbbr + 1,
            ExcdDayp = ExcdTime + 1
        }
        public enum ARMPResourcesRowsOrig
        {
            RsrcYear = 1,
            ClmnHead = RsrcYear + 1,
            RsrcStrt = ClmnHead + 1
        }
        public enum ARMPResourcesRowsImpr
        {
            RsrcYear = 1,
            ClmnHead = RsrcYear + 1,
            RsrcStrt = ClmnHead + 1
        }
        public enum ARMPResourcesRowsCnvt
        {
            ExcpDate = 1,
            RsrcAmei = ExcpDate + 1,
            RsrcName = RsrcAmei + 1
        }
        public enum ARMPResourcesColsOrig
        {
            WorkPlce = 1,
            RsrcName = WorkPlce + 1,
            RsrcAbbr = RsrcName + 1,
            RsrcAmei = RsrcAbbr + 1
        }
        public enum ARMPResourcesColsImpr
        {
            WorkPlce = 1,
            RsrcName = WorkPlce + 1,
            RsrcAbbr = RsrcName + 1,
            RsrcAmei = RsrcAbbr + 1
        }
        public enum ARMPResourcesColsCnvt
        {
            RsrcStrt = ARMPTasksColsCnvt.WorkPlan + 1
        }
        public enum ARMPExceptionsRowsOrig
        {
            ExcpStrt = 3
        }
        public enum ARMPExceptionsRowsImpr
        {
            ExcpStrt = 1
        }
        public enum ARMPExceptionsRowsCnvt
        {
            RsrcExcd = ARMPResourcesRowsCnvt.RsrcName + 1,
            RsrcWork = RsrcExcd + 1,
            RsrcTodo = RsrcWork + 1,
            RsrcPlan = RsrcTodo + 1
        }
        public enum ARMPExceptionsColsOrig
        {
            WorkPlce = 1,
            RsrcName = WorkPlce + 1,
            StrtTime = RsrcName + 1,
            ExcpStrt = StrtTime + 1
        }
        public enum ARMPExceptionsColsImpr
        {
            WorkPlce = 1,
            RsrcName = 2,
            ExcpStrt = 3
        }
        public enum ARMPExceptionsColsCnvt
        {
            ExcpStrt = ARMPTasksColsCnvt.WorkPlan + 1
        }
        public enum ARMPTasksRowsOrig
        {
            ClmnHead = 1,
            TaskRows = ClmnHead + 1
        }
        public enum ARMPTasksRowsImpr
        {
            TaskRows = 1
        }
        public enum ARMPTasksRowsCnvt
        {
            TaskTitl = 2,
            TaskStrt = ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan + 1
        }
        public enum ARMPTasksColsOrig
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1,
            RstrStrt = OperNmbr + 1,
            RstrFnsh = RstrStrt + 1,
            OrdrStrt = RstrFnsh + 1,
            GateDate = OrdrStrt + 1,
            OrdrDesc = GateDate + 1,
            OperDesc = OrdrDesc + 1,
            StatTech = OperDesc + 1,
            OrdrType = StatTech + 1,
            UserStat = OrdrType + 1,
            RevsDesc = UserStat + 1,
            DuraNmbr = RevsDesc + 1,
            DuraNorm = DuraNmbr + 1,
            DuraUnit = DuraNorm + 1,
            WorkNorm = DuraUnit + 1,
            WorkUnit = WorkNorm + 1,
            WorkReal = WorkUnit + 1
        }
        public enum ARMPTasksColsImpr
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1,
            RstrStrt = OperNmbr + 1,
            RstrFnsh = RstrStrt + 1,
            OrdrStrt = RstrFnsh + 1,
            GateDate = OrdrStrt + 1,
            OrdrDesc = GateDate + 1,
            OperDesc = OrdrDesc + 1,
            StatTech = OperDesc + 1,
            OrdrType = StatTech + 1,
            UserStat = OrdrType + 1,
            RevsDesc = UserStat + 1,
            DuraNmbr = RevsDesc + 1,
            DuraNorm = DuraNmbr + 1,
            DuraUnit = DuraNorm + 1,
            WorkNorm = DuraUnit + 1,
            WorkUnit = WorkNorm + 1,
            WorkReal = WorkUnit + 1
        }
        public enum ARMPTasksColsCnvt
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1,
            RstrStrt = OperNmbr + 1,
            RstrFnsh = RstrStrt + 1,
            OrdrDesc = RstrFnsh + 1,
            OperDesc = OrdrDesc + 1,
            WorkUnit = OperDesc + 1,
            WorkNorm = WorkUnit + 1,
            WorkHour = WorkNorm + 1,
            WorkReal = WorkHour + 1,
            WorkTodo = WorkReal + 1,
            WorkPlan = WorkTodo + 1
        };

        public string[] ARMPTasksColsHead =
        {
            "Uitvoerende werkplek",
            "Verantwoordelijke werkplek",
            "Prioriteit",
            "Order nummer",
            "Operatie nummer",
            "Start datum",
            "Eind datum",
            "Order beschrijving",
            "Operatie beschrijving",
            "Werktijd eenheid",
            "Werktijd",
            "Werktijd in uur",
            "Werktijd gewerkt",
            "Werktijd todo",
            "Werktijd gepland"
        };

        public enum ARMPSummRowsCnvt
        {
            SummRow1 = ARMPPlanExcelLayout.ARMPExceptionsRowsCnvt.RsrcWork,
            SummRow2 = SummRow1 + 1,
            SummRow3 = SummRow2 + 1
        }

        public enum ARMPSummColsCnvt
        {
            SummCol1 = ARMPPlanExcelLayout.ARMPTasksColsCnvt.WorkHour,
            SummCol2 = SummCol1 + 1,
            SummCol3 = SummCol2 + 1,
            SummCol4 = SummCol3 + 1
        }

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

        // Holds the workplaces which are in planned orders
        public List<string> ARMPWorkplaces = new List<string>();
        // Holds the resources which are in planned workplaces
        public List<Resource> ARMPResources = new List<Resource>();
        public List<Resource> ARMPResourcesFiltered = new List<Resource>();
    }
}
