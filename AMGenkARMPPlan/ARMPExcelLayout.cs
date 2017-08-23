using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    class ARMPExcelLayout
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
            RsrcAbbr = ExcpDate + 1,
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
            RsrcStrt = ARMPTasksColsCnvt.RsrcTime
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
            RsrcExcd = ARMPResourcesRowsCnvt.RsrcAbbr + 1,
            RsrcWork = RsrcExcd + 1,
            RsrcTodo = RsrcWork + 1,
            RsrcPlan = RsrcTodo + 1
        }
        public enum ARMPExceptionsColsOrig
        {
            WorkPlce = 1,
            RsrcName = WorkPlce + 1,
            StrtTime = RsrcName + 1,
            ExcpStrt = RsrcName + 1
        }
        public enum ARMPExceptionsColsImpr
        {
            WorkPlce = 1,
            RsrcName = 2,
            ExcpStrt = 3
        }
        public enum ARMPExceptionsColsCnvt
        {
            ExcpStrt = ARMPTasksColsCnvt.RsrcTime
        }
        public enum ARMPTasksRowsOrig
        {
            ClmnHead = 1,
            TaskRows = ClmnHead + 1
        }
        public enum ARMPTasksRowsCnvt
        {
            TaskStrt = ARMPExcelLayout.ARMPExceptionsRowsCnvt.RsrcPlan + 1
        }
        public enum ARMPTasksColsOrig
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1,
            OrdrStrt = OperNmbr + 1,
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
            WorkUnit = WorkNorm + 1
        }
        public enum ARMPTasksColsImpr
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1, 
            OrdrStrt = OperNmbr + 1,
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
            WorkUnit = WorkNorm + 1
        }
        public enum ARMPTasksColsCnvt
        {
            WorkPlce = 1,
            MainWork = WorkPlce + 1,
            OrdrPrio = MainWork + 1,
            OrdrNmbr = OrdrPrio + 1,
            OperNmbr = OrdrNmbr + 1,
            OrdrStrt = OperNmbr + 1,
            GateDate = OrdrStrt + 1,
            OrdrDesc = GateDate + 1,
            OperDesc = OrdrDesc + 1,
            WorkUnit = OperDesc + 1,
            WorkNorm = WorkUnit + 1,
            WorkHour = WorkNorm + 1,
            WorkTodo = WorkHour + 1,
            WorkPlan = WorkTodo + 1,
            RsrcTime = WorkPlan + 1
        }

        public int ARMPResourcesRow { get; set; }
        public int ARMPResourcesCol { get; set; }

        public int ARMPExceptionsRow { get; set; }
        public int ARMPExceptionsCol { get; set; }

        public int ARMPTasksRow { get; set; }
        public int ARMPTasksRowA { get; set; }
        public int ARMPTasksRowB { get; set; }
        public int ARMPTasksRowC { get; set; }
        public int ARMPTasksRowO { get; set; }

    public int ARMPTasksCol { get; set; }

        public DateTime ARMPStrtDate { get; set; }
        public DateTime ARMPFnshDate { get; set; }
    }
}
