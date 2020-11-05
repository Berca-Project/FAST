using System;
using System.Collections.Generic;

namespace Fast.Web.Models.Report
{
    public class ReportProdActualPlanningModel
    {
        public ReportProdActualPlanningModel()
        {
            ShiftReports = new List<ReportActualPlanningShiftModel>();
            ShiftReportBlocks = new List<ShiftReportBlock>();
        }

        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Shift { get; set; }
        public List<long> LocationIDList { get; set; }
        public string Granularity { get; set; }
        public string Product { get; set; }
        public string Type { get; set; }
        public long UserID { get; set; }
        public int Year { get; set; }
        public int Week { get; set; }
        public List<ReportActualPlanningShiftModel> ShiftReports { get; set; }
        public List<ShiftReportBlock> ShiftReportBlocks { get; set; }
    }

    public class ShiftReportBlock
    {
        public ReportActualPlanningShiftModel Plan { get; set; }
        public ReportActualPlanningShiftModel Actual { get; set; }
        public ReportActualPlanningShiftModel Total { get; set; }
    }

    public class ReportActualPlanningShiftModel
    {
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string UOM { get; set; }
        public List<ShiftModel> ShiftList { get; set; }
        public decimal Total { get; set; }
        public string TotalStr { get; set; }
        public string Location { get; set; }
        public long LocationID { get; set; }
        public string Remark { get; set; }

        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public string LU { get; set; }
        public string Market { get; set; }
        public decimal Actual { get; set; }
        public decimal Plan { get; set; }
    }

    public class ShiftModel
    {
        public DateTime Date { get; set; }
        public decimal Shift1 { get; set; }
        public decimal Shift2 { get; set; }
        public decimal Shift3 { get; set; }
        public decimal AllShift { get; set; }
        public string Shift1Str { get; set; }
        public string Shift2Str { get; set; }
        public string Shift3Str { get; set; }
        public string AllShiftStr { get; set; }
    }
}