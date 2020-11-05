using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web;

namespace Fast.Web.Models
{
    public class EmployeeOvertimeModel
    {
        public long ID { get; set; }
        public string EmployeeID { get; set; }
        public string FullName { get; set; }
        public string DepartmentDesc { get; set; }
        public string PositionDesc { get; set; }
        public string BasetownLocation { get; set; }
        public string CostCenter { get; set; }
        public DateTime Date { get; set; }
        public string ClockIn { get; set; }
        public string ClockOut { get; set; }
        public string ActualIn { get; set; }
        public string ActualOut { get; set; }
        public double Overtime { get; set; }
        public string OvertimeCategory { get; set; }
        public string Comments { get; set; }
        public AccessRightDBModel Access { get; set; }
        public double TotalOvertime { get; set; }
        public int TotalManPower { get; set; }
        public string YTDYear { get; set; }
        public string Month { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }
        public long ProdCenterID { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
    }

    public class OvertimeRootModel
    {
        public OvertimeRootModel()
        {
            Parents = new List<OvertimeParentModel>();
        }
        public List<OvertimeParentModel> Parents { get; set; }

    }
    public class OvertimeParentModel
    {
        public OvertimeParentModel()
        {
            Children = new List<OvertimeModel>();
        }

        public string YTDYear { get; set; }
        public string Month { get; set; }
        public double TotalPercentage { get; set; }
        public List<OvertimeModel> Children { get; set; }
    }

    public class OvertimeModel
    {
        public string OvertimeCategory { get; set; }
        public double TotalOvertime { get; set; }
        public int TotalManPower { get; set; }
        public double Percentage { get; set; }
        public string PercentageStr { get; set; }
    }

    public class OvertimeDashboardModel
    {
        public string LastMonthName { get; set; }
        public string LastTwoMonthName { get; set; }
        public string LastThreeMonthName { get; set; }
        public OvertimeData LastMonth { get; set; }
        public OvertimeData LastTwoMonth { get; set; }
        public OvertimeData LastThreeMonth { get; set; }

        public OvertimeDashboardModel()
        {
            LastMonth = new OvertimeData();
            LastTwoMonth = new OvertimeData();
            LastThreeMonth = new OvertimeData();
        }
    }

    public class OvertimeData
    {
        public double Blank { get; set; }
        public double Rework { get; set; }
        public double Other { get; set; }
        public double Daily { get; set; }
        public double Emergency { get; set; }
        public double BackupLeave { get; set; }
        public double Maintenance { get; set; }
        public double Leave { get; set; }
        public double Volume { get; set; }
        public double Training { get; set; }
        public double Project { get; set; }
    }
}