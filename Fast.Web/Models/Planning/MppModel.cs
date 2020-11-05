using System;
using System.Collections.Generic;
using System.Web;

namespace Fast.Web.Models
{
    public class MppModel : BaseModel
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int Year { get; set; }
        public int Week { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public string StatusMPP { get; set; }
        public string JobTitle { get; set; }
        public string RoleName { get; set; }
        public string EmployeeID { get; set; }
        public string EmployeeIDLS1 { get; set; }
        public string EmployeeIDLS2 { get; set; }
        public string EmployeeName { get; set; }
        public string EmployeeMachine { get; set; }
        public long WPPID { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public string Remark { get; set; }
        public string GroupType { get; set; }
        public string GroupName { get; set; }

        // custom fields
        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
        public long JobTitleID { get; set; }
        public long GroupTypeID { get; set; }
        public Nullable<DateTime> StartDateCalendar { get; set; }
        public Nullable<DateTime> EndDateCalendar { get; set; }
        public string IDNormal { get; set; }
        public long IDLS1 { get; set; }
        public long IDLS2 { get; set; }
        public bool IsWPPExist { get; set; }
        public bool IsHoliday { get; set; }

        public long PcID { get; set; }
        public long DepID { get; set; }
        public long SubDepID { get; set; }
        public string OvertimeDate { get; set; }

        public int SumAvailable { get; set; }
        public int SumHalfAssigned { get; set; }
        public int SumFullyAssigned { get; set; }
        public int SumOverload { get; set; }

        public long LocID { get; set; }
        public string LocType { get; set; }
        public string Canteen { get; set; }
        public string CostCenter { get; set; }
    }

    public class MppDashboardModel
    {
        public int Total { get; set; }
        public int Available { get; set; }
        public int HalfAssigned { get; set; }
        public int FullyAssigned { get; set; }
        public int Overload { get; set; }
        public int Allocated { get; set; }
        public int Unallocated { get; set; }
        public int Unused { get; set; }
        public int NotExistInWPP { get; set; }
    }

    public class MppSummaryModel
    {
        public DateTime Date { get; set; }
        public string JobTitle { get; set; }
        public int Total { get; set; }
        public int Assigned { get; set; }
        public int Idle { get; set; }
    }

    public class MppAllocationModel
    {
        public List<MppModel> MppList { get; set; }
        public List<MachineAllocationModel> AllocationList { get; set; }
    }

    public class MPPMiniModel
    {
        public string Group { get; set; }
        public string JobTitle { get; set; }
        public string EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpMachine { get; set; }
        public long EmpMachineLocationID { get; set; }
    }

    public class MPPExcelModel
    {
        public string JobTitle { get; set; }
        public string EmpIdA { get; set; }
        public string EmpNameA { get; set; }
        public string EmpMachineA { get; set; }
        public long EmpMachineALocationID { get; set; }
        public string EmpIdB { get; set; }
        public string EmpNameB { get; set; }
        public string EmpMachineB { get; set; }
        public long EmpMachineBLocationID { get; set; }
        public string EmpIdC { get; set; }
        public string EmpNameC { get; set; }
        public string EmpMachineC { get; set; }
        public long EmpMachineCLocationID { get; set; }
        public string EmpIdD { get; set; }
        public string EmpNameD { get; set; }
        public string EmpMachineD { get; set; }
        public long EmpMachineDLocationID { get; set; }
    }

    public class MPPOvertimeModel
    {
        public List<string> CategoryList { get; set; }
        public List<string> MonthList { get; set; }
        public string YearYTD { get; set; }
        public List<OvertimeParentModel> OvertimeList { get; set; }
    }
}

