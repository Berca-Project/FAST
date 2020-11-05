using System;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class MachineModel : BaseModel
    {
        public string Code { get; set; }
        public string Items { get; set; }
        public string LegalEntity { get; set; }
        public string MachineBrand { get; set; }
        public long MachineTypeID { get; set; }
        public string MachineType { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }
        public string MachineSN { get; set; }
        public string SubProcess { get; set; }
        public string SamID { get; set; }
        public string Notes { get; set; }
        public string LinkUp { get; set; }
        public string Type { get; set; }
        public string Cluster { get; set; }
        public Nullable<int> OrderNumber { get; set; }
        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
        public Nullable<decimal> DesignSpeed { get; set; }
        public Nullable<decimal> CellophanerSpeed { get; set; }
        public bool IsShift1FullyAssigned { get; set; }
        public bool IsShift2FullyAssigned { get; set; }
        public bool IsShift3FullyAssigned { get; set; }
        public bool IsShift1Zero { get; set; }
        public bool IsShift2Zero { get; set; }
        public bool IsShift3Zero { get; set; }
        public bool IsMaker { get; set; }
        public bool IsPacker { get; set; }
        public bool IsSP { get; set; }
        public bool IsPP { get; set; }
        public bool IsOther { get; set; }
        public bool IsExistInWpp { get; set; }
        public bool IsSelectedDate { get; set; }
    }

    public class MppMachineModel
    {
        public List<MachineModel> MachineList { get; set; }
        public int CurrentDateMachineCount { get; set; }
        public int NextDateMachineCount { get; set; }
        public string CurrentDate { get; set; }
        public string NextDate { get; set; }
        public string SelectedDateShift1Percentage { get; set; }
        public string SelectedDateShift2Percentage { get; set; }
        public string SelectedDateShift3Percentage { get; set; }
        public string NextDateShift1Percentage { get; set; }
        public string NextDateShift2Percentage { get; set; }
        public string NextDateShift3Percentage { get; set; }
        public string NextDateOverallPercentage { get; set; }
        public string SelectedDateOverallPercentage { get; set; }
    }
}