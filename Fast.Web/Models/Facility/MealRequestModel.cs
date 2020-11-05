using System;

namespace Fast.Web.Models
{
    public class MealRequestModel : BaseModel
    {
        public string CostCenter { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime Date { get; set; }
        public string ProductionCenterID { get; set; }
        public string ProductionCenter { get; set; }
        public string Canteen { get; set; }
        public string EmployeeID { get; set; }
        public string EmployeeFullname { get; set; }
        public int TotalGuest { get; set; }
        public string GuestType { get; set; }
        public string Guest { get; set; }
        public string Company { get; set; }
        public string Purpose { get; set; }
        public string Department { get; set; }
        public string Shift { get; set; }
        public int Shift1 { get; set; }
        public int Shift2 { get; set; }
        public int Shift3 { get; set; }
        public int NS { get; set; }
        public string Phone { get; set; }
        public string RequestType { get; set; }
        public string PIC { get; set; }
    }
}