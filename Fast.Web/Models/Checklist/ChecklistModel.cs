using System;
using System.Web;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistModel : BaseModel
    {
        public string CreatorEmployeeID { get; set; }
        public string MenuTitle { get; set; }
        public string Header { get; set; }
        public HttpPostedFileBase IconFile { get; set; }
        public string Icon { get; set; }
        public int ColumnHeader { get; set; }
        public int ColumnContent { get; set; }
        public int FrequencyAmount { get; set; }
        public int FrequencyDivider { get; set; }
        public string FrequencyUnit { get; set; }
        public List<long> Location1 { get; set; }
        public List<long> Location2 { get; set; }
        public List<long> Location3 { get; set; }
        public List<string> Components { get; set; }
        public int SubmitCount { get; set; }
        public SimpleEmployeeModel Creator { get; set; }
        public List<ChecklistLocationModel> Locations { get; set; }
        //public List<long> UserIDs { get; set; }
        public ChecklistModel()
        {
            Locations = new List<ChecklistLocationModel>();
        }
    }

    public class ChecklistEditModel : BaseModel
    {
        public ChecklistModel Checklist { get; set; }
        public List<ChecklistComponentModel> Components { get; set; }
    }

    public class SimpleEmployeeModel
    {
        public string EmployeeID { get; set; }
        public string FullName { get; set; }
    }

    public class ParentChilds
    {
        public LocationModel Parent { get; set; }
        public List<long> Childs { get; set; }
    }

    public class ChecklistReportAdherenceModel
    {
        public ChecklistModel Checklist { get; set; }
        public List<ChecklistSubmitModel> ReportItems { get; set; }
        public float TotalAdherence { get; set; }
    }

    public class ChecklistReportSubmitModel
    {
        public long CheklistSubmitID { get; set; }
        public string User { get; set; }
        public string Header { get; set; }
        public int Counter { get; set; }
        public int YesOn { get; set; }
        public string Date { get; set; }
        public DateTime Datetime { get; set; }
        
    }

    public class ChecklistReportSubmitGroupModel
    {
        public string Date { get; set; }
        public string User { get; set; }
        public string Header { get; set; }
        public int SubmitCount { get; set; }
        public int Counter { get; set; }
        public int YesOn { get; set; }
        public float Percentage { get; set; }
    }

    public class ChecklistReportRawModel
    {
        public string User { get; set; }
        public string Component { get; set; }
        public string Value { get; set; }
        public string Date { get; set; }
    }

    public class ChecklistReportGroupModel
    {
        public string Date { get; set; }
        public string User { get; set; }
        public string Component { get; set; }
        public int Counter { get; set; }
        public int YesOn { get; set; }
        public float Percentage { get; set; }
    }
    public class ChecklistChartModel
    {
        public string Date { get; set; }
        public List<double> Average { get; set; }
    }
    public class ChecklistReportModel
    {
        public ChecklistModel Checklist { get; set; }
        public List<string> Header { get; set; }
        public List<ChecklistReportSubmitGroupModel> ReportItems { get; set; }
        public List<ChecklistChartModel> Charts { get; set; }
    }

    public class ChecklistReportSummaryModel
    {
        public ChecklistModel Checklist { get; set; }
        public List<ChecklistSubmitModel> Submissions { get; set; }
        public List<EmployeeModel> Submitters { get; set; }
    }

    public class ChecklistReportRawDataModel
    {

    }

}