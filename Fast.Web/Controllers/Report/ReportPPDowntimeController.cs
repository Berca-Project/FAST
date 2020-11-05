using Fast.Application.Interfaces;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using System;
using System.Web.Mvc;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.Script.Serialization;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Fast.Web.Resources;
using Fast.Web.Models.Report;
using Fast.Web.Models;
using System.Linq;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("ReportPPDowntime")]
    public class ReportPPDowntimeController : BaseController<PPLPHModel>
    {      
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly IWeeksAppService _weeksAppService;
        private readonly IUserAppService _userAppService;
        private readonly IUserRoleAppService _userRoleAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly IJobTitleAppService _jobTitleAppService;

        public ReportPPDowntimeController(
       
         ILoggerAppService logger,
         IReferenceAppService referenceAppService,
         ILocationAppService locationAppService,
         IEmployeeAppService employeeAppService,
         IWeeksAppService weeksAppService,
         IUserAppService userAppService,
         IUserRoleAppService userRoleAppService,
         IMachineAppService machineAppService,
         IReferenceDetailAppService referenceDetailAppService,
         IJobTitleAppService jobTitleAppService)
        {           
            _logger = logger;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _employeeAppService = employeeAppService;
            _weeksAppService = weeksAppService;
            _userAppService = userAppService;
            _userRoleAppService = userRoleAppService;
            _machineAppService = machineAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _jobTitleAppService = jobTitleAppService;
        }
        public ActionResult Index()
        {
            //LoadLeader();
            GetTempData();
            GetFilterTempData();
            ViewBag.isWest = 0;
            if (AccountProdCenterID == 4 || AccountProdCenterID == 5)
            {
                ViewBag.isWest = 1;
            }
            ViewBag.empID = AccountEmployeeID;
            ViewBag.isSupervisor = isSupervisor();
            return View();
        }
        private bool isSupervisor()
        {
            return AccountRoleList.Contains("SUPERVISOR") || AccountRoleList.Contains("Supervisor");
        }
        
        public ActionResult Edit(long id, string table)
        {
            //LoadLeader();
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            var myQuery = @"SELECT * FROM " + table + " WHERE ID = " + id + "";

            DataSet dset = new DataSet();
            using (SqlConnection con = new SqlConnection(strConString))
            {
                SqlCommand cmd = new SqlCommand(myQuery, con);
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dset);
                }
            }

            string data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]);
            List<PPReportDowntimeModel> dataModel = data.DeserializeToPPReportDowntimeModelList();// new PPReportDowntimeModel();
            //dataModel.FilterLine = table;
            //dataModel.LineProd = ;
            PPReportDowntimeModel finalData = dataModel[0];

            DateTime startDt = finalData.StartDateTime.ToLocalTime();
            DateTime endDt = finalData.EndDateTime.ToLocalTime();

            finalData.FilterLine = table;
            finalData.StartDate = startDt.Date;
            finalData.StartTime = startDt.ToString("HH:mm");
            finalData.EndDate = endDt.Date;
            finalData.EndTime = endDt.ToString("HH:mm");

            return PartialView(finalData);
        }
        public ActionResult Delete(long id, string table)
        {
            try
            {
                TempData["FilterFound"] = true;
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;

                #region Data
                var myQuerySelect = @"SELECT * FROM " + table + " WHERE ID = " + id + "";

                DataSet dsetSelect = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuerySelect, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dsetSelect);
                    }
                }

                string data = DataTableToJSONWithJavaScriptSerializer(dsetSelect.Tables[0]);
                List<PPReportDowntimeModel> dataModel = data.DeserializeToPPReportDowntimeModelList();
                PPReportDowntimeModel resultData = dataModel[0];


                #endregion

                var myQuery = string.Format("DELETE FROM {0} WHERE ID = {1};",table, id);
                var myQueryInsert = string.Format("INSERT INTO DTDeleted (PK,BatchID,TableName,ModifiedBy,ModifiedDate) values ({0},{1},'{2}','{3}','{4}')", resultData.PK, resultData.BatchID, table, GetFullNameByEmployeeId(AccountEmployeeID), DateTime.Now);

                DataSet dset = new DataSet();
                DataSet dset2= new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd2 = new SqlCommand(myQueryInsert, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd2))
                    {
                        da.Fill(dset2);
                    }
                }
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult Edit(PPReportDowntimeModel data)
        {
            try
            {
                TempData["FilterFound"] = true;
                TimeSpan tsStart = TimeSpan.Parse(data.StartTime + ":00");
                TimeSpan tsEnd = TimeSpan.Parse(data.EndTime + ":00");
                DateTime startDateTime = data.StartDate.AddSeconds(tsStart.TotalSeconds);
                DateTime endDateTime = data.EndDate.AddSeconds(tsEnd.TotalSeconds);
                TimeSpan duration = endDateTime - startDateTime;

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;

                var myQuery = string.Format("UPDATE {0} Set StartDateTime = '{1}',EndDateTime ='{2}',Duration={3},LogUser='{4}',LineProd='{5}',PlanUnplan='{6}',Category='{7}',UpDownStream='{8}',Issue='',Description='{9}',Remark='{10}',Status='{11}' where id = {12}; ", data.FilterLine, startDateTime, endDateTime, duration.TotalMinutes, GetFullNameByEmployeeId(AccountEmployeeID), data.LineProd, data.PlanUnplan, data.Category, data.UpDownStream, data.Description, data.Remark, data.Status, data.ID);

                if (isSupervisor())
                {
                    //myQuery = @"UPDATE " + data.FilterLine + " Set StartDateTime = '" + startDateTime + "',EndDateTime ='" + endDateTime + "',Duration=" + duration.TotalMinutes + ",ApprovedBy='" + GetFullNameByEmployeeId(AccountEmployeeID) + "',ApproverEmployeeID='" + AccountEmployeeID + "',LineProd='" + data.LineProd + "',PlanUnplan='" + data.PlanUnplan + "',Category='" + data.Category + "',UpDownStream='" + data.UpDownStream + "',Issue='',Description='" + data.Description + "',Remark='" + data.Remark + "',Status='" + data.Status + "' where " +
                    //            "id = " + data.ID;
                    myQuery = string.Format("UPDATE {0} Set StartDateTime = '{1}',EndDateTime ='{2}',Duration={3},LineProd='{4}',PlanUnplan='{5}',Category='{6}',UpDownStream='{7}',Issue='',Description='{8}',Remark='{9}',Status='{10}',ApprovedBy='{11}',ApproverEmployeeID='{12}' where id = {13}; ", data.FilterLine, startDateTime, endDateTime, duration.TotalMinutes, data.LineProd, data.PlanUnplan, data.Category, data.UpDownStream, data.Description, data.Remark, data.Status, GetFullNameByEmployeeId(AccountEmployeeID), AccountEmployeeID, data.ID);
                }

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }
                SetTrueTempData(UIResources.UpdateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UpdateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult AddNew(PPReportDowntimeModel data)
        {
            try
            {
                TempData["FilterFound"] = true;
                TimeSpan tsStart = TimeSpan.Parse(data.StartTime+":00");
                TimeSpan tsEnd = TimeSpan.Parse(data.EndTime + ":00");
                DateTime startDateTime = data.StartDate.AddSeconds(tsStart.TotalSeconds);
                DateTime endDateTime = data.EndDate.AddSeconds(tsEnd.TotalSeconds);
                TimeSpan duration = endDateTime - startDateTime;

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                //var myQuery = @"Insert Into " + data.FilterLine + "(PK,BatchID,StartDateTime,EndDateTime,Duration,LogUser,SubmitDateTime,LineProd,PlanUnplan,Category,UpDownStream,Issue,Description,Remark,Status,UserEmployeeID,ApproverEmployeeID,LocationID,Location) values" +
                //                "(0,"+data.BatchID+",'"+ startDateTime+ "','"+endDateTime+ "',"+duration.TotalMinutes+",'"+ GetFullNameByEmployeeId(AccountEmployeeID)+ "','"+DateTime.Now+"','"+data.LineProd+"','"+data.PlanUnplan+"','"+data.Category+"','"+data.UpDownStream+"','"+data.Issue+"','"+data.Description+ "','" +data.Remark +"','" +data.Status +"'," + AccountEmployeeID+"," + data.ApproverEmployeeID+"," +AccountDepartmentID +",'" + AccountLocation.Substring(0, 5) +"')";
                var myQuery = @"Insert Into " + data.FilterLine + "(PK,BatchID,StartDateTime,EndDateTime,Duration,LogUser,SubmitDateTime,LineProd,PlanUnplan,Category,UpDownStream,Issue,Description,Remark,Status,UserEmployeeID,LocationID,Location,Flag) values" +
                                "(0,"+data.BatchID+",'"+ startDateTime+ "','"+endDateTime+ "',"+duration.TotalMinutes+",'"+ GetFullNameByEmployeeId(AccountEmployeeID)+ "','"+DateTime.Now+"','"+data.LineProd+"','"+data.PlanUnplan+"','"+data.Category+"','"+data.UpDownStream+"','','"+data.Description+ "','" +data.Remark +"','" +data.Status +"'," + AccountEmployeeID+"," +AccountDepartmentID +",'" + AccountLocation.Substring(0, 5) +"','User')";
                //if (area != "All Area")
                //    myQuery += " AND [LineProd] = '" + area + "'";
                //if (equipment != "All Equipment")
                //    myQuery += " AND [Remark] = '" + equipment + "'";
                //if (batch != "")
                //    myQuery += " AND [BatchID] = '" + batch + "'";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }
                SetTrueTempData(UIResources.UpdateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UpdateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult GetReportWithParam(string startDate, string endDate, string type, string batch, string area, string equipment) // param nya menyesuaikan
        {
            try
            {
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = "";
                DateTime dtStartDate = DateTime.ParseExact(startDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(endDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                dtEndDate = dtEndDate.AddDays(1);
                string startD = dtStartDate.ToString("yyyy-MM-dd");
                string endD = dtEndDate.ToString("yyyy-MM-dd");
                myQuery = @"SELECT * FROM " + type + " WHERE [StartDateTime] >= '" + startD + " 06:00:00' AND [StartDateTime] < '" + endD + " 06:00:00'";
                if (area != "All Area")
                {
                    string a = string.Join(",", area.Split(',').Select(x => "'" + x + "'"));
                    myQuery += " AND [LineProd] IN (" + a + ")";
                }
                if (equipment != "All Equipment")
                {
                    myQuery += " AND [Remark] = '" + equipment + "'";
                }
                if (batch != "")
                {
                    myQuery += " AND [BatchID] = '" + batch + "'";
                }
                myQuery += " order by StartDateTime";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }
                SetFilterTempData(startDate, endDate, type, area, equipment, batch);
                var data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]);

                return Json(new { Status = "True", Data = data }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False", Error = ex.Message });
            }
        }
        [HttpPost]
        public ActionResult ValidateDowntime(string startDate, string endDate, string type, string batch, string area, string equipment) // param nya menyesuaikan
        {
            try
            {
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = "";
                var myQueryUpdate = "";
                DateTime dtStartDate = DateTime.ParseExact(startDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(endDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                dtEndDate = dtEndDate.AddDays(1);
                string startD = dtStartDate.ToString("yyyy-MM-dd");
                string endD = dtEndDate.ToString("yyyy-MM-dd");
                myQuery = @"SELECT * FROM " + type + " WHERE [StartDateTime] >= '" + startD + " 06:00:00' AND [StartDateTime] < '" + endD + " 06:00:00'";
                //var myQueryUpdate = @"UPDATE " + type + " SET ApprovedBy='" + GetFullNameByEmployeeId(AccountEmployeeID) + "' WHERE ( ApproverEmployeeID='" + AccountEmployeeID + "') AND (convert(date,[StartDateTime]) BETWEEN '" + startDate + "' AND '" + endDate + "' )";
                myQueryUpdate = @"UPDATE " + type + " SET ApprovedBy='" + GetFullNameByEmployeeId(AccountEmployeeID) + "', ApproverEmployeeID='" + AccountEmployeeID + "' WHERE ApproverEmployeeID IS NULL AND (convert(date,[StartDateTime]) BETWEEN '" + startDate + "' AND '" + endDate + "' )";
                if (area != "All Area")
                {
                    string a = string.Join(",", area.Split(',').Select(x => "'" + x + "'"));
                    myQuery += " AND [LineProd] IN (" + a + ")";
                    myQueryUpdate += " AND [LineProd] IN (" + a + ")";
                }
                if (equipment != "All Equipment")
                {
                    myQuery += " AND [Remark] = '" + equipment + "'";
                    myQueryUpdate += " AND [Remark] = '" + equipment + "'";
                }
                if (batch != "")
                {
                    myQuery += " AND [BatchID] = '" + batch + "'";
                    myQueryUpdate += " AND [BatchID] = '" + batch + "'";
                }

                myQuery += " order by StartDateTime";

                DataSet dset = new DataSet();
                DataSet dset2 = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    SqlCommand cmd2 = new SqlCommand(myQueryUpdate, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd2))
                    {
                        da.Fill(dset2);
                    }
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                var data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]);

                return Json(new { Status = "True", Data = data }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False", Error = ex.Message });
            }
        }

        [HttpPost]
        public ActionResult ExtractExcel(string data)
        {
            try
            {
                List<Dictionary<string, string>> s = string.IsNullOrEmpty(data) ?
                new List<Dictionary<string, string>>() :
                JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(data);

                List<string> listHeader = new List<string>
                {
                    "BatchID",
                    "StartDateTime",
                    "EndDateTime",
                    "Duration",
                    "LogUser",
                    "SubmitDateTime",
                    "LineProd",
                    "PlanUnplan",
                    "Category",
                    "UpDownStream",
                    "Issue",
                    "Description",
                    "Remark",
                    "Status",
                    "ApprovedBy",
                };

                List<List<string>> table = new List<List<string>>();
                table.Add(listHeader);
                s.ForEach(d =>
                {
                    List<string> row = new List<string>();
                    foreach (string k in d.Keys)
                    {
                        if (listHeader.Contains(k))
                        {
                            if (d[k] != null && d[k].Contains("/Date("))
                            {
                                string date = d[k].Substring(6, d[k].Length - 8);
                                row.Add((new DateTime(1970, 1, 1)).AddMilliseconds(double.Parse(date)).ToString("dd-MMM-yy hh:mm"));
                            }
                            else row.Add(d[k]);
                        }
                    }
                    table.Add(row);

                });

                Session["ReportPPDowntime"] = ExcelGenerator.PPRawDataExtract(new Dictionary<string, List<List<string>>>()
                {
                    { "ReportPPDowntime", table }
                });

                return Json(new { status = true }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { status = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult DownloadExcel()
        {
            if (Session["ReportPPDowntime"] != null)
            {
                byte[] data = Session["ReportPPDowntime"] as byte[];
                Session["ReportPPDowntime"] = null;
                return File(data, "application/octet-stream", "ReportPPDowntime.xlsx");
            }
            return new EmptyResult();
        }

        public string DataTableToJSONWithJavaScriptSerializer(DataTable table)
        {
            JavaScriptSerializer jsSerializer = new JavaScriptSerializer();
            List<Dictionary<string, object>> parentRow = new List<Dictionary<string, object>>();
            Dictionary<string, object> childRow;
            foreach (DataRow row in table.Rows)
            {
                childRow = new Dictionary<string, object>();
                foreach (DataColumn col in table.Columns)
                {
                    childRow.Add(col.ColumnName, row[col]);
                }
                parentRow.Add(childRow);
            }
            return jsSerializer.Serialize(parentRow);
        }

        public string GetFullNameByEmployeeId(string empID)
        {
            string findUser = _employeeAppService.GetBy("EmployeeID", empID, false);
            EmployeeModel employeeModel = findUser.DeserializeToEmployee();
            return employeeModel.FullName.Replace("'", "''");
        }
        private void SetFilterTempData(string dStart,string dEnd,string FilterLine, string Area, string Equipment, string BatchNo)
        {
            //TempData["FilterFound"] = true;
            TempData["DateStart"] = dStart;
            TempData["DateEnd"] = dEnd;
            TempData["FilterLine"] = FilterLine;
            TempData["Area"] = Area;
            TempData["Equipment"] = Equipment;
            TempData["BatchNo"] = BatchNo;
        }
        private void GetFilterTempData()
        {
            ViewBag.FilterFound = TempData["FilterFound"];
            ViewBag.DateStart = TempData["DateStart"];
            ViewBag.DateEnd = TempData["DateEnd"];
            ViewBag.FilterLine = TempData["FilterLine"];
            ViewBag.Area = TempData["Area"];
            ViewBag.Equipment = TempData["Equipment"];
            ViewBag.BatchNo = TempData["BatchNo"];
            TempData["FilterFound"] = false;
        }
    }
}
