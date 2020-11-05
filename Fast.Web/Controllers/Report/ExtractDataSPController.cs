using Fast.Application.Interfaces;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using Fast.Infra.CrossCutting.Common;
using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Reflection;
using Newtonsoft.Json;

namespace Fast.Web.Controllers.Report
{
    public static class ExtensionChanif
    {
        public static List<T> ToList<T>(this DataTable table) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            List<T> result = new List<T>();

            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties);
                result.Add(item);
            }

            return result;
        }

        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                if (property.PropertyType == typeof(System.DayOfWeek))
                {
                    DayOfWeek day = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), row[property.Name].ToString());
                    property.SetValue(item, day, null);
                }
                else
                {
                    if (row[property.Name] == DBNull.Value)
                        property.SetValue(item, null, null);
                    else
                        property.SetValue(item, @row[property.Name], null);
                }
            }
            return item;
        }
    }

    public class ExtractDataSPController : BaseController<ExtractDataModel>
    {
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly ILPHAppService _lphAppService;
        private readonly ILPHComponentsAppService _lphComponentsAppService;
        private readonly ILPHValuesAppService _lphValuesAppService;
        private readonly ILPHExtrasAppService _lphExtrasAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly IUserAppService _userAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly ILoggerAppService _logger;

        public ExtractDataSPController(
            ILPHAppService lphAppService,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            ILPHComponentsAppService lphComponentsAppService,
            ILPHValuesAppService lphValuesAppService,
            ILPHExtrasAppService lphExtrasAppService,
            ILPHSubmissionsAppService lPHSubmissionsAppService,
            ILPHApprovalsAppService lphApprovalsAppService,
            IUserAppService userAppService,
            IEmployeeAppService employeeAppService,
            ILoggerAppService logger
             )
        {
            _lphAppService = lphAppService;
            _locationAppService = locationAppService;
            _referenceAppService = referenceAppService;
            _lphComponentsAppService = lphComponentsAppService;
            _lphValuesAppService = lphValuesAppService;
            _lphExtrasAppService = lphExtrasAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _userAppService = userAppService;
            _employeeAppService = employeeAppService;
            _lphApprovalAppService = lphApprovalsAppService;

            _logger = logger;
        }
        // GET: ExtractDataSP
        public ActionResult Index()
        {
            GetTempData();

            LocationTreeModel LocationTree = GetLocationTreeModel();
            ViewBag.LocationTree = LocationTree;

            var LPHType = BindDropDownLPHType();
            LPHType.Insert(0, new SelectListItem
            {
                Text = "All LPH SP",
                Value = "All LPH SP"
            });
            ViewBag.TypeList = LPHType;

            ViewBag.StatusList = BindDropDownStatus();

            return View();
        }

        [HttpPost]
        public ActionResult ExtractExcel(ExtractDataModel model)
        {
            try
            {
                UserModel user = (UserModel)Session["UserLogon"];
                if (!user.LocationID.HasValue)
                {
                    SetFalseTempData("Location for the logged user is invalid.");
                    return RedirectToAction("Index");
                }

                if (model.StartDate > model.EndDate)
                {
                    SetFalseTempData("Start Date must be less than End Date.");
                    return RedirectToAction("Index");
                }
                long locationID = model.SubDepartmentID;

                // Getting all data lph               
                string lphs = _lphAppService.GetAll(true);
                List<LPHModel> lphList = lphs.DeserializeToLPHList();
                //lphList = lphList.Where(x => x.LocationID == locationID).ToList();
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(model.ProdCenterID, "productioncenter");
                lphList = lphList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
                lphList = lphList.Where(x => x.ModifiedDate >= model.StartDate.AddHours(-12) && x.ModifiedDate <= model.EndDate.AddHours(12)).ToList();

                List<LPHModel> sortedLph = new List<LPHModel>();

                foreach (var item in lphList)
                {
                    var temp = item.MenuTitle;
                    string type = temp.Replace("Controller", "");
                    if (type == model.LPHType)
                    {
                        sortedLph.Add(item);
                    }
                }
                int countSortedLph = sortedLph.Count();
                if (countSortedLph == 0)
                {
                    SetFalseTempData(UIResources.DataNotFound);
                    return RedirectToAction("Index");
                }
                ArrayList myList = new ArrayList(50);
                if (sortedLph.Count() > 0)
                {
                    foreach (var item in sortedLph)
                    {
                        long lphId = item.ID;
                        string comps = _lphComponentsAppService.GetAll(true);
                        List<LPHComponentsModel> componentList = comps.DeserializeToLPHComponentList();

                        string values = _lphValuesAppService.GetAll(true);
                        List<LPHValuesModel> valueList = values.DeserializeToLPHValueList();

                        string extras = _lphExtrasAppService.GetAll(true);
                        List<LPHExtrasModel> extraList = extras.DeserializeToLPHExtraList();

                        List<string> compList = new List<string>();
                        List<string> valList = new List<string>();                                           
                        List<Extra> extList = new List<Extra>();

                        componentList = componentList.Where(x => x.LPHID == lphId).ToList();
                        extraList = extraList.Where(x => x.LPHID == lphId).ToList();

                        foreach (var ext in extraList)
                        {
                            Extra exs = new Extra();
                            exs.ID = ext.ID;
                            exs.Header = ext.HeaderName;
                            exs.Field = ext.FieldName;
                            exs.Value = ext.Value;
                            extList.Add(exs);
                        }

                        foreach (var comp in componentList)
                        {
                            compList.Add(comp.ComponentName);
                        }
                        foreach (var itemComp in componentList)
                        {
                            var resValue = valueList.Where(x => x.LPHComponentID == itemComp.ID).ToList().FirstOrDefault();
                            valList.Add(resValue.Value);
                        }
                     
                        myList.Add(compList);
                        myList.Add(valList);
                        myList.Add(extList);
                    }
                }

                byte[] excelData = ExcelGenerator.RawDataExtract(AccountName, model.LPHType, model.StartDate.ToString("dd-MMM-yy"), model.EndDate.ToString("dd-MMM-yy"), myList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Raw Data - " + model.LPHType + ".xlsx");
                Response.BinaryWrite(excelData);
                Response.End();

                SetTrueTempData(UIResources.ExtractSuccess);
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");

        }

        //chanif: aku buat fungsi baru saja supaya selama pengerjaan fungsi yg lama tetap dapat digunakan
        [HttpPost]
        public ActionResult GenerateRawData(ExtractDataModel model)
        {
            try
            {
                if (model.StartDate > model.EndDate)
                {
                    SetFalseTempData("Start Date must be less than End Date.");
                    return RedirectToAction("Index");
                }

                //supaya lebih efisien, getall yo ajur jum
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0")); //chanif: ambil yg submitted saja

                // shift 3 sudah disimpan sesuai tanggal kan?
                filters.Add(new QueryFilter("Date", model.StartDate.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", model.EndDate.ToString(), Operator.LessThanOrEqualTo));

                if (model.LPHType != "All LPH SP")
                    filters.Add(new QueryFilter("LPHHeader", model.LPHType+"Controller"));

                string submissionList = _lphSubmissionsAppService.Find(filters);
                List<LPHSubmissionsModel> submissions = submissionList.DeserializeToLPHSubmissionsList();

                //ambil data sesuai filter lokasi
                List<long> locations = new List<long>();
                var location = "All Locations";
                if (!string.IsNullOrEmpty(model.location3))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location3));
                    submissions = submissions.Where(x => x.LocationID == Int64.Parse(model.location3)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location2))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location2));

                    locations.Add(Int64.Parse(model.location2));

                    string subdeps = _locationAppService.FindBy("ParentID", model.location2, true);
                    var subdepsM = subdeps.DeserializeToLocationList();

                    foreach (var subdep in subdepsM)
                    {
                        locations.Add(subdep.ID);
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location1))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location1));

                    locations.Add(Int64.Parse(model.location1));

                    string deps = _locationAppService.FindBy("ParentID", model.location1, true);
                    var depsM = deps.DeserializeToLocationList();

                    foreach (var dep in depsM)
                    {
                        locations.Add(dep.ID);

                        string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }

                //mengurangi data yg perlu dicek
                if (model.Status == "approved")
                {
                    submissions = submissions.Where(x => x.IsComplete == true).ToList();
                } else if (model.Status == "submitted")
                {
                    submissions = submissions.Where(x => x.IsComplete == false).ToList();
                }

                if (submissions.Count() > 0)
                {
                    //ambil by min & max ID; sepertinya lebih cepat drpd find berkali2, benar tak?
                    var minSubmissionID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                    var maxSubmissionID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                    filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("IsDeleted", "0"));
                    filters.Add(new QueryFilter("LPHSubmissionID", minSubmissionID.ToString(), Operator.GreaterThanOrEqual));
                    filters.Add(new QueryFilter("LPHSubmissionID", maxSubmissionID.ToString(), Operator.LessThanOrEqualTo));

                    string approvals = _lphApprovalAppService.Find(filters);
                    List<LPHApprovalsModel> approvalList = approvals.DeserializeToLPHApprovalList();

                    //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                    foreach (var item in submissions.ToList())
                    {
                        string lphM = _lphAppService.GetById(item.LPHID); //harusnya getbyid lebih cepat drpd find where lphid >= lowestID
                        LPHModel lphModel = lphM.DeserializeToLPH();

                        // chanif: exclude LPH yang sudah dihapus
                        if (lphModel == null)
                        {
                            submissions.Remove(item);
                            continue;
                        }
                        else if (lphModel.IsDeleted)
                        {
                            submissions.Remove(item);
                            continue;
                        }


                        long tempApprover = 0;
                        LPHApprovalsModel approvalModel = approvalList.Where(x => x.LPHSubmissionID == item.ID).LastOrDefault();
                        if (approvalModel != null && (approvalModel.Status.Trim() != "" && (approvalModel.Status.Trim().ToLower() == model.Status || model.Status == "all")))
                        {
                            item.Status = approvalModel.Status;
                            tempApprover = approvalModel.ApproverID;
                            // masuk gan
                        }
                        else
                        {
                            // aneh, buang aja drpd error
                            submissions.Remove(item);
                            continue;
                        }


                        //item.CreatedAt = (DateTime)lphModel.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create

                        string user = _userAppService.GetById(item.UserID);
                        var userM = user.DeserializeToUser();

                        item.Location = String.IsNullOrEmpty(userM.Location) ? "" : userM.Location;
                        item.Creator = "[" + userM.EmployeeID + "] " + userM.UserName;

                        user = _userAppService.GetById(tempApprover);
                        userM = user.DeserializeToUser();
                        item.Approver = "[" + userM.EmployeeID + "] " + userM.UserName;

                        item.LPHHeader = item.LPHHeader.Replace("Controller", "");
                    }

                    if (submissions.Count() > 0)
                    {
                        ExcelPackage Ep = new ExcelPackage();
                        ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Submissions");

                        using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                        {
                            var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                            excelImage.SetPosition(0, 0, 0, 0);
                        }

                        Sheet.Cells["A3"].Value = UIResources.Title;
                        Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                        Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                        Sheet.Cells["B3"].Value = "Report Raw Data LPH SP";
                        Sheet.Cells["B4"].Value = AccountName;
                        Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                        Sheet.Cells[8, 1].Value = "LPH Type";
                        Sheet.Cells[9, 1].Value = "Location";
                        Sheet.Cells[10, 1].Value = "Date Range";

                        Sheet.Cells[8, 2].Value = model.LPHType;
                        Sheet.Cells[9, 2].Value = location;
                        Sheet.Cells[10, 2].Value = model.StartDate.ToString("dd-MMM-yy") + "  -  " + model.EndDate.ToString("dd-MMM-yy");


                        var colHeaderGroup = 12;
                        var colHeader = 13;
                        var rowHeader = 1;
                        var rowHeaderStart = rowHeader;
                        Sheet.Cells[colHeaderGroup, rowHeader].Value = "SUBMISSION INFORMATION";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Submission ID";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Location";
                        Sheet.Cells[colHeader, rowHeader++].Value = "LPH Type";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Submitter";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Status";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Approver";
                        Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                        using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                        }

                        var colContent = 14;
                        foreach (var item in submissions)
                        {
                            Sheet.Cells[colContent, 1].Value = item.ID;
                            Sheet.Cells[colContent, 2].Value = item.Location;
                            Sheet.Cells[colContent, 3].Value = item.LPHHeader;
                            Sheet.Cells[colContent, 4].Value = item.Creator;
                            Sheet.Cells[colContent, 5].Value = item.Status;
                            Sheet.Cells[colContent, 6].Value = item.Approver;
                            colContent++;
                        }
                        Sheet.Cells[13, 1, colContent, 6].AutoFitColumns();

                        // bagian yg sulit mulai dari sini. semangat......

                        var LPHIDs = submissions.Select(x => x.LPHID).ToList();
                        var minLPHID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.LPHID);
                        var maxLPHID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.LPHID);

                        filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("IsDeleted", "0"));
                        filters.Add(new QueryFilter("LPHID", minLPHID.ToString(), Operator.GreaterThanOrEqual));
                        filters.Add(new QueryFilter("LPHID", maxLPHID.ToString(), Operator.LessThanOrEqualTo));

                        string comps = _lphComponentsAppService.Find(filters);
                        List<LPHComponentsModel> componentList = comps.DeserializeToLPHComponentList();

                        // exclude data yg ndak masuk di submission
                        componentList = componentList.Where(x => LPHIDs.Contains(x.LPHID)).ToList();


                        if (componentList.Count() > 0)
                        {
                            // populate values

                            var compoIDs = componentList.Select(x => x.ID).ToList();
                            var minCompoID = componentList.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                            var maxCompoID = componentList.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                            filters = new List<QueryFilter>();
                            filters.Add(new QueryFilter("IsDeleted", "0"));
                            filters.Add(new QueryFilter("LPHComponentID", minCompoID.ToString(), Operator.GreaterThanOrEqual));
                            filters.Add(new QueryFilter("LPHComponentID", maxCompoID.ToString(), Operator.LessThanOrEqualTo));

                            string values = _lphValuesAppService.Find(filters);
                            List<LPHValuesModel> valueList = values.DeserializeToLPHValueList();

                            // exclude value yg ndak masuk di componentList
                            valueList = valueList.Where(x => compoIDs.Contains(x.LPHComponentID)).ToList();

                            // header untuk component
                            rowHeaderStart = rowHeader;
                            var compoNames = componentList.Select(x => x.ComponentName).Distinct().ToList();
                            Sheet.Cells[colHeaderGroup, rowHeader].Value = "LPH CONTENT";
                            foreach (var item in compoNames.ToList())
                            {
                                var content = item.Replace("generalInfo-", "").Replace("TeamLeader", "Supervisor");
                                if (String.IsNullOrWhiteSpace(content))
                                {
                                    compoNames.Remove(item);
                                    continue;
                                }
                                Sheet.Cells[colHeader, rowHeader++].Value = content;
                            }
                            Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                            using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                            }

                            // valuenya
                            colContent = 14;
                            var rowContent = 7;
                            foreach (var item in submissions)
                            {
                                rowContent = 7;
                                foreach (var name in compoNames)
                                {
                                    dynamic content = "";

                                    var compo = componentList.Where(x => x.LPHID == item.LPHID && x.ComponentName == name).FirstOrDefault();
                                    if (compo != null && compo.ID != 0)
                                    {
                                        var value = valueList.Where(x => x.LPHComponentID == compo.ID).FirstOrDefault();

                                        if (value != null && !String.IsNullOrEmpty(value.Value))
                                        {
                                            if (value.ValueType.Trim() == "Numeric")
                                            {
                                                if (value.Value.Contains("."))
                                                {
                                                    //tak anggap double
                                                    double number = 0;
                                                    Double.TryParse(value.Value, out number);
                                                    content = number;
                                                } else
                                                {
                                                    //tak anggep long
                                                    Int64 number = 0;
                                                    Int64.TryParse(value.Value, out number);
                                                    content = number;
                                                }
                                            }
                                            else if (value.ValueType.Trim() == "ImageURL")
                                            {
                                                if (value.Value == "_no_image.png")
                                                {
                                                    content = "no image";
                                                }
                                                else
                                                {
                                                    content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + item.LPHHeader.ToLower() + "/" + value.Value;
                                                }
                                            }
                                            else
                                            {
                                                content = value.Value;
                                            }
                                        }
                                    }

                                    Sheet.Cells[colContent, rowContent++].Value = content;
                                }

                                colContent++;
                            }

                            // populate extra; bagian tersulit; core of the core :D
                            filters = new List<QueryFilter>();
                            filters.Add(new QueryFilter("IsDeleted", "0"));
                            filters.Add(new QueryFilter("LPHID", minLPHID.ToString(), Operator.GreaterThanOrEqual));
                            filters.Add(new QueryFilter("LPHID", maxLPHID.ToString(), Operator.LessThanOrEqualTo));

                            string extras = _lphExtrasAppService.Find(filters);
                            List<LPHExtrasModel> extraList = extras.DeserializeToLPHExtraList();
                            // exclude data yg ndak masuk di submission
                            extraList = extraList.Where(x => LPHIDs.Contains(x.LPHID)).ToList();

                            var headerExtra = extraList.OrderBy(x=>x.ID).Select(x => x.HeaderName).Distinct().ToList();

                            //summary extra
                            rowHeaderStart = rowHeader;
                            Sheet.Cells[colHeaderGroup, rowHeader].Value = "EXTRAS (Detailed in next Sheets)";
                            foreach (var item in headerExtra)
                            {
                                Sheet.Cells[colHeader, rowHeader++].Value = item;
                            }
                            Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                            using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.PaleGreen);
                            }

                            var extraDivider = new Dictionary<string, int>();
                            foreach (var item in headerExtra)
                            {
                                extraDivider.Add(item, extraList.Where(x => x.HeaderName == item).Select(x => x.FieldName).Distinct().ToList().Count());
                            }

                            colContent = 14;
                            var rowContentTemp = rowContent;
                            foreach (var item in submissions)
                            {
                                rowContent = rowContentTemp;

                                foreach (var header in headerExtra)
                                {
                                    int content = 0;

                                    var extra = extraList.Where(x => x.LPHID == item.LPHID && x.HeaderName == header).ToList();
                                    if (extra != null && extra.Count() > 0)
                                    {
                                        content = extra.Count() / extraDivider[header];
                                    }

                                    Sheet.Cells[colContent, rowContent++].Value = content;
                                }

                                colContent++;
                            }
                            Sheet.Cells[13, 7, 13, rowContent - 1].AutoFitColumns();
                            Sheet.Cells[12, 1, colContent - 1, rowContent - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                            // generate sheet
                            foreach (var item in headerExtra)
                            {
                                if (extraDivider[item] > 0)
                                {
                                    Sheet = Ep.Workbook.Worksheets.Add(item);
                                    Sheet.Cells[1, 1].Value = item.ToUpper();
                                    Sheet.Cells[1, 1].Style.Font.Bold = true;
                                    Sheet.Cells[1, 1].Style.Font.Size = 16;

                                    var thisExtra = extraList.Where(x => x.HeaderName == item).OrderByDescending(x => x.ID).ToList();
                                    var ths = thisExtra.Select(x => x.FieldName).Distinct().ToList();
                                    /*
                                    var ths = new List<string>();
                                    foreach (var th in thisExtra)
                                    {
                                        if (ths.Contains(th.FieldName))
                                        {

                                        } else
                                        {
                                            ths.Add(th.FieldName);
                                        }
                                    }
                                    */

                                    thisExtra = thisExtra.OrderBy(x => x.LPHID).ToList();

                                    Sheet.Cells[3, 1].Value = "SUBMISSION INFORMATION";
                                    Sheet.Cells[4, 1].Value = "Submission ID";
                                    Sheet.Cells[4, 2].Value = "Location";
                                    Sheet.Cells[4, 3].Value = "LPH Type";
                                    Sheet.Cells[3, 1, 3, 3].Merge = true;
                                    using (var range = Sheet.Cells[3, 1, 4, 3])
                                    {
                                        range.Style.Font.Bold = true;
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                                    }

                                    Sheet.Cells[3, 4].Value = "TABLE CONTENT";
                                    var td = 4;
                                    foreach (var th in ths)
                                    {
                                        Sheet.Cells[4, td++].Value = th;
                                    }
                                    Sheet.Cells[3, 4, 3, td - 1].Merge = true;
                                    using (var range = Sheet.Cells[3, 4, 4, td - 1])
                                    {
                                        range.Style.Font.Bold = true;
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                    }

                                    var baris = 4;
                                    long tempLPHID = 0;
                                    var counter = 1;
                                    int flag = 0;
                                    foreach (var isi in thisExtra.ToList())
                                    {
                                        if (tempLPHID != isi.LPHID)
                                        {
                                            baris++;
                                            counter = 1;

                                            var LPHSubData = submissions.Where(x => x.LPHID == isi.LPHID).FirstOrDefault();

                                            Sheet.Cells[baris, counter++].Value = LPHSubData.ID;
                                            Sheet.Cells[baris, counter++].Value = LPHSubData.Location;
                                            Sheet.Cells[baris, counter++].Value = LPHSubData.LPHHeader;

                                            var currentExtra = thisExtra.Where(x => x.LPHID == LPHSubData.LPHID).ToList();
                                            foreach (var th in ths)
                                            {
                                                var tempval = currentExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                                dynamic content = "";

                                                if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                                {
                                                    if (tempval.ValueType.Trim() == "Numeric")
                                                    {
                                                        if (tempval.Value.Contains("."))
                                                        {
                                                            //tak anggap double
                                                            double number = 0;
                                                            Double.TryParse(tempval.Value, out number);
                                                            content = number;
                                                        }
                                                        else
                                                        {
                                                            //tak anggep long
                                                            Int64 number = 0;
                                                            Int64.TryParse(tempval.Value, out number);
                                                            content = number;
                                                        }
                                                    }
                                                    else if (tempval.ValueType.Trim() == "ImageURL")
                                                    {
                                                        if (tempval.Value == "_no_image.png")
                                                        {
                                                            content = "no image";
                                                        }
                                                        else
                                                        {
                                                            content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + LPHSubData.LPHHeader.ToLower() + "/" + tempval.Value;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        content = tempval.Value;
                                                    }
                                                }

                                                Sheet.Cells[baris, counter++].Value = content;

                                                thisExtra.Remove(tempval);
                                            }

                                            tempLPHID = isi.LPHID;
                                        }

                                        // else do nothing; agak aneh ya? soale urutan kesimpennya kacau; bingung nampilinnya gimana biar urut

                                        // 1 LPHID multi content
                                        flag++;
                                        if (flag % extraDivider[item] == 0)
                                        {
                                            tempLPHID = 0;
                                        }
                                    }

                                    Sheet.Cells[4, 1, baris, counter - 1].AutoFitColumns();
                                    Sheet.Cells[3, 1, baris, counter - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                            }

                            Response.Clear();
                            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            Response.AddHeader("content-disposition", "attachment;filename=Report_raw_data_LPH_SP.xlsx");
                            Response.BinaryWrite(Ep.GetAsByteArray());
                            Response.End();

                            SetTrueTempData(UIResources.ExtractSuccess);
                        }
                        else
                        {
                            SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                        }
                    }
                    else
                    {
                        SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                    }
                }
                else
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GenerateRawDataV2(ExtractDataModel model)
        {
            try
            {
                if (model.StartDate > model.EndDate)
                {
                    SetFalseTempData("Start Date must be less than End Date.");
                    return RedirectToAction("Index");
                }

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM LPHSubmissions WHERE (convert(date,[Date]) BETWEEN '" + model.StartDate.Date.ToShortDateString() + "' AND '" + model.EndDate.Date.ToShortDateString() + "' AND IsDeleted = 0)";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<LPHSubmissionsModel> submissions = jsondata.DeserializeToLPHSubmissionsList();

                if (model.LPHType != "All LPH SP")
                    submissions = submissions.Where(x => x.LPHHeader == model.LPHType + "Controller").ToList();

                //ambil data sesuai filter lokasi
                List<long> locations = new List<long>();
                var location = "All Locations";
                if (!string.IsNullOrEmpty(model.location3))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location3));
                    submissions = submissions.Where(x => x.LocationID == Int64.Parse(model.location3)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location2))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location2));

                    locations.Add(Int64.Parse(model.location2));

                    string subdeps = _locationAppService.FindBy("ParentID", model.location2, true);
                    var subdepsM = subdeps.DeserializeToLocationList();

                    foreach (var subdep in subdepsM)
                    {
                        locations.Add(subdep.ID);
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location1))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location1));

                    locations.Add(Int64.Parse(model.location1));

                    string deps = _locationAppService.FindBy("ParentID", model.location1, true);
                    var depsM = deps.DeserializeToLocationList();

                    foreach (var dep in depsM)
                    {
                        locations.Add(dep.ID);

                        string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }

                //mengurangi data yg perlu dicek
                if (model.Status == "approved")
                {
                    submissions = submissions.Where(x => x.IsComplete == true).ToList();
                }
                else if (model.Status == "submitted")
                {
                    submissions = submissions.Where(x => x.IsComplete == false).ToList();
                }

                if (submissions.Count() > 0)
                {
                    var submissionsID = submissions.Select(x => x.ID).ToList();
                    var LPHsID = submissions.Select(x => x.LPHID).ToList();

                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM LPHApprovals WHERE IsDeleted = 0 AND LPHSubmissionID IN (" + string.Join(",", submissionsID) + ")";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }


                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    List<LPHApprovalsModel> approvalList = jsondata.DeserializeToLPHApprovalList();

                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM LPHs WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", LPHsID) + ")";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }

                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    List<LPHModel> LPHList = jsondata.DeserializeToLPHList();


                    var Creators = approvalList.Select(x => x.UserID).ToList();
                    var Approvers = approvalList.Select(x => x.ApproverID).ToList();
                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM Users WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", (Creators.Concat(Approvers).ToList())) + ")";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }

                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    List<UserModel> UserList = jsondata.DeserializeToUserList();

                    //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                    foreach (var item in submissions.ToList())
                    {
                        LPHModel lphModel = LPHList.Where(x => x.ID == item.LPHID).FirstOrDefault();

                        // chanif: exclude LPH yang sudah dihapus
                        if (lphModel == null)
                        {
                            submissions.Remove(item);
                            continue;
                        }
                        else if (lphModel.IsDeleted)
                        {
                            submissions.Remove(item);
                            continue;
                        }


                        long tempApprover = 0;
                        LPHApprovalsModel approvalModel = approvalList.Where(x => x.LPHSubmissionID == item.ID).LastOrDefault();
                        if (approvalModel != null && (approvalModel.Status.Trim() != "" && (approvalModel.Status.Trim().ToLower() == model.Status || model.Status == "all")))
                        {
                            item.Status = approvalModel.Status;
                            tempApprover = approvalModel.ApproverID;
                            // masuk gan
                        }
                        else
                        {
                            // aneh, buang aja drpd error
                            submissions.Remove(item);
                            continue;
                        }


                        //item.CreatedAt = (DateTime)lphModel.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create

                        var userM = UserList.Where(x => x.ID == item.UserID).FirstOrDefault();

                        item.Location = String.IsNullOrEmpty(userM.Location) ? "" : userM.Location;
                        item.Creator = "[" + userM.EmployeeID + "] " + userM.UserName;

                        userM = UserList.Where(x => x.ID == tempApprover).FirstOrDefault();
                        item.Approver = "[" + userM.EmployeeID + "] " + userM.UserName;

                        item.LPHHeader = item.LPHHeader.Replace("Controller", "");
                    }

                    if (submissions.Count() > 0)
                    {
                        ExcelPackage Ep = new ExcelPackage();
                        ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Submissions");

                        using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                        {
                            var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                            excelImage.SetPosition(0, 0, 0, 0);
                        }

                        Sheet.Cells["A3"].Value = UIResources.Title;
                        Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                        Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                        Sheet.Cells["B3"].Value = "Report Raw Data LPH SP";
                        Sheet.Cells["B4"].Value = AccountName;
                        Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                        Sheet.Cells[8, 1].Value = "LPH Type";
                        Sheet.Cells[9, 1].Value = "Location";
                        Sheet.Cells[10, 1].Value = "Date Range";

                        Sheet.Cells[8, 2].Value = model.LPHType;
                        Sheet.Cells[9, 2].Value = location;
                        Sheet.Cells[10, 2].Value = model.StartDate.ToString("dd-MMM-yy") + "  -  " + model.EndDate.ToString("dd-MMM-yy");


                        var colHeaderGroup = 12;
                        var colHeader = 13;
                        var rowHeader = 1;
                        var rowHeaderStart = rowHeader;
                        Sheet.Cells[colHeaderGroup, rowHeader].Value = "SUBMISSION INFORMATION";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Submission ID";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Location";
                        Sheet.Cells[colHeader, rowHeader++].Value = "LPH Type";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Submitter";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Status";
                        Sheet.Cells[colHeader, rowHeader++].Value = "Approver";
                        Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                        using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                        }

                        var colContent = 14;
                        foreach (var item in submissions)
                        {
                            Sheet.Cells[colContent, 1].Value = item.ID;
                            Sheet.Cells[colContent, 2].Value = item.Location;
                            Sheet.Cells[colContent, 3].Value = item.LPHHeader;
                            Sheet.Cells[colContent, 4].Value = item.Creator;
                            Sheet.Cells[colContent, 5].Value = item.Status;
                            Sheet.Cells[colContent, 6].Value = item.Approver;
                            colContent++;
                        }
                        Sheet.Cells[13, 1, colContent, 6].AutoFitColumns();

                        // bagian yg sulit mulai dari sini. semangat......

                        var LPHIDs = submissions.Select(x => x.LPHID).ToList();
                        
                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        myQuery = @"SELECT * FROM LPHComponents WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                        dset = new DataSet();
                        using (SqlConnection con = new SqlConnection(strConString))
                        {
                            SqlCommand cmd = new SqlCommand(myQuery, con);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dset);
                            }
                        }

                        jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                        List<LPHComponentsModel> componentList = jsondata.DeserializeToLPHComponentList();

                        if (componentList.Count() > 0)
                        {
                            // populate values
                            var compoIDs = componentList.Select(x => x.ID).ToList();
                            var minCompoID = componentList.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                            var maxCompoID = componentList.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHValues WHERE IsDeleted = 0 AND LPHComponentID BETWEEN "+ minCompoID + " AND " + maxCompoID;

                            dset = new DataSet();
                            using (SqlConnection con = new SqlConnection(strConString))
                            {
                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dset);
                                }
                            }

                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                            List<LPHValuesModel> valueList = jsondata.DeserializeToLPHValueList();

                            // exclude value yg ndak masuk di componentList
                            //valueList = valueList.Where(x => compoIDs.Contains(x.LPHComponentID)).ToList();

                            // header untuk component
                            rowHeaderStart = rowHeader;
                            var compoNames = componentList.Select(x => x.ComponentName).Distinct().ToList();
                            Sheet.Cells[colHeaderGroup, rowHeader].Value = "LPH CONTENT";
                            foreach (var item in compoNames.ToList())
                            {
                                var content = item.Replace("generalInfo-", "").Replace("TeamLeader", "Supervisor");
                                if (String.IsNullOrWhiteSpace(content))
                                {
                                    compoNames.Remove(item);
                                    continue;
                                }
                                Sheet.Cells[colHeader, rowHeader++].Value = content;
                            }
                            Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                            using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                            }

                            // valuenya
                            colContent = 14;
                            var rowContent = 7;
                            foreach (var item in submissions)
                            {
                                var compoHere = componentList.Where(x => x.LPHID == item.LPHID).ToList();

                                minCompoID = compoHere.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                                maxCompoID = compoHere.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                                List<LPHValuesModel> valueHere = valueList.Where(x => x.LPHComponentID >= minCompoID && x.LPHComponentID <= maxCompoID).ToList();

                                rowContent = 7;
                                foreach (var name in compoNames)
                                {
                                    dynamic content = "";

                                    var compo = componentList.Where(x => x.LPHID == item.LPHID && x.ComponentName == name).FirstOrDefault();
                                    if (compo != null && compo.ID != 0)
                                    {
                                        var value = valueHere.Where(x => x.LPHComponentID == compo.ID).FirstOrDefault();

                                        if (value != null && !String.IsNullOrEmpty(value.Value))
                                        {
                                            if (value.ValueType.Trim() == "Numeric")
                                            {
                                                if (value.Value.Contains("."))
                                                {
                                                    //tak anggap double
                                                    double number = 0;
                                                    Double.TryParse(value.Value, out number);
                                                    content = number;
                                                }
                                                else
                                                {
                                                    //tak anggep long
                                                    Int64 number = 0;
                                                    Int64.TryParse(value.Value, out number);
                                                    content = number;
                                                }
                                            }
                                            else if (value.ValueType.Trim() == "ImageURL")
                                            {
                                                if (value.Value == "_no_image.png")
                                                {
                                                    content = "no image";
                                                }
                                                else
                                                {
                                                    content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + item.LPHHeader.ToLower() + "/" + value.Value;
                                                }
                                            }
                                            else
                                            {
                                                content = value.Value;
                                            }
                                        }
                                    }

                                    Sheet.Cells[colContent, rowContent++].Value = content;
                                }

                                colContent++;
                            }

                            // populate extra; bagian tersulit; core of the core :D
                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHExtras WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                            dset = new DataSet();
                            using (SqlConnection con = new SqlConnection(strConString))
                            {
                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dset);
                                }
                            }

                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                            List<LPHExtrasModel> extraList = jsondata.DeserializeToLPHExtraList();

                            var headerExtra = extraList.OrderBy(x => x.ID).Select(x => x.HeaderName).Distinct().ToList();

                            //summary extra
                            rowHeaderStart = rowHeader;
                            Sheet.Cells[colHeaderGroup, rowHeader].Value = "EXTRAS (Detailed in next Sheets)";
                            foreach (var item in headerExtra)
                            {
                                Sheet.Cells[colHeader, rowHeader++].Value = item;
                            }
                            Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                            using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.PaleGreen);
                            }

                            var extraDivider = new Dictionary<string, int>();
                            foreach (var item in headerExtra)
                            {
                                extraDivider.Add(item, extraList.Where(x => x.HeaderName == item).Select(x => x.FieldName).Distinct().ToList().Count());
                            }

                            colContent = 14;
                            var rowContentTemp = rowContent;
                            foreach (var item in submissions)
                            {
                                rowContent = rowContentTemp;

                                foreach (var header in headerExtra)
                                {
                                    int content = 0;

                                    var extra = extraList.Where(x => x.LPHID == item.LPHID && x.HeaderName == header).ToList();
                                    if (extra != null && extra.Count() > 0)
                                    {
                                        content = extra.Count() / extraDivider[header];
                                    }

                                    Sheet.Cells[colContent, rowContent++].Value = content;
                                }

                                colContent++;
                            }
                            Sheet.Cells[13, 7, 13, rowContent - 1].AutoFitColumns();
                            Sheet.Cells[12, 1, colContent - 1, rowContent - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                            // generate sheet
                            foreach (var item in headerExtra)
                            {
                                if (extraDivider[item] > 0)
                                {
                                    Sheet = Ep.Workbook.Worksheets.Add(item);
                                    Sheet.Cells[1, 1].Value = item.ToUpper();
                                    Sheet.Cells[1, 1].Style.Font.Bold = true;
                                    Sheet.Cells[1, 1].Style.Font.Size = 16;

                                    var thisExtra = extraList.Where(x => x.HeaderName == item).OrderByDescending(x => x.ID).ToList();
                                    var ths = thisExtra.Select(x => x.FieldName).Distinct().ToList();
                                    /*
                                    var ths = new List<string>();
                                    foreach (var th in thisExtra)
                                    {
                                        if (ths.Contains(th.FieldName))
                                        {

                                        } else
                                        {
                                            ths.Add(th.FieldName);
                                        }
                                    }
                                    */

                                    thisExtra = thisExtra.OrderBy(x => x.LPHID).ToList();

                                    Sheet.Cells[3, 1].Value = "SUBMISSION INFORMATION";
                                    Sheet.Cells[4, 1].Value = "Submission ID";
                                    Sheet.Cells[4, 2].Value = "Location";
                                    Sheet.Cells[4, 3].Value = "LPH Type";
                                    Sheet.Cells[3, 1, 3, 3].Merge = true;
                                    using (var range = Sheet.Cells[3, 1, 4, 3])
                                    {
                                        range.Style.Font.Bold = true;
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                                    }

                                    Sheet.Cells[3, 4].Value = "TABLE CONTENT";
                                    var td = 4;
                                    foreach (var th in ths)
                                    {
                                        Sheet.Cells[4, td++].Value = th;
                                    }
                                    Sheet.Cells[3, 4, 3, td - 1].Merge = true;
                                    using (var range = Sheet.Cells[3, 4, 4, td - 1])
                                    {
                                        range.Style.Font.Bold = true;
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                    }

                                    var baris = 4;
                                    long tempLPHID = 0;
                                    var counter = 1;
                                    int flag = 0;
                                    foreach (var isi in thisExtra.ToList())
                                    {
                                        if (tempLPHID != isi.LPHID)
                                        {
                                            baris++;
                                            counter = 1;

                                            var LPHSubData = submissions.Where(x => x.LPHID == isi.LPHID).FirstOrDefault();

                                            Sheet.Cells[baris, counter++].Value = LPHSubData.ID;
                                            Sheet.Cells[baris, counter++].Value = LPHSubData.Location;
                                            Sheet.Cells[baris, counter++].Value = LPHSubData.LPHHeader;

                                            var currentExtra = thisExtra.Where(x => x.LPHID == LPHSubData.LPHID).ToList();
                                            foreach (var th in ths)
                                            {
                                                var tempval = currentExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                                dynamic content = "";

                                                if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                                {
                                                    if (tempval.ValueType.Trim() == "Numeric")
                                                    {
                                                        if (tempval.Value.Contains("."))
                                                        {
                                                            //tak anggap double
                                                            double number = 0;
                                                            Double.TryParse(tempval.Value, out number);
                                                            content = number;
                                                        }
                                                        else
                                                        {
                                                            //tak anggep long
                                                            Int64 number = 0;
                                                            Int64.TryParse(tempval.Value, out number);
                                                            content = number;
                                                        }
                                                    }
                                                    else if (tempval.ValueType.Trim() == "ImageURL")
                                                    {
                                                        if (tempval.Value == "_no_image.png")
                                                        {
                                                            content = "no image";
                                                        }
                                                        else
                                                        {
                                                            content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + LPHSubData.LPHHeader.ToLower() + "/" + tempval.Value;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        content = tempval.Value;
                                                    }
                                                }

                                                Sheet.Cells[baris, counter++].Value = content;

                                                thisExtra.Remove(tempval);
                                            }

                                            tempLPHID = isi.LPHID;
                                        }

                                        // else do nothing; agak aneh ya? soale urutan kesimpennya kacau; bingung nampilinnya gimana biar urut

                                        // 1 LPHID multi content
                                        flag++;
                                        if (flag % extraDivider[item] == 0)
                                        {
                                            tempLPHID = 0;
                                        }
                                    }

                                    Sheet.Cells[4, 1, baris, counter - 1].AutoFitColumns();
                                    Sheet.Cells[3, 1, baris, counter - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                }
                            }

                            Response.Clear();
                            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            Response.AddHeader("content-disposition", "attachment;filename=Report_raw_data_LPH_SP.xlsx");
                            Response.BinaryWrite(Ep.GetAsByteArray());
                            Response.End();

                            SetTrueTempData(UIResources.ExtractSuccess);
                        }
                        else
                        {
                            SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                        }
                    }
                    else
                    {
                        SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                    }
                }
                else
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GenerateRawDataCategorized(ExtractDataModel model)
        {
            try
            {
                if (model.StartDate > model.EndDate)
                {
                    SetFalseTempData("Start Date must be less than End Date.");
                    return RedirectToAction("Index");
                }

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM LPHSubmissions WHERE (convert(date,[Date]) BETWEEN '" + model.StartDate.Date.ToShortDateString() + "' AND '" + model.EndDate.Date.ToShortDateString() + "' AND IsDeleted = 0)";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<LPHSubmissionsModel> submissions = jsondata.DeserializeToLPHSubmissionsList();

                var categories = new List<string>();

                if (model.LPHType == "All LPH SP")
                {
                    var LPHType = BindDropDownLPHType();
                    foreach(var type in LPHType)
                    {
                        categories.Add(type.Value + "Controller");
                    }
                }
                else
                {
                    categories.Add(model.LPHType + "Controller");
                }

                

                //ambil data sesuai filter lokasi
                List<long> locations = new List<long>();
                var location = "All Locations";
                if (!string.IsNullOrEmpty(model.location3))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location3));
                    submissions = submissions.Where(x => x.LocationID == Int64.Parse(model.location3)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location2))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location2));

                    locations.Add(Int64.Parse(model.location2));

                    string subdeps = _locationAppService.FindBy("ParentID", model.location2, true);
                    var subdepsM = subdeps.DeserializeToLocationList();

                    foreach (var subdep in subdepsM)
                    {
                        locations.Add(subdep.ID);
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location1))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location1));

                    locations.Add(Int64.Parse(model.location1));

                    string deps = _locationAppService.FindBy("ParentID", model.location1, true);
                    var depsM = deps.DeserializeToLocationList();

                    foreach (var dep in depsM)
                    {
                        locations.Add(dep.ID);

                        string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }

                //mengurangi data yg perlu dicek
                if (model.Status == "approved")
                {
                    submissions = submissions.Where(x => x.IsComplete == true).ToList();
                }
                else if (model.Status == "submitted")
                {
                    submissions = submissions.Where(x => x.IsComplete == false).ToList();
                }

                if (submissions.Count() > 0)
                {
                    ExcelPackage Ep = new ExcelPackage();

                    foreach (var category in categories)
                    {
                        var submissionHere = submissions.Where(x => x.LPHHeader == category).ToList();

                        if (submissionHere.Count > 0)
                        {
                            var submissionsID = submissionHere.Select(x => x.ID).ToList();
                            var LPHsID = submissionHere.Select(x => x.LPHID).ToList();

                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHApprovals WHERE IsDeleted = 0 AND LPHSubmissionID IN (" + string.Join(",", submissionsID) + ")";

                            dset = new DataSet();
                            using (SqlConnection con = new SqlConnection(strConString))
                            {
                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dset);
                                }
                            }

                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                            List<LPHApprovalsModel> approvalList = jsondata.DeserializeToLPHApprovalList();

                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHs WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", LPHsID) + ")";

                            dset = new DataSet();
                            using (SqlConnection con = new SqlConnection(strConString))
                            {
                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dset);
                                }
                            }

                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                            List<LPHModel> LPHList = jsondata.DeserializeToLPHList();

                            var Creators = approvalList.Select(x => x.UserID).ToList();
                            var Approvers = approvalList.Select(x => x.ApproverID).ToList();
                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM Users WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", (Creators.Concat(Approvers).Distinct().ToList())) + ")";

                            dset = new DataSet();
                            using (SqlConnection con = new SqlConnection(strConString))
                            {
                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dset);
                                }
                            }

                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                            List<UserModel> UserList = jsondata.DeserializeToUserList();

                            //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                            foreach (var item in submissionHere.ToList())
                            {
                                LPHModel lphModel = LPHList.Where(x => x.ID == item.LPHID).FirstOrDefault();

                                // chanif: exclude LPH yang sudah dihapus
                                if (lphModel == null)
                                {
                                    submissionHere.Remove(item);
                                    continue;
                                }
                                else if (lphModel.IsDeleted)
                                {
                                    submissionHere.Remove(item);
                                    continue;
                                }


                                long tempApprover = 0;
                                LPHApprovalsModel approvalModel = approvalList.Where(x => x.LPHSubmissionID == item.ID && x.Status.Trim() != "").LastOrDefault();
                                if (approvalModel != null && (approvalModel.Status.Trim() != "" && (approvalModel.Status.Trim().ToLower() == model.Status || model.Status == "all")))
                                {
                                    item.Status = approvalModel.Status;
                                    tempApprover = approvalModel.ApproverID;
                                    // masuk gan
                                }
                                else
                                {
                                    // aneh, buang aja drpd error
                                    submissionHere.Remove(item);
                                    continue;
                                }

                                //item.CreatedAt = (DateTime)lphModel.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create
                                var userM = UserList.Where(x => x.ID == item.UserID).FirstOrDefault();
                                if (userM == null)
                                {
                                    item.Location = "";
                                    item.Creator = "";
                                }
                                else
                                {
                                    item.Location = String.IsNullOrEmpty(userM.Location) ? "" : userM.Location;
                                    item.Creator = "[" + userM.EmployeeID + "] " + userM.UserName;
                                }

                                userM = UserList.Where(x => x.ID == tempApprover).FirstOrDefault();
                                if (userM == null)
                                    item.Approver = "";
                                else
                                    item.Approver = "[" + userM.EmployeeID + "] " + userM.UserName;

                                item.LPHHeader = item.LPHHeader.Replace("Controller", "");
                            }

                            if (submissionHere.Count() > 0)
                            {
                                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(category.Replace("Controller", ""));

                                using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                                {
                                    var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                                    excelImage.SetPosition(0, 0, 0, 0);
                                }

                                Sheet.Cells["A3"].Value = UIResources.Title;
                                Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                                Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                                Sheet.Cells["B3"].Value = "Report Raw Data LPH SP";
                                Sheet.Cells["B4"].Value = AccountName;
                                Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                                Sheet.Cells[8, 1].Value = "LPH Type";
                                Sheet.Cells[9, 1].Value = "Location";
                                Sheet.Cells[10, 1].Value = "Date Range";

                                Sheet.Cells[8, 2].Value = category.Replace("Controller", "");
                                Sheet.Cells[9, 2].Value = location;
                                Sheet.Cells[10, 2].Value = model.StartDate.ToString("dd-MMM-yy") + "  -  " + model.EndDate.ToString("dd-MMM-yy");

                                var colHeaderGroup = 12;
                                var colHeader = 13;
                                var rowHeader = 1;
                                var rowHeaderStart = rowHeader;
                                Sheet.Cells[colHeaderGroup, rowHeader].Value = "SUBMISSION INFORMATION";
                                Sheet.Cells[colHeader, rowHeader++].Value = "Submission ID";
                                Sheet.Cells[colHeader, rowHeader++].Value = "Location";
                                Sheet.Cells[colHeader, rowHeader++].Value = "LPH Type";
                                Sheet.Cells[colHeader, rowHeader++].Value = "Submitter";
                                Sheet.Cells[colHeader, rowHeader++].Value = "Status";
                                Sheet.Cells[colHeader, rowHeader++].Value = "Approver";
                                Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                                using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                                {
                                    range.Style.Font.Bold = true;
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                                }

                                var colContent = 14;
                                foreach (var item in submissionHere)
                                {
                                    Sheet.Cells[colContent, 1].Value = item.ID;
                                    Sheet.Cells[colContent, 2].Value = item.Location;
                                    Sheet.Cells[colContent, 3].Value = item.LPHHeader;
                                    Sheet.Cells[colContent, 4].Value = item.Creator;
                                    Sheet.Cells[colContent, 5].Value = item.Status;
                                    Sheet.Cells[colContent, 6].Value = item.Approver;
                                    colContent++;
                                }
                                Sheet.Cells[13, 1, colContent, 6].AutoFitColumns();

                                // bagian yg sulit mulai dari sini. semangat......

                                var LPHIDs = submissionHere.Select(x => x.LPHID).ToList();

                                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                myQuery = @"SELECT * FROM LPHComponents WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                                dset = new DataSet();
                                using (SqlConnection con = new SqlConnection(strConString))
                                {
                                    SqlCommand cmd = new SqlCommand(myQuery, con);
                                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                    {
                                        da.Fill(dset);
                                    }
                                }

                                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                List<LPHComponentsModel> componentList = jsondata.DeserializeToLPHComponentList();

                                if (componentList.Count() > 0)
                                {
                                    // populate values
                                    var compoIDs = componentList.Select(x => x.ID).ToList();
                                    var minCompoID = componentList.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                                    var maxCompoID = componentList.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                    myQuery = @"SELECT * FROM LPHValues WHERE IsDeleted = 0 AND LPHComponentID BETWEEN " + minCompoID + " AND " + maxCompoID;

                                    dset = new DataSet();
                                    using (SqlConnection con = new SqlConnection(strConString))
                                    {
                                        SqlCommand cmd = new SqlCommand(myQuery, con);
                                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                        {
                                            da.Fill(dset);
                                        }
                                    }

                                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                    List<LPHValuesModel> valueList = jsondata.DeserializeToLPHValueList();
                                    
                                    // header untuk component
                                    rowHeaderStart = rowHeader;
                                    //var kolom_kosong = new List<int>();

                                    var compoNames = componentList.Where(x=>x.LPHID == LPHIDs[LPHIDs.Count()-1]).Select(x => x.ComponentName).ToList();
                                    //ar compoNames = componentList.Select(x => x.ComponentName).Distinct().ToList();
                                    Sheet.Cells[colHeaderGroup, rowHeader].Value = "LPH CONTENT";
                                    foreach (var item in compoNames)
                                    {
                                        var content = item.Replace("generalInfo-", "").Replace("GeneralInfo-", "").Replace("TeamLeader", "Supervisor");
                                        /*
                                         * karena gak ada pengecekan lagi di value, jadi sulit kalau ada komponen yg kosong
                                         * 
                                        if (String.IsNullOrWhiteSpace(content))
                                        {
                                            compoNames.Remove(item);
                                            continue;
                                        }
                                        */
                                        if (String.IsNullOrWhiteSpace(content))
                                        {
                                            //kolom_kosong.Add(rowHeader);
                                            //continue;
                                        }

                                        Sheet.Cells[colHeader, rowHeader++].Value = content;
                                    }
                                    Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;
                                    using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                                    {
                                        range.Style.Font.Bold = true;
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                    }

                                    // valuenya
                                    colContent = 14;
                                    var rowContent = 7;
                                    foreach (var item in submissionHere)
                                    {
                                        var compoHere = componentList.Where(x => x.LPHID == item.LPHID).ToList();

                                        compoIDs = compoHere.Select(x => x.ID).ToList();
                                        minCompoID = compoHere.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                                        maxCompoID = compoHere.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                                        List<LPHValuesModel> valueHere = valueList.Where(x => x.LPHComponentID >= minCompoID && x.LPHComponentID <= maxCompoID).ToList();
                                        valueHere = valueHere.Where(x => compoIDs.Contains(x.LPHComponentID)).ToList();
                                        valueHere = valueHere.OrderBy(x => x.LPHComponentID).ToList();

                                        rowContent = 7;
                                        var counter = rowContent;
                                        foreach (var value in valueHere)
                                        {
                                            //if (kolom_kosong.Contains(counter))
                                            //{
                                            //} else
                                            //{
                                                dynamic content = "";
                                                if (value != null && !String.IsNullOrEmpty(value.Value))
                                                {
                                                    if (value.ValueType.Trim() == "Numeric")
                                                    {
                                                        if (value.Value.Contains("."))
                                                        {
                                                            //tak anggap double
                                                            double number = 0;
                                                            Double.TryParse(value.Value, out number);
                                                            content = number;
                                                        }
                                                        else
                                                        {
                                                            //tak anggep long
                                                            Int64 number = 0;
                                                            Int64.TryParse(value.Value, out number);
                                                            content = number;
                                                        }
                                                    }
                                                    else if (value.ValueType.Trim() == "ImageURL")
                                                    {
                                                        if (value.Value == "_no_image.png")
                                                        {
                                                            content = "no image";
                                                        }
                                                        else
                                                        {
                                                            content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + item.LPHHeader.ToLower() + "/" + value.Value;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        content = value.Value;
                                                    }
                                                }

                                                Sheet.Cells[colContent, rowContent++].Value = content;
                                            //}

                                            counter++;
                                        }

                                        colContent++;
                                    }
                                    

                                    // populate extra; bagian tersulit; core of the core :D
                                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                    myQuery = @"SELECT * FROM LPHExtras WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                                    dset = new DataSet();
                                    using (SqlConnection con = new SqlConnection(strConString))
                                    {
                                        SqlCommand cmd = new SqlCommand(myQuery, con);
                                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                        {
                                            da.Fill(dset);
                                        }
                                    }

                                    

                                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                    List<LPHExtrasModel> extraList = jsondata.DeserializeToLPHExtraList();

                                    var headerExtra = extraList.Select(x => x.HeaderName).Distinct().ToList();

                                    
                                    if (headerExtra.Count > 0)
                                    {
                                        //summary extra
                                        rowHeaderStart = rowHeader;
                                        Sheet.Cells[colHeaderGroup, rowHeader].Value = "EXTRAS (Detailed in next Sheets)";

                                        foreach (var item in headerExtra)
                                        {
                                            Sheet.Cells[colHeader, rowHeader++].Value = item;
                                        }

                                        Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeaderGroup, rowHeader - 1].Merge = true;

                                        using (var range = Sheet.Cells[colHeaderGroup, rowHeaderStart, colHeader, rowHeader - 1])
                                        {
                                            range.Style.Font.Bold = true;
                                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            range.Style.Fill.BackgroundColor.SetColor(Color.PaleGreen);
                                        }

                                        var extraDivider = new Dictionary<string, int>();
                                        foreach (var item in headerExtra)
                                        {
                                            extraDivider.Add(item, extraList.Where(x => x.HeaderName == item).Select(x => x.FieldName).Distinct().ToList().Count());
                                        }

                                        colContent = 14;
                                        var rowContentTemp = rowContent;
                                        foreach (var item in submissionHere)
                                        {
                                            rowContent = rowContentTemp;

                                            foreach (var header in headerExtra)
                                            {
                                                int content = 0;

                                                var extra = extraList.Where(x => x.LPHID == item.LPHID && x.HeaderName == header).ToList();
                                                if (extra != null && extra.Count() > 0)
                                                {
                                                    content = extra.Count() / extraDivider[header];
                                                }

                                                Sheet.Cells[colContent, rowContent++].Value = content;
                                            }

                                            colContent++;
                                        }
                                        Sheet.Cells[13, 7, 13, rowContent - 1].AutoFitColumns();
                                        Sheet.Cells[12, 1, colContent - 1, rowContent - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                        // generate sheet
                                        foreach (var item in headerExtra)
                                        {
                                            if (extraDivider[item] > 0)
                                            {

                                                var item_text = item;
                                                if (category == "GWGeneralController")
                                                {
                                                    if (item == "receiveWHS")
                                                        item_text = "incoming";
                                                    else if (item == "pouring")
                                                        item_text = "usage";
                                                }


                                                Sheet = Ep.Workbook.Worksheets.Add(category.Replace("Controller", "") + " - " + item_text);
                                                Sheet.Cells[1, 1].Value = item_text.ToUpper();
                                                Sheet.Cells[1, 1].Style.Font.Bold = true;
                                                Sheet.Cells[1, 1].Style.Font.Size = 16;

                                                var thisExtra = extraList.Where(x => x.HeaderName == item).OrderByDescending(x => x.ID).ToList();
                                                var ths = thisExtra.Select(x => x.FieldName).Distinct().ToList();

                                                thisExtra = thisExtra.OrderBy(x => x.LPHID).ToList();

                                                Sheet.Cells[3, 1].Value = "SUBMISSION INFORMATION";
                                                Sheet.Cells[4, 1].Value = "Submission ID";
                                                Sheet.Cells[4, 2].Value = "Location";
                                                Sheet.Cells[4, 3].Value = "LPH Type";
                                                Sheet.Cells[4, 4].Value = "Submitter";
                                                Sheet.Cells[4, 5].Value = "Status";
                                                Sheet.Cells[4, 6].Value = "Approver";
                                                Sheet.Cells[4, 7].Value = "Date";
                                                Sheet.Cells[4, 8].Value = "Shift";
                                                Sheet.Cells[4, 9].Value = "Machine";

                                                Sheet.Cells[3, 1, 3, 9].Merge = true;
                                                using (var range = Sheet.Cells[3, 1, 4, 9])
                                                {
                                                    range.Style.Font.Bold = true;
                                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                                                }

                                                Sheet.Cells[3, 10].Value = "TABLE CONTENT";
                                                var td = 10;
                                                foreach (var th in ths)
                                                {
                                                    Sheet.Cells[4, td++].Value = th;
                                                }
                                                Sheet.Cells[3, 10, 3, td - 1].Merge = true;
                                                using (var range = Sheet.Cells[3, 10, 4, td - 1])
                                                {
                                                    range.Style.Font.Bold = true;
                                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                                }

                                                var baris = 4;
                                                long tempLPHID = 0;
                                                var counter = 1;
                                                int flag = 0;
                                                foreach (var isi in thisExtra.ToList())
                                                {
                                                    if (tempLPHID != isi.LPHID)
                                                    {
                                                        baris++;
                                                        counter = 1;

                                                        var LPHSubData = submissionHere.Where(x => x.LPHID == isi.LPHID).FirstOrDefault();

                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.ID;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Location;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.LPHHeader;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Creator;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Status;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Approver;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Date.ToString("dd-MMM-yy");
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Shift;
                                                        Sheet.Cells[baris, counter++].Value = LPHSubData.Machine;

                                                        var currentExtra = thisExtra.Where(x => x.LPHID == LPHSubData.LPHID).ToList();
                                                        foreach (var th in ths)
                                                        {
                                                            var tempval = currentExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                                            dynamic content = "";

                                                            if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                                            {
                                                                if (tempval.ValueType.Trim() == "Numeric")
                                                                {
                                                                    if (tempval.Value.Contains("."))
                                                                    {
                                                                        //tak anggap double
                                                                        double number = 0;
                                                                        Double.TryParse(tempval.Value, out number);
                                                                        content = number;
                                                                    }
                                                                    else
                                                                    {
                                                                        //tak anggep long
                                                                        Int64 number = 0;
                                                                        Int64.TryParse(tempval.Value, out number);
                                                                        content = number;
                                                                    }
                                                                }
                                                                else if (tempval.ValueType.Trim() == "ImageURL")
                                                                {
                                                                    if (tempval.Value == "_no_image.png")
                                                                    {
                                                                        content = "no image";
                                                                    }
                                                                    else
                                                                    {
                                                                        content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + LPHSubData.LPHHeader.ToLower() + "/" + tempval.Value;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    content = tempval.Value;
                                                                }
                                                            }

                                                            

                                                            if (category == "GWGeneralController" || category == "GWGeneral")
                                                            {
                                                                if (item == "pouring" || item == "usage")
                                                                {
                                                                    if (th == "NetWeight" || th == "GrossWeight")
                                                                    {
                                                                        content = double.Parse(content);
                                                                        Sheet.Cells[baris, counter].Style.Numberformat.Format = "0.00";
                                                                    }
                                                                }
                                                                
                                                                /* cuma buat coba
                                                                if (item == "crr")
                                                                {
                                                                    if (th.Trim() == "WeightVal")
                                                                    {
                                                                        Sheet.Cells[baris, counter+1].Style.Numberformat.Format = "0.00";
                                                                    }
                                                                }
                                                                */
                                                            }

                                                            Sheet.Cells[baris, counter++].Value = content;

                                                            thisExtra.Remove(tempval);
                                                        }

                                                        tempLPHID = isi.LPHID;
                                                    }

                                                    // else do nothing; agak aneh ya? soale urutan kesimpennya kacau; bingung nampilinnya gimana biar urut

                                                    // 1 LPHID multi content
                                                    flag++;
                                                    if (flag % extraDivider[item] == 0)
                                                    {
                                                        tempLPHID = 0;
                                                    }
                                                }

                                                Sheet.Cells[4, 1, baris, counter - 1].AutoFitColumns();
                                                Sheet.Cells[3, 1, baris, counter - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                                
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                                }
                            }
                        }
                    }

                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename=Report_raw_data_LPH_SP.xlsx");
                    Response.BinaryWrite(Ep.GetAsByteArray());
                    Response.End();

                    SetTrueTempData(UIResources.ExtractSuccess);
                }
                else
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GenerateRawDataSAP(ExtractDataModel model)
        {
            try
            {
                if (model.StartDate > model.EndDate)
                {
                    SetFalseTempData("Start Date must be less than End Date.");
                    return RedirectToAction("Index");
                }


                // Define langsung karena emang perlu yg fix, gak gonta ganti

                var defineCompo = new Dictionary<string, List<int>>();
                defineCompo.Add("MakerController", new List<int> { 0,1,2,3,4,9,10,11,12,13,14,15,16,17,18,19,20,21,26,27,28,29,30,31,32,33,34,35,36 });
                defineCompo.Add("PackerController", new List<int> { 0,1,2,3,4,6,7,9,10,11,12,13,14,15,17,18,20,21,23,24,26,27,29,37,38,39,44,45,46,47,48,49,50,51,52,53,54,59,63,64,65,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,92,93,95,96,99,100 });
                defineCompo.Add("CasePackerController", new List<int> { 0,1,2,3,4,8,9,10,11,12,13,14,15,16,17 });
                defineCompo.Add("FilterController", new List<int> { 0,1,2,3,4,9,10,11,12,14,20,21,40,41,42,43,44,45 });
                defineCompo.Add("RipperController", new List<int> { 0,1,2,3,4,9,10,11,12 });
                //defineCompo.Add("RobotController", new List<int> { -1 });
                defineCompo.Add("GWGeneralController", new List<int> { 0,1,2,3,4,5,6,7 });
                defineCompo.Add("GWReworkController", new List<int> { 0,1,2,3,6 });

                var defineExtra = new Dictionary<string, List<string>>();
                defineExtra.Add("MakerController", new List<string> { "trsrate" });
                defineExtra.Add("PackerController", new List<string> { "cbw", "bandroltake" });
                defineExtra.Add("CasePackerController", new List<string> { });
                defineExtra.Add("FilterController", new List<string> { });
                defineExtra.Add("RipperController", new List<string> { "inputTable", "output" });
                //defineExtra.Add("RobotController", new List<string> { "" });
                defineExtra.Add("GWGeneralController", new List<string> { "crr", "sapon", "filterWaste", "filter", "receiveWHS", "pouring" });
                defineExtra.Add("GWReworkController", new List<string> { "stockwip", "claim" });

                var extraHeader = new Dictionary<string, List<string>>();
                extraHeader.Add("trsrate", new List<string> { "date","time","speed","CTW","sample","sampling_time","recovery","remark" });
                extraHeader.Add("cbw", new List<string> { "Filename","Remark" });
                extraHeader.Add("bandroltake", new List<string> { "take","kode","weightTaker" });
                extraHeader.Add("inputTable", new List<string> { "PONumberInput","FACodeInput","AsalLUInput","MachineNumValInput","MachineInput","WeightInput","TypeInput","ResourceCigaretteInput","CRCodeInput","TypeCFInput" });
                extraHeader.Add("output", new List<string> { "RemarkOutput","WeightOutput","TypeOutput" });
                extraHeader.Add("crr", new List<string> { "PO_num","FACode","MachineVal","MachineNumVal","CRCode","ResourceCigarette","ItemVal","TipeCF","WeightVal" });
                extraHeader.Add("sapon", new List<string> { "MachineVal2","MachineNumVal2","WeightVal2" });
                extraHeader.Add("filterWaste", new List<string> { "PO_numFW","FACodeFW","AsalLUFW","MachineValFW","MachineNumValFW","FRCodeFW","WeightValFW" });
                extraHeader.Add("filter", new List<string> { "MachineVal","ItemCodeVal","BatchVal","TotalFilterVal","TakeTimeVal","ProductionDateVal" });
                extraHeader.Add("receiveWHS", new List<string> { "Code","WLSLosNumber","BatchNo","BoxNo","Date","NetWeight","GrossWeight" });
                extraHeader.Add("pouring", new List<string> { "Code","WLSLosNumber","BatchNo","BoxNo","Date","NetWeight","GrossWeight","PouringDate" });
                extraHeader.Add("stockwip", new List<string> { "LastVal" });
                extraHeader.Add("claim", new List<string> { "MachineNoVal","EarlyAVal","TotalAVal","SendAVal","FinalAVal" });


                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM LPHSubmissions WHERE (convert(date,[Date]) BETWEEN '" + model.StartDate.Date.ToShortDateString() + "' AND '" + model.EndDate.Date.ToShortDateString() + "' AND IsDeleted = 0)";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<LPHSubmissionsModel> submissions = jsondata.DeserializeToLPHSubmissionsList();

                var categories = new List<string>();

                if (model.LPHType == "All LPH SP")
                {
                    var LPHType = BindDropDownLPHType();
                    foreach (var type in LPHType)
                    {
                        categories.Add(type.Value + "Controller");
                    }
                }
                else
                {
                    categories.Add(model.LPHType + "Controller");
                }


                //ambil data sesuai filter lokasi
                List<long> locations = new List<long>();
                var location = "All Locations";
                if (!string.IsNullOrEmpty(model.location3))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location3));
                    submissions = submissions.Where(x => x.LocationID == Int64.Parse(model.location3)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location2))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location2));

                    locations.Add(Int64.Parse(model.location2));

                    string subdeps = _locationAppService.FindBy("ParentID", model.location2, true);
                    var subdepsM = subdeps.DeserializeToLocationList();

                    foreach (var subdep in subdepsM)
                    {
                        locations.Add(subdep.ID);
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }
                else if (!string.IsNullOrEmpty(model.location1))
                {
                    location = _locationAppService.GetLocationFullCode(Int64.Parse(model.location1));

                    locations.Add(Int64.Parse(model.location1));

                    string deps = _locationAppService.FindBy("ParentID", model.location1, true);
                    var depsM = deps.DeserializeToLocationList();

                    foreach (var dep in depsM)
                    {
                        locations.Add(dep.ID);

                        string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }

                //mengurangi data yg perlu dicek
                if (model.Status == "approved")
                {
                    submissions = submissions.Where(x => x.IsComplete == true).ToList();
                }
                else if (model.Status == "submitted")
                {
                    submissions = submissions.Where(x => x.IsComplete == false).ToList();
                }

                if (submissions.Count() > 0)
                {
                    ExcelPackage Ep = new ExcelPackage();

                    foreach (var category in categories)
                    {
                        if (defineCompo.ContainsKey(category))
                        {
                            var submissionHere = submissions.Where(x => x.LPHHeader == category).ToList();

                            if (submissionHere.Count > 0)
                            {
                                var submissionsID = submissionHere.Select(x => x.ID).ToList();
                                var LPHsID = submissionHere.Select(x => x.LPHID).ToList();

                                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                myQuery = @"SELECT * FROM LPHApprovals WHERE IsDeleted = 0 AND LPHSubmissionID IN (" + string.Join(",", submissionsID) + ")";

                                dset = new DataSet();
                                using (SqlConnection con = new SqlConnection(strConString))
                                {
                                    SqlCommand cmd = new SqlCommand(myQuery, con);
                                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                    {
                                        da.Fill(dset);
                                    }
                                }

                                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                List<LPHApprovalsModel> approvalList = jsondata.DeserializeToLPHApprovalList();

                                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                myQuery = @"SELECT * FROM LPHs WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", LPHsID) + ")";

                                dset = new DataSet();
                                using (SqlConnection con = new SqlConnection(strConString))
                                {
                                    SqlCommand cmd = new SqlCommand(myQuery, con);
                                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                    {
                                        da.Fill(dset);
                                    }
                                }

                                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                List<LPHModel> LPHList = jsondata.DeserializeToLPHList();

                                var Creators = approvalList.Select(x => x.UserID).ToList();
                                var Approvers = approvalList.Select(x => x.ApproverID).ToList();
                                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                myQuery = @"SELECT * FROM Users WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", (Creators.Concat(Approvers).Distinct().ToList())) + ")";

                                dset = new DataSet();
                                using (SqlConnection con = new SqlConnection(strConString))
                                {
                                    SqlCommand cmd = new SqlCommand(myQuery, con);
                                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                    {
                                        da.Fill(dset);
                                    }
                                }

                                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                List<UserModel> UserList = jsondata.DeserializeToUserList();

                                //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                                foreach (var item in submissionHere.ToList())
                                {
                                    LPHModel lphModel = LPHList.Where(x => x.ID == item.LPHID).FirstOrDefault();

                                    // chanif: exclude LPH yang sudah dihapus
                                    if (lphModel == null)
                                    {
                                        submissionHere.Remove(item);
                                        continue;
                                    }
                                    else if (lphModel.IsDeleted)
                                    {
                                        submissionHere.Remove(item);
                                        continue;
                                    }


                                    long tempApprover = 0;
                                    LPHApprovalsModel approvalModel = approvalList.Where(x => x.LPHSubmissionID == item.ID && x.Status.Trim() != "").LastOrDefault();
                                    if (approvalModel != null && (approvalModel.Status.Trim() != "" && (approvalModel.Status.Trim().ToLower() == model.Status || model.Status == "all")))
                                    {
                                        item.Status = approvalModel.Status;
                                        tempApprover = approvalModel.ApproverID;
                                        // masuk gan
                                    }
                                    else
                                    {
                                        // aneh, buang aja drpd error
                                        submissionHere.Remove(item);
                                        continue;
                                    }

                                    //item.CreatedAt = (DateTime)lphModel.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create
                                    var userM = UserList.Where(x => x.ID == item.UserID).FirstOrDefault();
                                    if (userM == null)
                                    {
                                        item.Location = "";
                                        item.Creator = "";
                                    }
                                    else
                                    {
                                        item.Location = String.IsNullOrEmpty(userM.Location) ? "" : userM.Location;
                                        item.Creator = "[" + userM.EmployeeID + "] " + userM.UserName;
                                    }

                                    userM = UserList.Where(x => x.ID == tempApprover).FirstOrDefault();
                                    if (userM == null)
                                        item.Approver = "";
                                    else
                                        item.Approver = "[" + userM.EmployeeID + "] " + userM.UserName;

                                    item.LPHHeader = item.LPHHeader.Replace("Controller", "");
                                }

                                if (submissionHere.Count() > 0)
                                {
                                    ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(category.Replace("Controller", ""));

                                    Sheet.Cells[1, 1].Value = "LPH Type :";
                                    Sheet.Cells[1, 2].Value = category.Replace("Controller", "");
                                    Sheet.Cells[1, 1].Style.Font.Bold = true;
                                    Sheet.Cells[1, 2].Style.Font.Bold = true;

                                    var colHeader = 3;
                                    var rowHeader = 1;
                                    var rowHeaderStart = rowHeader;
                                    Sheet.Cells[colHeader, rowHeader++].Value = "Submission ID";

                                    var colContent = 4;
                                    foreach (var item in submissionHere)
                                    {
                                        Sheet.Cells[colContent, 1].Value = item.ID;
                                        colContent++;
                                    }
                                    //Sheet.Cells[13, 1, colContent, 6].AutoFitColumns();

                                    // bagian yg sulit mulai dari sini. semangat......

                                    var LPHIDs = submissionHere.Select(x => x.LPHID).ToList();

                                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                    myQuery = @"SELECT * FROM LPHComponents WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                                    dset = new DataSet();
                                    using (SqlConnection con = new SqlConnection(strConString))
                                    {
                                        SqlCommand cmd = new SqlCommand(myQuery, con);
                                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                        {
                                            da.Fill(dset);
                                        }
                                    }

                                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                    List<LPHComponentsModel> componentList = jsondata.DeserializeToLPHComponentList();

                                    if (componentList.Count() > 0)
                                    {
                                        // populate values
                                        var compoIDs = componentList.Select(x => x.ID).ToList();
                                        var minCompoID = componentList.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                                        var maxCompoID = componentList.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                        myQuery = @"SELECT * FROM LPHValues WHERE IsDeleted = 0 AND LPHComponentID BETWEEN " + minCompoID + " AND " + maxCompoID;

                                        dset = new DataSet();
                                        using (SqlConnection con = new SqlConnection(strConString))
                                        {
                                            SqlCommand cmd = new SqlCommand(myQuery, con);
                                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                            {
                                                da.Fill(dset);
                                            }
                                        }

                                        jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                        List<LPHValuesModel> valueList = jsondata.DeserializeToLPHValueList();

                                        // header untuk component
                                        rowHeaderStart = rowHeader;
                                        var compoNames = componentList.Where(x => x.LPHID == LPHIDs[LPHIDs.Count() - 1]).Select(x => x.ComponentName).ToList();
                                        //ar compoNames = componentList.Select(x => x.ComponentName).Distinct().ToList();
                                        var counter = -1;
                                        foreach (var item in compoNames)
                                        {
                                            counter++;

                                            if (defineCompo[category].Contains(counter))
                                            {
                                                var content = item.Replace("generalInfo-", "").Replace("GeneralInfo-", "").Replace("TeamLeader", "Supervisor");
                                                Sheet.Cells[colHeader, rowHeader++].Value = content;
                                            }
                                        }

                                        using (var range = Sheet.Cells[3, 1, colHeader, rowHeader - 1])
                                        {
                                            range.Style.Font.Bold = true;
                                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                        }

                                        // valuenya
                                        
                                        colContent = 4;
                                        var rowContent = 2;
                                        foreach (var item in submissionHere)
                                        {
                                            var compoHere = componentList.Where(x => x.LPHID == item.LPHID).ToList();

                                            compoIDs = compoHere.Select(x => x.ID).ToList();
                                            minCompoID = compoHere.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                                            maxCompoID = compoHere.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                                            List<LPHValuesModel> valueHere = valueList.Where(x => x.LPHComponentID >= minCompoID && x.LPHComponentID <= maxCompoID).ToList();
                                            valueHere = valueHere.Where(x => compoIDs.Contains(x.LPHComponentID)).ToList();
                                            valueHere = valueHere.OrderBy(x => x.LPHComponentID).ToList();

                                            rowContent = 2;
                                            counter = -1;
                                            foreach (var value in valueHere)
                                            {
                                                counter++;

                                                if (defineCompo[category].Contains(counter))
                                                {
                                                    dynamic content = "";
                                                    if (value != null && !String.IsNullOrEmpty(value.Value))
                                                    {
                                                        if (value.ValueType.Trim() == "Numeric")
                                                        {
                                                            if (value.Value.Contains("."))
                                                            {
                                                                //tak anggap double
                                                                double number = 0;
                                                                Double.TryParse(value.Value, out number);
                                                                content = number;
                                                            }
                                                            else
                                                            {
                                                                //tak anggep long
                                                                Int64 number = 0;
                                                                Int64.TryParse(value.Value, out number);
                                                                content = number;
                                                            }
                                                        }
                                                        else if (value.ValueType.Trim() == "ImageURL")
                                                        {
                                                            if (value.Value == "_no_image.png")
                                                            {
                                                                content = "no image";
                                                            }
                                                            else
                                                            {
                                                                content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + item.LPHHeader.ToLower() + "/" + value.Value;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            content = value.Value;
                                                        }
                                                    }

                                                    Sheet.Cells[colContent, rowContent++].Value = content;
                                                }
                                            }

                                            colContent++;
                                        }

                                        Sheet.Cells[3, 1, colContent-1, rowContent-1].AutoFitColumns();
                                        Sheet.Cells[3, 1, colContent-1, rowContent-1].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                                        // populate extra; bagian tersulit; core of the core :D
                                        if (defineExtra[category].Count() > 0)
                                        {
                                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                            myQuery = @"SELECT * FROM LPHExtras WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

                                            dset = new DataSet();
                                            using (SqlConnection con = new SqlConnection(strConString))
                                            {
                                                SqlCommand cmd = new SqlCommand(myQuery, con);
                                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                                {
                                                    da.Fill(dset);
                                                }
                                            }

                                            jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                                            List<LPHExtrasModel> extraList = jsondata.DeserializeToLPHExtraList();

                                            foreach (var item in defineExtra[category])
                                            {
                                                Sheet = Ep.Workbook.Worksheets.Add(category.Replace("Controller", "") + " - " + item);
                                                Sheet.Cells[1, 1].Value = "Table Name :";
                                                Sheet.Cells[1, 1].Style.Font.Bold = true;
                                                Sheet.Cells[1, 2].Value = item.ToUpper();
                                                Sheet.Cells[1, 2].Style.Font.Bold = true;

                                                var extraDivider = new Dictionary<string, int>();
                                                extraDivider.Add(item, extraList.Where(x => x.HeaderName == item).Select(x => x.FieldName).Distinct().ToList().Count());

                                                if (extraDivider[item] > 0 && extraHeader.ContainsKey(item))
                                                {
                                                    var thisExtra = extraList.Where(x => x.HeaderName == item).OrderBy(x => x.LPHID).ToList();
                                                    var ths = extraHeader[item];
                                                    //thisExtra = thisExtra.OrderBy(x => x.LPHID).ToList();

                                                    Sheet.Cells[3, 1].Value = "Submission ID";
                                                    var td = 2;
                                                    foreach (var th in ths)
                                                    {
                                                        Sheet.Cells[3, td++].Value = th;
                                                    }
                                                    using (var range = Sheet.Cells[3, 1, 3, td - 1])
                                                    {
                                                        range.Style.Font.Bold = true;
                                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                        range.Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                                                    }

                                                    var baris = 3;
                                                    long tempLPHID = 0;
                                                    counter = 1;
                                                    int flag = 0;
                                                    foreach (var isi in thisExtra.ToList())
                                                    {
                                                        if (tempLPHID != isi.LPHID)
                                                        {
                                                            baris++;
                                                            counter = 1;

                                                            var LPHSubData = submissionHere.Where(x => x.LPHID == isi.LPHID).FirstOrDefault();
                                                            Sheet.Cells[baris, counter++].Value = LPHSubData.ID;

                                                            var currentExtra = thisExtra.Where(x => x.LPHID == LPHSubData.LPHID).ToList();
                                                            foreach (var th in ths)
                                                            {
                                                                var tempval = currentExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                                                dynamic content = "";

                                                                if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                                                {
                                                                    if (tempval.ValueType.Trim() == "Numeric")
                                                                    {
                                                                        if (tempval.Value.Contains("."))
                                                                        {
                                                                            //tak anggap double
                                                                            double number = 0;
                                                                            Double.TryParse(tempval.Value, out number);
                                                                            content = number;
                                                                        }
                                                                        else
                                                                        {
                                                                            //tak anggep long
                                                                            Int64 number = 0;
                                                                            Int64.TryParse(tempval.Value, out number);
                                                                            content = number;
                                                                        }
                                                                    }
                                                                    else if (tempval.ValueType.Trim() == "ImageURL")
                                                                    {
                                                                        if (tempval.Value == "_no_image.png")
                                                                        {
                                                                            content = "no image";
                                                                        }
                                                                        else
                                                                        {
                                                                            content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + LPHSubData.LPHHeader.ToLower() + "/" + tempval.Value;
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        content = tempval.Value;
                                                                    }
                                                                }

                                                                if (category == "GWGeneralController")
                                                                {
                                                                    if (item == "pouring")
                                                                    {
                                                                        if (th == "NetWeight" || th == "GrossWeight")
                                                                        {
                                                                            content = double.Parse(content);
                                                                            Sheet.Cells[baris, counter].Style.Numberformat.Format = "0.00";
                                                                        }
                                                                    }
                                                                }

                                                                Sheet.Cells[baris, counter++].Value = content;

                                                                thisExtra.Remove(tempval);
                                                            }

                                                            tempLPHID = isi.LPHID;
                                                        }

                                                        // else do nothing; agak aneh ya? soale urutan kesimpennya kacau; bingung nampilinnya gimana biar urut

                                                        // 1 LPHID multi content
                                                        flag++;
                                                        if (flag % extraDivider[item] == 0)
                                                        {
                                                            tempLPHID = 0;
                                                        }
                                                    }

                                                    Sheet.Cells[3, 1, baris, counter - 1].AutoFitColumns();
                                                    Sheet.Cells[3, 1, baris, counter - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                                }
                                                else
                                                {
                                                    Sheet.Cells[3, 1].Value = "No Data";
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                                    }
                                }
                            }
                        }
                    }

                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename=Report_raw_data_LPH_SP_for_SAP.xlsx");
                    Response.BinaryWrite(Ep.GetAsByteArray());
                    Response.End();

                    SetTrueTempData(UIResources.ExtractSuccess);
                }
                else
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }

        public static List<T> ConvertToList<T>(DataTable dt)
        {
            var columnNames = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var properties = typeof(T).GetProperties();
            return dt.AsEnumerable().Select(row => {
                var objT = Activator.CreateInstance<T>();
                foreach (var pro in properties)
                {
                    if (columnNames.Contains(pro.Name))
                    {
                        try
                        {
                            pro.SetValue(objT, row[pro.Name]);
                        }
                        catch (Exception ex) { }
                    }
                }
                return objT;
            }).ToList();
        }

        private List<SelectListItem> BindDropDownLPHType()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Maker",
                Value = "Maker"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Packer",
                Value = "Packer"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "CasePacker",
                Value = "CasePacker"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Filter",
                Value = "Filter"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Laser",
                Value = "Laser"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Ripper",
                Value = "Ripper"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Robot",
                Value = "Robot"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "GWGeneral",
                Value = "GWGeneral"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "GWRework",
                Value = "GWRework"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Mentholator",
                Value = "Mentholator"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "CigarilloMaker",
                Value = "CigarilloMaker"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "CigarilloPacker",
                Value = "CigarilloPacker"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "TDCPacker",
                Value = "TDCPacker"
            });
            return _menuList;
        }
        private List<SelectListItem> BindDropDownStatus()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Approved",
                Value = "approved"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Submitted Only",
                Value = "submitted"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Submitted + Approved",
                Value = "all"
            });
            
            return _menuList;
        }
        
        public class Extra
        {
            public long ID { get; set; }
            public string Header { get; set; }
            public string Field { get; set; }
            public string Value { get; set; }
        }


        private LocationTreeModel GetLocationTreeModel()
        {
            LocationTreeModel model = new LocationTreeModel();

            int index = 1;

            // get production center list
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var item in pcList)
            {
                string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
                ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
                model.ProductionCenters.Add(pcModel);
            }

            // get department list
            foreach (var pc in model.ProductionCenters)
            {
                LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
                string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
                List<LocationModel> departmentList = departments.DeserializeToLocationList();

                foreach (var d in departmentList)
                {
                    string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
                    DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

                    string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                    List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                    foreach (var subdeb in subdepartmentList)
                    {
                        string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
                        deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
                    }

                    pc.Departments.Add(deptModel);
                }
            }

            return model;
        }
    }
}