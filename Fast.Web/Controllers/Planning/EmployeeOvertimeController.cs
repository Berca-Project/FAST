using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
    [CustomAuthorize("employeeovertime")]
    public class EmployeeOvertimeController : BaseController<EmployeeOvertimeModel>
    {
        private readonly IEmployeeAppService _empService;
        private readonly IEmployeeAllAppService _empAllService;
        private readonly IEmployeeOvertimeAppService _empOvertimeService;
        private readonly IMenuAppService _menuService;
        private readonly IUserAppService _userAppService;
        private readonly IMppAppService _mppAppService;
        private readonly ILoggerAppService _logger;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;

        public EmployeeOvertimeController(
            IEmployeeAppService empService,
            IMppAppService mppAppService,
            IEmployeeAllAppService empAllService,
            ILoggerAppService logger,
            IUserAppService userAppService,
            IEmployeeOvertimeAppService empOvertimeService,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IMenuAppService menuService)
        {
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _userAppService = userAppService;
            _empService = empService;
            _empAllService = empAllService;
            _empOvertimeService = empOvertimeService;
            _menuService = menuService;
            _logger = logger;
            _mppAppService = mppAppService;
        }

        [HttpPost]
        public JsonResult AutoCompleteSpv(string prefix)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            if (prefix.All(Char.IsDigit))
                filters.Add(new QueryFilter("EmployeeID", prefix, Operator.StartsWith));
            else
                filters.Add(new QueryFilter("FullName", prefix, Operator.Contains));

            string emplist = _empAllService.Find(filters);
            List<EmployeeModel> empModelList = emplist.DeserializeToEmployeeList();

            string empAlllist = _empAllService.GetAll();
            List<EmployeeModel> empAllListModel = empAlllist.DeserializeToEmployeeList();
            empAllListModel = empAllListModel.Where(x => !string.IsNullOrEmpty(x.ReportToID1)).ToList();

            List<string> spvList = empAllListModel.Select(x => x.ReportToID1).Distinct().ToList();

            List<EmployeeModel> spvModelList = empModelList.Where(x => spvList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

            if (prefix.All(Char.IsDigit))
            {
                spvModelList = spvModelList.OrderBy(x => x.EmployeeID).ToList();
            }
            else
            {
                spvModelList = spvModelList.OrderBy(x => x.FullName).ToList();
            }

            return Json(spvModelList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Index()
        {
            GetTempData();

            ViewBag.StartDate = TempData["StartDate"];
            ViewBag.EndDate = TempData["EndDate"];

            EmployeeOvertimeModel model = new EmployeeOvertimeModel();
            model.Access = GetAccess(WebConstants.MenuSlug.EMP_OVERTIME, _menuService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            return View(model);
        }

        public ActionResult Edit(int id, string startDate = "", string endDate = "")
        {
            string empOvertime = _empOvertimeService.GetById(id);
            EmployeeOvertimeModel empOvertimeModel = empOvertime.DeserializeToEmployeeOvertime();
            string emp = _empService.GetBy("EmployeeID", empOvertimeModel.EmployeeID);
            EmployeeModel empModel = emp.DeserializeToEmployee();
            empOvertimeModel.FullName = empModel.FullName;
            empOvertimeModel.StartDate = startDate;
            empOvertimeModel.EndDate = endDate;

            ViewBag.OvertimeCategoryList = DropDownHelper.BindDropDownOvertimeCategory();

            return PartialView(empOvertimeModel);
        }

        [HttpPost]
        public ActionResult Edit(EmployeeOvertimeModel model)
        {
            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.EMP_OVERTIME, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);

                    _logger.LogError(ModelState.GetModelStateErrors(), AccountID);

                    return RedirectToAction("Index");
                }

                string empOver = _empOvertimeService.GetById(model.ID, true);
                EmployeeOvertimeModel empOverModel = empOver.DeserializeToEmployeeOvertime();

                empOverModel.OvertimeCategory = model.OvertimeCategory;

                string data = JsonHelper<EmployeeOvertimeModel>.Serialize(empOverModel);

                _empOvertimeService.Update(data);

                SetTrueTempData(UIResources.EditSucceed);

                TempData["StartDate"] = model.StartDate;
                TempData["EndDate"] = model.EndDate;
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult UpdateCategory()
        {
            return PartialView();
        }

        [HttpPost]
        public ActionResult UpdateCategory(EmployeeOvertimeModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidData);
                    return RedirectToAction("Index");
                }

                if (model.PostedFilename != null && model.PostedFilename.ContentLength > 0)
                {
                    Stream stream = model.PostedFilename.InputStream;

                    if (model.PostedFilename.FileName.ToLower().EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (model.PostedFilename.FileName.ToLower().EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileNotSupported);
                        return RedirectToAction("Index");
                    }

                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;
                    DataTable dt = new DataTable();
                    DataTable dt_ = reader.AsDataSet().Tables[0];

                    List<EmployeeOvertimeModel> empList = new List<EmployeeOvertimeModel>();

                    var categoryList = DropDownHelper.BindDropDownOvertimeCategory();

                    for (int index = 1; index < dt_.Rows.Count; index++)
                    {
                        string empId = dt_.Rows[index][0].ToString();
                        string overDate = dt_.Rows[index][2].ToString();
                        string overCategory = dt_.Rows[index][3].ToString();

                        if (empId == string.Empty || overCategory == string.Empty || overDate == string.Empty)
                        {
                            continue;
                        }

                        if (categoryList.Any(x => x.Value == overCategory))
                        {
                            EmployeeOvertimeModel empModel = new EmployeeOvertimeModel();
                            empModel.EmployeeID = dt_.Rows[index][0].ToString();
                            DateTime resultDate;
                            if (DateTime.TryParseExact(overDate, "dd-MMM-yy", CultureInfo.CurrentCulture, DateTimeStyles.None, out resultDate))
                            {
                                empModel.Date = resultDate;
                                empModel.OvertimeCategory = overCategory;

                                empList.Add(empModel);
                            }
                            else
                            {
                                SetFalseTempData("Invalid date format ( dd-MMM-yy ) " + overDate);
                                return RedirectToAction("Index");
                            }
                        }
                        else
                        {
                            SetFalseTempData("Invalid overtime category " + overCategory);
                            return RedirectToAction("Index");
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    if (empList.Count > 0)
                    {
                        string empOvertimeAllList = _empOvertimeService.GetAll();
                        List<EmployeeOvertimeModel> empOvertimeAllModelList = empOvertimeAllList.DeserializeToEmployeeOvertimeList();
                        List<EmployeeOvertimeModel> updateOvertimeList = new List<EmployeeOvertimeModel>();

                        foreach (var item in empList)
                        {
                            EmployeeOvertimeModel empModel = empOvertimeAllModelList.Where(x => x.EmployeeID == item.EmployeeID && x.Date == item.Date).FirstOrDefault();
                            if (empModel != null && empModel.OvertimeCategory != item.OvertimeCategory)
                            {
                                empModel.OvertimeCategory = item.OvertimeCategory;
                                updateOvertimeList.Add(empModel);
                            }
                        }

                        if (UpdateOvertimeCategory(updateOvertimeList))
                        {
                            SetTrueTempData(UIResources.UploadSucceed);
                        }
                        else
                        {
                            SetFalseTempData("Upload failed. Please try it again");
                        }
                    }
                    else
                    {
                        SetFalseTempData(UIResources.InvalidData);
                    }
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UploadFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                reader = null;
            }

            return RedirectToAction("Index");
        }

        private bool UpdateOvertimeCategory(List<EmployeeOvertimeModel> empUpdateList)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            string updateEmp = "UPDATE EmployeeOvertimes SET OvertimeCategory = @OvertimeCategory WHERE EmployeeID = @EmployeeID AND Date = @Date";

            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    foreach (var emp in empUpdateList)
                    {
                        SqlCommand command = new SqlCommand(updateEmp, connection, transaction);
                        command.Parameters.Add("@EmployeeID", SqlDbType.Char).Value = emp.EmployeeID;
                        command.Parameters.Add("@Date", SqlDbType.DateTime).Value = emp.Date;
                        command.Parameters.Add("@OvertimeCategory", SqlDbType.Char).Value = emp.OvertimeCategory;
                        command.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();

                    return false;
                }
            }

            return true;
        }

        public ActionResult DownloadTemplate(string supervisorID, string startDate, string endDate)
        {
            try
            {
                // Getting all data    			
                DateTime startDateFilter = DateTime.Parse(startDate);
                DateTime endDateFilter = DateTime.Parse(endDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", startDateFilter.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateFilter.ToString(), Operator.LessThanOrEqualTo));

                string empOvertimeAllList = _empOvertimeService.Find(filters);
                List<EmployeeOvertimeModel> empOvertimeAllModelList = empOvertimeAllList.DeserializeToEmployeeOvertimeList();

                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                string userList = _userAppService.GetAll(true);
                List<UserModel> userModelList = userList.DeserializeToUserList();

                List<EmployeeModel> empOvertimeList = new List<EmployeeModel>();
                List<EmployeeOvertimeModel> result = new List<EmployeeOvertimeModel>();

                if (!string.IsNullOrEmpty(supervisorID))
                {
                    var empModelListSpv = empModelList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == supervisorID || x.EmployeeID.Trim() == supervisorID).ToList();
                    var userModelListSpv = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == supervisorID || x.EmployeeID.Trim() == supervisorID).ToList();

                    foreach (var user in userModelListSpv)
                    {
                        EmployeeModel emp = empModelListSpv.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                        if (emp != null)
                        {
                            empOvertimeList.Add(emp);

                            var userModelListSpv1 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user.EmployeeID && x.SupervisorID.Trim() != supervisorID).ToList();
                            foreach (var user1 in userModelListSpv1)
                            {
                                EmployeeModel emp1 = empModelList.Where(x => x.EmployeeID.Trim() == user1.EmployeeID.Trim()).FirstOrDefault();
                                if (emp1 != null)
                                {
                                    empOvertimeList.Add(emp1);

                                    var userModelListSpv2 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user1.EmployeeID).ToList();
                                    foreach (var user2 in userModelListSpv2)
                                    {
                                        EmployeeModel emp2 = empModelList.Where(x => x.EmployeeID.Trim() == user2.EmployeeID.Trim()).FirstOrDefault();
                                        if (emp2 != null)
                                        {
                                            empOvertimeList.Add(emp2);

                                            var userModelListSpv3 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user2.EmployeeID).ToList();
                                            foreach (var user3 in userModelListSpv3)
                                            {
                                                EmployeeModel emp3 = empModelList.Where(x => x.EmployeeID.Trim() == user3.EmployeeID.Trim()).FirstOrDefault();
                                                if (emp3 != null)
                                                {
                                                    empOvertimeList.Add(emp3);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (var item in empOvertimeList)
                    {
                        var tempList = empOvertimeAllModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).ToList();
                        result.AddRange(tempList);
                    }

                    foreach (var item in result)
                    {
                        var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        if (emp != null)
                            item.FullName = emp.FullName;
                    }
                }
                else
                {
                    result.Add(new EmployeeOvertimeModel { EmployeeID = "000000001", FullName = "Testing Name", Date = DateTime.Now, OvertimeCategory = "Daily Work" });
                }

                var categoryList = DropDownHelper.BindDropDownOvertimeCategory();

                byte[] excelData = ExcelGenerator.ExportOvertimeCategory(result, categoryList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Template-Overtime-Category.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();

                ViewBag.Result = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult GenerateExcel(string startDateParam, string endDateParam, long locID, string locType)
        {
            try
            {
                // Getting all data    			
                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                if (!string.IsNullOrEmpty(startDateParam))
                {
                    startDate = DateTime.Parse(startDateParam);
                }

                if (!string.IsNullOrEmpty(endDateParam))
                {
                    endDate = DateTime.Parse(endDateParam);
                }

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", startDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));

                string empOvertimeList = _empOvertimeService.Find(filters);
                List<EmployeeOvertimeModel> empOvertimeModelList = empOvertimeList.DeserializeToEmployeeOvertimeList();

                string userList = _userAppService.FindBy("IsFast", "true");
                List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();
                List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

                empOvertimeModelList = empOvertimeModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

                // Filter Search
                if (locID > 0 && locType != null)
                {
                    List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);

                    empOvertimeModelList = empOvertimeModelList.Where(x => !string.IsNullOrEmpty(x.Location) && locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
                }

                foreach (var item in empOvertimeModelList)
                {
                    EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                    item.FullName = emp == null ? string.Empty : emp.FullName;
                }

                byte[] excelData = ExcelGenerator.ExportEmployeeOvertime(empOvertimeModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Employee-Overtime.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult GeneratePDF(string startDateParam, string endDateParam)
        {
            try
            {
                // Getting all data    			
                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                DateTime startDate = DateTime.Now.AddDays(-1);
                DateTime endDate = DateTime.Now.AddDays(30);

                if (!string.IsNullOrEmpty(startDateParam))
                {
                    startDate = DateTime.Parse(startDateParam);
                }

                if (!string.IsNullOrEmpty(endDateParam))
                {
                    endDate = DateTime.Parse(endDateParam);
                }

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", startDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                if (!string.IsNullOrEmpty(endDateParam))
                    filters.Add(new QueryFilter("Date", endDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));

                string empOvertimeList = _empOvertimeService.Find(filters);
                List<EmployeeOvertimeModel> empOvertimeModelList = empOvertimeList.DeserializeToEmployeeOvertimeList();

                string userList = _userAppService.FindBy("IsFast", "true");
                List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

                List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

                empOvertimeModelList = empOvertimeModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).OrderBy(x => x.Date).ToList();

                foreach (var item in empOvertimeModelList)
                {
                    EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                    item.FullName = emp == null ? string.Empty : emp.FullName;
                }

                Document pdfDoc = new Document(PageSize.A3.Rotate(), 10, 10, 10, 10);
                PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
                pdfDoc.Open();

                Image image = Image.GetInstance(Server.MapPath("~/Content/theme/images/fast-blue.jpg"));
                image.ScaleAbsolute(193, 38);
                pdfDoc.Add(image);

                BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.EMBEDDED);
                Font font = new Font(bf, 12);
                pdfDoc.Add(new Paragraph(new Chunk("Title          : Master Employee Overtimes", font)));
                pdfDoc.Add(new Paragraph(new Chunk("Generated By   : " + AccountName, font)));
                pdfDoc.Add(new Paragraph(new Chunk("Generated Date : " + DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), font)));

                //Horizontal Line
                Paragraph line = new Paragraph(new Chunk(new LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
                pdfDoc.Add(line);

                //Table
                PdfPTable table = new PdfPTable(12);
                table.WidthPercentage = 100;
                //0=Left, 1=Centre, 2=Right
                table.HorizontalAlignment = 0;
                table.SpacingBefore = 20f;
                table.SpacingAfter = 30f;

                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                Chunk chunk = new Chunk("Emp ID");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Emp Name");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Department");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Position");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Base Location");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Cost Center");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Date");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Clock In");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Clock Out");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Actual In");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Actual Out");
                cell.AddElement(chunk);
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                chunk = new Chunk("Overtime");
                cell.AddElement(chunk);
                table.AddCell(cell);

                foreach (var item in empOvertimeModelList)
                {
                    table.AddCell(item.EmployeeID);
                    table.AddCell(item.FullName);
                    table.AddCell(item.DepartmentDesc);
                    table.AddCell(item.PositionDesc);
                    table.AddCell(item.BasetownLocation);
                    table.AddCell(item.CostCenter);
                    table.AddCell(item.Date.ToString("dd-MMM-yy"));
                    table.AddCell(item.ClockIn);
                    table.AddCell(item.ClockOut);
                    table.AddCell(item.ActualIn);
                    table.AddCell(item.ActualOut);
                    table.AddCell(item.Overtime.ToString());
                }

                pdfDoc.Add(table);
                pdfWriter.CloseStream = false;
                pdfDoc.Close();

                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Employee-Overtime.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.Write(pdfDoc);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetAllByParam(string startDateParam, string endDateParam, long locID, string locType)
        {
            try
            {
                var draw = Request.Form.GetValues("draw").FirstOrDefault();
                var start = Request.Form.GetValues("start").FirstOrDefault();
                var length = Request.Form.GetValues("length").FirstOrDefault();
                var sortColumn = Request.Form.GetValues("columns[" + Request.Form.GetValues("order[0][column]").FirstOrDefault() + "][name]").FirstOrDefault();
                var sortColumnDir = Request.Form.GetValues("order[0][dir]").FirstOrDefault();
                var searchValue = Request.Form.GetValues("search[value]").FirstOrDefault();

                // Paging Size (10,20,50,100)    
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;


                // Getting all data    			
                DateTime startDate = DateTime.Parse(startDateParam);
                DateTime endDate = DateTime.Parse(endDateParam);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", startDate.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDate.ToString(), Operator.LessThanOrEqualTo));

                string empOvertimeList = _empOvertimeService.Find(filters);
                List<EmployeeOvertimeModel> empOvertimeModelList = empOvertimeList.DeserializeToEmployeeOvertimeList();

                if (locID > 0 && locType != null)
                {
                    List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);
                    empOvertimeModelList = empOvertimeModelList.Where(x => !string.IsNullOrEmpty(x.Location) && locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
                }

                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                string userList = _userAppService.FindBy("IsFast", "true");
                List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

                List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

                empOvertimeModelList = empOvertimeModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

                int recordsTotal = empOvertimeModelList.Count();
                sortColumn = sortColumn == "ID" ? "" : sortColumn;
                bool isLoaded = false;

                if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
                {
                    isLoaded = true;
                    Dictionary<long, string> locationMap = new Dictionary<long, string>();
                    foreach (var item in empOvertimeModelList)
                    {
                        EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                        item.FullName = emp == null ? string.Empty : emp.FullName;
                    }
                }

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    empOvertimeModelList = empOvertimeModelList.Where(m => m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                                                m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                empOvertimeModelList = empOvertimeModelList.OrderBy(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                empOvertimeModelList = empOvertimeModelList.OrderBy(x => x.FullName).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                empOvertimeModelList = empOvertimeModelList.OrderByDescending(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                empOvertimeModelList = empOvertimeModelList.OrderByDescending(x => x.FullName).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = empOvertimeModelList.Count();

                // Paging     
                var data = empOvertimeModelList.Skip(skip).Take(pageSize).ToList();

                if (!isLoaded)
                {
                    foreach (var item in data)
                    {
                        EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                        item.FullName = emp == null ? string.Empty : emp.FullName;
                    }
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<EmployeeOvertimeModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAll()
        {
            try
            {
                var draw = Request.Form.GetValues("draw").FirstOrDefault();
                var start = Request.Form.GetValues("start").FirstOrDefault();
                var length = Request.Form.GetValues("length").FirstOrDefault();
                var sortColumn = Request.Form.GetValues("columns[" + Request.Form.GetValues("order[0][column]").FirstOrDefault() + "][name]").FirstOrDefault();
                var sortColumnDir = Request.Form.GetValues("order[0][dir]").FirstOrDefault();
                var searchValue = Request.Form.GetValues("search[value]").FirstOrDefault();

                // Paging Size (10,20,50,100)    
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                // Getting all data    							
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", DateTime.Now.AddDays(-100).ToString("yyyy-MM-dd")));

                string empOvertimeList = _empOvertimeService.Find(filters);
                List<EmployeeOvertimeModel> empOvertimeModelList = empOvertimeList.DeserializeToEmployeeOvertimeList();

                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                string userList = _userAppService.FindBy("IsFast", "true");
                List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

                List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

                empOvertimeModelList = empOvertimeModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).OrderBy(x => x.Date).ToList();

                int recordsTotal = empOvertimeModelList.Count();
                sortColumn = sortColumn == "ID" ? "" : sortColumn;
                bool isLoaded = false;

                if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
                {
                    isLoaded = true;
                    Dictionary<string, string> empMap = new Dictionary<string, string>();
                    foreach (var item in empOvertimeModelList)
                    {
                        if (empMap.ContainsKey(item.EmployeeID.Trim()))
                        {
                            string name;
                            empMap.TryGetValue(item.EmployeeID.Trim(), out name);
                            item.FullName = name;
                        }
                        else
                        {
                            EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                            item.FullName = emp == null ? string.Empty : emp.FullName;
                            empMap.Add(item.EmployeeID.Trim(), item.FullName);
                        }
                    }
                }

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    empOvertimeModelList = empOvertimeModelList.Where(m => m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                                                m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                empOvertimeModelList = empOvertimeModelList.OrderBy(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                empOvertimeModelList = empOvertimeModelList.OrderBy(x => x.FullName).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                empOvertimeModelList = empOvertimeModelList.OrderByDescending(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                empOvertimeModelList = empOvertimeModelList.OrderByDescending(x => x.FullName).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = empOvertimeModelList.Count();

                // Paging     
                var data = empOvertimeModelList.Skip(skip).Take(pageSize).ToList();

                if (!isLoaded)
                {
                    foreach (var item in data)
                    {
                        EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                        item.FullName = emp == null ? string.Empty : emp.FullName;
                    }
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<EmployeeOvertimeModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult OvertimeDashboard(string location)
        {
            DateTime firstDayOfLast3Month = new DateTime(DateTime.Now.AddMonths(-3).Year, DateTime.Now.AddMonths(-3).Month, 1);
            DateTime lastDayOfLast3Month = firstDayOfLast3Month.AddMonths(1).AddDays(-1);

            DateTime firstDayOfLast2Month = new DateTime(DateTime.Now.AddMonths(-2).Year, DateTime.Now.AddMonths(-2).Month, 1);
            DateTime lastDayOfLast2Month = firstDayOfLast2Month.AddMonths(1).AddDays(-1);

            DateTime firstDayOfLastMonth = new DateTime(DateTime.Now.AddMonths(-1).Year, DateTime.Now.AddMonths(-1).Month, 1);
            DateTime lastDayOfLastMonth = firstDayOfLastMonth.AddMonths(1).AddDays(-1);

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", firstDayOfLast3Month.ToString(), Operator.GreaterThanOrEqual));
            filters.Add(new QueryFilter("Date", lastDayOfLastMonth.ToString(), Operator.LessThanOrEqualTo));

            string empOverTimeList = _empOvertimeService.Find(filters);
            List<EmployeeOvertimeModel> result = empOverTimeList.DeserializeToEmployeeOvertimeList();

            string userList = _userAppService.FindBy("IsFast", "true");
            List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();
            List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

            result = result.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

            if (location != null && location != "All")
            {
                // filter by selected PC
                result = result.Where(x => x.Location != null && x.Location.StartsWith("ID-" + location)).ToList();
            }

            OvertimeDashboardModel model = new OvertimeDashboardModel();
            model.LastMonthName = firstDayOfLastMonth.ToString("MMMM");
            model.LastTwoMonthName = firstDayOfLast2Month.ToString("MMMM");
            model.LastThreeMonthName = firstDayOfLast3Month.ToString("MMMM");

            for (var day = firstDayOfLast3Month.Date; day.Date <= lastDayOfLast3Month.Date; day = day.AddDays(1))
            {
                var overtimeList = result.Where(x => x.Date == day).ToList();

                model.LastThreeMonth.Blank += overtimeList.Where(x => string.IsNullOrEmpty(x.OvertimeCategory)).Sum(x => x.Overtime);
                model.LastThreeMonth.Rework += overtimeList.Where(x => x.OvertimeCategory == "Rework").Sum(x => x.Overtime);
                model.LastThreeMonth.Other += overtimeList.Where(x => x.OvertimeCategory == "Other").Sum(x => x.Overtime);
                model.LastThreeMonth.Daily += overtimeList.Where(x => x.OvertimeCategory == "Daily").Sum(x => x.Overtime);
                model.LastThreeMonth.BackupLeave += overtimeList.Where(x => x.OvertimeCategory == "Back Up Leave" || x.OvertimeCategory == "Backup Leave").Sum(x => x.Overtime);
                model.LastThreeMonth.Leave += overtimeList.Where(x => x.OvertimeCategory == "Leave").Sum(x => x.Overtime);
                model.LastThreeMonth.Maintenance += overtimeList.Where(x => x.OvertimeCategory == "Maintenance").Sum(x => x.Overtime);
                model.LastThreeMonth.Volume += overtimeList.Where(x => x.OvertimeCategory == "Volume").Sum(x => x.Overtime);
                model.LastThreeMonth.Training += overtimeList.Where(x => x.OvertimeCategory == "Training").Sum(x => x.Overtime);
                model.LastThreeMonth.Project += overtimeList.Where(x => x.OvertimeCategory == "Project" || x.OvertimeCategory == "Project Activity").Sum(x => x.Overtime);
            }

            model.LastThreeMonth.Blank = model.LastThreeMonth.Blank == 0 ? 0 : model.LastThreeMonth.Blank / 1000;
            model.LastThreeMonth.Rework = model.LastThreeMonth.Rework == 0 ? 0 : model.LastThreeMonth.Rework / 1000;
            model.LastThreeMonth.Other = model.LastThreeMonth.Other == 0 ? 0 : model.LastThreeMonth.Other / 1000;
            model.LastThreeMonth.Daily = model.LastThreeMonth.Daily == 0 ? 0 : model.LastThreeMonth.Daily / 1000;
            model.LastThreeMonth.BackupLeave = model.LastThreeMonth.BackupLeave == 0 ? 0 : model.LastThreeMonth.BackupLeave / 1000;
            model.LastThreeMonth.Leave = model.LastThreeMonth.Leave == 0 ? 0 : model.LastThreeMonth.Leave / 1000;
            model.LastThreeMonth.Maintenance = model.LastThreeMonth.Maintenance == 0 ? 0 : model.LastThreeMonth.Maintenance / 1000;
            model.LastThreeMonth.Volume = model.LastThreeMonth.Volume == 0 ? 0 : model.LastThreeMonth.Volume / 1000;
            model.LastThreeMonth.Training = model.LastThreeMonth.Training == 0 ? 0 : model.LastThreeMonth.Training / 1000;
            model.LastThreeMonth.Project = model.LastThreeMonth.Project == 0 ? 0 : model.LastThreeMonth.Project / 1000;

            for (var day = firstDayOfLast2Month.Date; day.Date <= lastDayOfLast2Month.Date; day = day.AddDays(1))
            {
                var overtimeList = result.Where(x => x.Date == day).ToList();

                model.LastTwoMonth.Blank += overtimeList.Where(x => string.IsNullOrEmpty(x.OvertimeCategory)).Sum(x => x.Overtime);
                model.LastTwoMonth.Rework += overtimeList.Where(x => x.OvertimeCategory == "Rework").Sum(x => x.Overtime);
                model.LastTwoMonth.Other += overtimeList.Where(x => x.OvertimeCategory == "Other").Sum(x => x.Overtime);
                model.LastTwoMonth.Daily += overtimeList.Where(x => x.OvertimeCategory == "Daily").Sum(x => x.Overtime);
                model.LastTwoMonth.BackupLeave += overtimeList.Where(x => x.OvertimeCategory == "Back Up Leave" || x.OvertimeCategory == "Backup Leave").Sum(x => x.Overtime);
                model.LastTwoMonth.Leave += overtimeList.Where(x => x.OvertimeCategory == "Leave").Sum(x => x.Overtime);
                model.LastTwoMonth.Maintenance += overtimeList.Where(x => x.OvertimeCategory == "Maintenance").Sum(x => x.Overtime);
                model.LastTwoMonth.Volume += overtimeList.Where(x => x.OvertimeCategory == "Volume").Sum(x => x.Overtime);
                model.LastTwoMonth.Training += overtimeList.Where(x => x.OvertimeCategory == "Training").Sum(x => x.Overtime);
                model.LastTwoMonth.Project += overtimeList.Where(x => x.OvertimeCategory == "Project" || x.OvertimeCategory == "Project Activity").Sum(x => x.Overtime);
            }

            model.LastTwoMonth.Blank = model.LastTwoMonth.Blank == 0 ? 0 : model.LastTwoMonth.Blank / 1000;
            model.LastTwoMonth.Rework = model.LastTwoMonth.Rework == 0 ? 0 : model.LastTwoMonth.Rework / 1000;
            model.LastTwoMonth.Other = model.LastTwoMonth.Other == 0 ? 0 : model.LastTwoMonth.Other / 1000;
            model.LastTwoMonth.Daily = model.LastTwoMonth.Daily == 0 ? 0 : model.LastTwoMonth.Daily / 1000;
            model.LastTwoMonth.BackupLeave = model.LastTwoMonth.BackupLeave == 0 ? 0 : model.LastTwoMonth.BackupLeave / 1000;
            model.LastTwoMonth.Leave = model.LastTwoMonth.Leave == 0 ? 0 : model.LastTwoMonth.Leave / 1000;
            model.LastTwoMonth.Maintenance = model.LastTwoMonth.Maintenance == 0 ? 0 : model.LastTwoMonth.Maintenance / 1000;
            model.LastTwoMonth.Volume = model.LastTwoMonth.Volume == 0 ? 0 : model.LastTwoMonth.Volume / 1000;
            model.LastTwoMonth.Training = model.LastTwoMonth.Training == 0 ? 0 : model.LastTwoMonth.Training / 1000;
            model.LastTwoMonth.Project = model.LastTwoMonth.Project == 0 ? 0 : model.LastTwoMonth.Project / 1000;

            for (var day = firstDayOfLastMonth.Date; day.Date <= lastDayOfLastMonth.Date; day = day.AddDays(1))
            {
                var overtimeList = result.Where(x => x.Date == day).ToList();

                model.LastMonth.Blank += overtimeList.Where(x => string.IsNullOrEmpty(x.OvertimeCategory)).Sum(x => x.Overtime);
                model.LastMonth.Rework += overtimeList.Where(x => x.OvertimeCategory == "Rework").Sum(x => x.Overtime);
                model.LastMonth.Other += overtimeList.Where(x => x.OvertimeCategory == "Other").Sum(x => x.Overtime);
                model.LastMonth.Daily += overtimeList.Where(x => x.OvertimeCategory == "Daily").Sum(x => x.Overtime);
                model.LastMonth.BackupLeave += overtimeList.Where(x => x.OvertimeCategory == "Back Up Leave" || x.OvertimeCategory == "Backup Leave").Sum(x => x.Overtime);
                model.LastMonth.Leave += overtimeList.Where(x => x.OvertimeCategory == "Leave").Sum(x => x.Overtime);
                model.LastMonth.Maintenance += overtimeList.Where(x => x.OvertimeCategory == "Maintenance").Sum(x => x.Overtime);
                model.LastMonth.Volume += overtimeList.Where(x => x.OvertimeCategory == "Volume").Sum(x => x.Overtime);
                model.LastMonth.Training += overtimeList.Where(x => x.OvertimeCategory == "Training").Sum(x => x.Overtime);
                model.LastMonth.Project += overtimeList.Where(x => x.OvertimeCategory == "Project" || x.OvertimeCategory == "Project Activity").Sum(x => x.Overtime);
            }

            model.LastMonth.Blank = model.LastMonth.Blank == 0 ? 0 : model.LastMonth.Blank / 1000;
            model.LastMonth.Rework = model.LastMonth.Rework == 0 ? 0 : model.LastMonth.Rework / 1000;
            model.LastMonth.Other = model.LastMonth.Other == 0 ? 0 : model.LastMonth.Other / 1000;
            model.LastMonth.Daily = model.LastMonth.Daily == 0 ? 0 : model.LastMonth.Daily / 1000;
            model.LastMonth.BackupLeave = model.LastMonth.BackupLeave == 0 ? 0 : model.LastMonth.BackupLeave / 1000;
            model.LastMonth.Leave = model.LastMonth.Leave == 0 ? 0 : model.LastMonth.Leave / 1000;
            model.LastMonth.Maintenance = model.LastMonth.Maintenance == 0 ? 0 : model.LastMonth.Maintenance / 1000;
            model.LastMonth.Volume = model.LastMonth.Volume == 0 ? 0 : model.LastMonth.Volume / 1000;
            model.LastMonth.Training = model.LastMonth.Training == 0 ? 0 : model.LastMonth.Training / 1000;
            model.LastMonth.Project = model.LastMonth.Project == 0 ? 0 : model.LastMonth.Project / 1000;

            // Returning Json Data    
            return Json(model, JsonRequestBehavior.AllowGet);
        }

        private byte[] GetFile(string filepath)
        {
            FileStream fs = System.IO.File.OpenRead(filepath);
            byte[] data = new byte[fs.Length];
            int br = fs.Read(data, 0, data.Length);
            if (br != fs.Length)
                throw new System.IO.IOException(filepath);
            return data;
        }
    }
}
