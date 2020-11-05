using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
    [CustomAuthorize("mealrequest")]
    public class MealRequestController : BaseController<MealRequestModel>
    {
        #region ::Init::
        private readonly IMealRequestAppService _mealRequestAppService;
        private readonly IMenuAppService _menuService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly ILoggerAppService _logger;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly IMppAppService _mppAppService;
        private readonly IUserAppService _userAppService;
        #endregion

        #region ::Constructor::
        public MealRequestController(
            ILoggerAppService logger,
            IMealRequestAppService mealRequestAppService,
            IReferenceAppService referenceAppService,
            IEmployeeAppService employeeAppService,
            ILocationAppService locationAppService,
            IMppAppService mppAppService,
            IUserAppService userAppService,
            IMenuAppService menuService)
        {
            _userAppService = userAppService;
            _mppAppService = mppAppService;
            _locationAppService = locationAppService;
            _employeeAppService = employeeAppService;
            _referenceAppService = referenceAppService;
            _mealRequestAppService = mealRequestAppService;
            _menuService = menuService;
            _logger = logger;
        }
        #endregion

        #region ::Public Methods::		
        public ActionResult Index()
        {
            GetTempData();

            MealRequestModel model = new MealRequestModel();
            model.Access = GetAccess(WebConstants.MenuSlug.MEAL_REQUEST, _menuService);

            return View(model);
        }

        public ActionResult DetailReport()
        {
            GetTempData();

            MealRequestModel model = new MealRequestModel();
            model.Access = GetAccess(WebConstants.MenuSlug.MEAL_REQUEST, _menuService);

            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            return View(model);
        }

        public ActionResult SummaryReport()
        {
            GetTempData();

            MealRequestModel model = new MealRequestModel();
            model.Access = GetAccess(WebConstants.MenuSlug.MEAL_REQUEST, _menuService);

            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            return View(model);
        }

        [HttpPost]
        public JsonResult AutoComplete(string prefix)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            if (prefix.All(Char.IsDigit))
                filters.Add(new QueryFilter("EmployeeID", prefix, Operator.StartsWith));
            else
                filters.Add(new QueryFilter("FullName", prefix, Operator.Contains));

            string emplist = _employeeAppService.Find(filters);
            List<EmployeeModel> empModelList = emplist.DeserializeToEmployeeList();

            if (prefix.All(Char.IsDigit))
            {
                empModelList = empModelList.OrderBy(x => x.EmployeeID).ToList();
            }
            else
            {
                empModelList = empModelList.OrderBy(x => x.FullName).ToList();
            }

            return Json(empModelList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Create()
        {
            ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService);
            ViewBag.CostCenterList = DropDownHelper.BindDropDownCostCenter(_referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, false);

            EmployeeModel emp = _employeeAppService.GetModelByEmpId(AccountEmployeeID);

            MealRequestModel model = new MealRequestModel();
            model.Company = "PT HM Sampoerna";
            model.EmployeeID = AccountEmployeeID;
            model.EmployeeFullname = emp.FullName;
            model.StartDate = DateTime.Now;
            model.EndDate = DateTime.Now;

            return PartialView(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(MealRequestModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                model.EmployeeID = string.IsNullOrEmpty(model.EmployeeID) ? AccountEmployeeID : model.EmployeeID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
                if (model.Department != "0")
                {
                    var fd = departList.Where(x => x.Value == model.Department).FirstOrDefault();
                    model.Department = fd == null ? "" : fd.Value;
                }

                List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                if (model.ProductionCenterID != "0")
                {
                    var pc = pcList.Where(x => x.Value == model.ProductionCenterID).FirstOrDefault();
                    model.ProductionCenter = pc == null ? "" : pc.Text;
                }

                List<MealRequestModel> requestList = new List<MealRequestModel>();
                for (var day = model.StartDate.Date; day.Date <= model.EndDate.Date; day = day.AddDays(1))
                {
                    MealRequestModel newModel = new MealRequestModel
                    {
                        StartDate = model.StartDate,
                        EndDate = model.EndDate,
                        Date = day,
                        Canteen = model.Canteen,
                        Company = model.Company,
                        CostCenter = model.CostCenter,
                        Department = model.Department,
                        EmployeeFullname = model.EmployeeFullname,
                        EmployeeID = model.EmployeeID,
                        Guest = model.Guest,
                        GuestType = model.GuestType,
                        ModifiedBy = model.ModifiedBy,
                        ModifiedDate = model.ModifiedDate,
                        Phone = model.Phone,
                        ProductionCenter = model.ProductionCenter,
                        ProductionCenterID = model.ProductionCenterID,
                        Purpose = model.Purpose,
                        RequestType = model.RequestType,
                        Shift = model.Shift,
                        TotalGuest = model.TotalGuest
                    };

                    requestList.Add(newModel);
                }

                string data = JsonHelper<MealRequestModel>.Serialize(requestList);
                _mealRequestAppService.AddRange(data);

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.CreateFailed);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Edit(int id)
        {
            string mealRequest = _mealRequestAppService.GetById(id);
            MealRequestModel mealRequestModel = mealRequest.DeserializeToMealRequest();

            string emp = _employeeAppService.GetBy("EmployeeID", mealRequestModel.EmployeeID);
            EmployeeModel empModel = emp.DeserializeToEmployee();

            mealRequestModel.EmployeeFullname = empModel.FullName;

            List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            if (!string.IsNullOrEmpty(mealRequestModel.ProductionCenter))
            {
                var pc = pcList.Where(x => x.Text == mealRequestModel.ProductionCenter).FirstOrDefault();
                mealRequestModel.ProductionCenterID = pc == null ? "" : pc.Value;
            }

            ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService);
            ViewBag.CostCenterList = DropDownHelper.BindDropDownCostCenter(_referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, false);

            return PartialView(mealRequestModel);
        }

        [HttpPost]
        public ActionResult Edit(MealRequestModel model)
        {
            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.TRAINING, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                model.EmployeeID = string.IsNullOrEmpty(model.EmployeeID) ? AccountEmployeeID : model.EmployeeID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
                if (model.Department != "0")
                {
                    var fd = departList.Where(x => x.Value == model.Department).FirstOrDefault();
                    model.Department = fd == null ? "" : fd.Value;
                }
                List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                if (model.ProductionCenterID != "0")
                {
                    var pc = pcList.Where(x => x.Value == model.ProductionCenterID).FirstOrDefault();
                    model.ProductionCenter = pc == null ? "" : pc.Text;
                }

                string data = JsonHelper<MealRequestModel>.Serialize(model);

                _mealRequestAppService.Update(data);

                SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                string tt = _mealRequestAppService.GetById(id, true);
                MealRequestModel ttm = tt.DeserializeToMealRequest();
                ttm.IsDeleted = true;

                string data = JsonHelper<MealRequestModel>.Serialize(ttm);
                _mealRequestAppService.Update(data);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GenerateExcel()
        {
            try
            {
                // Getting all data    			
                string mealRequests = _mealRequestAppService.GetAll(true);
                List<MealRequestModel> mealModelList = mealRequests.DeserializeToMealRequestList();

                string emps = _employeeAppService.GetAll();
                List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

                foreach (var item in mealModelList)
                {
                    var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                    item.EmployeeFullname = emp == null ? "" : emp.FullName;
                }

                byte[] excelData = ExcelGenerator.ExportMealRequest(mealModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Facility-Meal-Request.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.GenerateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID);
            }

            return RedirectToAction("Index");
        }

        public ActionResult GenerateExcelDetailReport(string startDate, string endDate, string location, string request)
        {
            try
            {
                DateTime startDateTemp = DateTime.Parse(startDate);
                DateTime endDateTemp = DateTime.Parse(endDate);

                List<MealRequestModel> mealModelList = new List<MealRequestModel>();

                Dictionary<long, string> canteenList = GetCanteenList();

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                // Getting all data    			
                if (request == "Form")
                {
                    if (location != "0")
                    {
                        List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                        if (pc != null)
                        {
                            filters.Add(new QueryFilter("ProductionCenter", pc.Text));
                        }
                    }

                    mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                    List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                    foreach (var item in mealModelList)
                    {
                        var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        item.PIC = emp == null ? "" : emp.FullName;

                        if (!string.IsNullOrEmpty(item.Department))
                        {
                            var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                            item.Department = fd == null ? "" : fd.Text;
                        }
                    }
                }
                else if (request == "MPP")
                {
                    List<EmployeeModel> empList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList().Where(x => x.SupervisorID != null).ToList();

                    List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
                        mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                        List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empList, canteenList);

                        List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empList, canteenList);

                        List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empList, canteenList);
                    }
                    else
                    {
                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        foreach (var pc in pcList)
                        {
                            if (pc.Value != "0")
                            {
                                List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
                                var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                                List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empList, canteenList);

                                List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empList, canteenList);

                                List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empList, canteenList);
                            }
                        }
                    }

                    mealModelList = mealModelList.OrderBy(x => x.Date).ToList();
                }
                else
                {

                    List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                        if (pc != null)
                        {
                            filters.Add(new QueryFilter("ProductionCenter", pc.Text));
                        }
                    }

                    mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                    List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                    foreach (var item in mealModelList)
                    {
                        var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        item.PIC = emp == null ? "" : emp.FullName;

                        if (!string.IsNullOrEmpty(item.Department))
                        {
                            var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                            item.Department = fd == null ? "" : fd.Text;
                        }
                    }

                    filters.Add(new QueryFilter("IsDeleted", "0"));
                    filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                    filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                    List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList().Where(x => x.SupervisorID != null).ToList();

                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
                        mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                        List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                        List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                        List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                    }
                    else
                    {
                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        foreach (var pc in pcList)
                        {
                            if (pc.Value != "0")
                            {
                                List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
                                var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                                List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                                List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                                List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                            }
                        }
                    }

                    mealModelList = mealModelList.OrderBy(x => x.Date).ToList();
                }

                byte[] excelData = ExcelGenerator.ExportMealRequestDetailReport(mealModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Facility-Meal-Request-Detail-Report.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.GenerateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID);
            }

            return RedirectToAction("Index");
        }

        public ActionResult GenerateExcelSummaryReport(string startDate, string endDate, string location)
        {
            try
            {
                DateTime startDateTemp = DateTime.Parse(startDate);
                DateTime endDateTemp = DateTime.Parse(endDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                // Getting all data    			
                List<MealRequestModel> mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

                #region ::Form::
                if (location != "0")
                {
                    var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                    if (pc != null)
                    {
                        mealModelList = mealModelList.Where(x => x.ProductionCenter == pc.Text).ToList();
                    }
                }

                string emps = _employeeAppService.GetAll();
                List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

                List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                foreach (var item in mealModelList)
                {
                    var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                    item.EmployeeFullname = emp == null ? "" : emp.FullName;

                    if (!string.IsNullOrEmpty(item.Department))
                    {
                        var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                        item.Department = fd == null ? "" : fd.Text;
                    }
                }

                List<MealRequestModel> result = new List<MealRequestModel>();

                foreach (var item in mealModelList)
                {
                    var temp = result.Where(x => x.ProductionCenter == item.ProductionCenter && x.Department == item.Department && x.CostCenter == item.CostCenter && x.Date.Date == item.Date.Date).FirstOrDefault();
                    if (temp == null)
                    {
                        int tempShift;
                        if (item.Shift == "NS")
                        {
                            item.NS = item.TotalGuest;
                        }
                        else if (int.TryParse(item.Shift, out tempShift))
                        {
                            if (tempShift == 1)
                            {
                                item.Shift1 = item.TotalGuest;
                            }
                            else if (tempShift == 2)
                            {
                                item.Shift2 = item.TotalGuest;
                            }
                            else
                            {
                                item.Shift3 = item.TotalGuest;
                            }
                        }

                        result.Add(item);
                    }
                    else
                    {
                        int tempShift;
                        if (item.Shift == "NS")
                        {
                            item.NS += item.TotalGuest;
                        }
                        else if (int.TryParse(item.Shift, out tempShift))
                        {
                            if (tempShift == 1)
                            {
                                temp.Shift1 += item.TotalGuest;
                            }
                            else if (tempShift == 2)
                            {
                                temp.Shift2 += item.TotalGuest;
                            }
                            else
                            {
                                temp.Shift3 += item.TotalGuest;
                            }
                        }
                    }
                }
                #endregion

                #region ::MPP::
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();
                Dictionary<long, string> canteenList = GetCanteenList();
                List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList();

                if (location == "0")
                {
                    foreach (var pc in pcList)
                    {
                        if (pc.Value == "0") continue;

                        long pcID = long.Parse(pc.Value);
                        List<long> pcLocIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                        List<MppModel> pcMppList = mppList.Where(x => pcLocIDList.Any(y => y == x.LocationID)).ToList();

                        foreach (var item in pcMppList)
                        {
                            GetCanteenAndCostCenter(canteenList, userList, empModelList, item);
                        }

                        List<string> mppCanteenList = pcMppList.Select(x => x.Canteen).Distinct().ToList();

                        foreach (var canteen in mppCanteenList)
                        {
                            var canteenMppList = pcMppList.Where(x => x.Canteen == canteen).ToList();

                            List<string> mppCostCenterList = canteenMppList.Select(x => x.CostCenter).Distinct().ToList();
                            foreach (var cc in mppCostCenterList)
                            {
                                var ccMppList = canteenMppList.Where(x => x.CostCenter == cc).ToList();
                                for (var day = startDateTemp.Date; day.Date <= endDateTemp.Date; day = day.AddDays(1))
                                {
                                    var tempList1 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "1").ToList();
                                    var tempList2 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "2").ToList();
                                    var tempList3 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "3").ToList();
                                    var tempList4 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "NS").ToList();

                                    if (tempList1.Count > 0 || tempList2.Count > 0 || tempList3.Count > 0 || tempList4.Count > 0)
                                    {
                                        MealRequestModel newSummaryModel = new MealRequestModel
                                        {
                                            ProductionCenter = pc.Text,
                                            Canteen = canteen,
                                            CostCenter = cc,
                                            Date = day,
                                            Shift1 = tempList1.Count,
                                            Shift2 = tempList2.Count,
                                            Shift3 = tempList3.Count,
                                            NS = tempList4.Count
                                        };

                                        result.Add(newSummaryModel);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    long pcID = long.Parse(location);
                    var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                    List<long> pcLocIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    List<MppModel> pcMppList = mppList.Where(x => pcLocIDList.Any(y => y == x.LocationID)).ToList();

                    foreach (var item in pcMppList)
                    {
                        GetCanteenAndCostCenter(canteenList, userList, empModelList, item);
                    }

                    List<string> mppCanteenList = pcMppList.Select(x => x.Canteen).Distinct().ToList();

                    foreach (var canteen in mppCanteenList)
                    {
                        var canteenMppList = pcMppList.Where(x => x.Canteen == canteen).ToList();

                        List<string> mppCostCenterList = canteenMppList.Select(x => x.CostCenter).Distinct().ToList();
                        foreach (var cc in mppCostCenterList)
                        {
                            var ccMppList = canteenMppList.Where(x => x.CostCenter == cc).ToList();
                            for (var day = startDateTemp.Date; day.Date <= endDateTemp.Date; day = day.AddDays(1))
                            {
                                var tempList1 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "1").ToList();
                                var tempList2 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "2").ToList();
                                var tempList3 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "3").ToList();
                                var tempList4 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "NS").ToList();

                                if (tempList1.Count > 0 || tempList2.Count > 0 || tempList3.Count > 0 || tempList4.Count > 0)
                                {
                                    MealRequestModel newSummaryModel = new MealRequestModel
                                    {
                                        ProductionCenter = pc.Text,
                                        Canteen = canteen,
                                        CostCenter = cc,
                                        Date = day,
                                        Shift1 = tempList1.Count,
                                        Shift2 = tempList2.Count,
                                        Shift3 = tempList3.Count,
                                        NS = tempList4.Count
                                    };

                                    result.Add(newSummaryModel);
                                }
                            }
                        }
                    }
                }
                #endregion

                byte[] excelData = ExcelGenerator.ExportMealRequestSummaryReport(result, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Facility-Meal-Request-Summary-Report.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.GenerateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID);
            }

            return RedirectToAction("Index");
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
                string meals = _mealRequestAppService.GetAll(true);
                List<MealRequestModel> mealModelList = meals.DeserializeToMealRequestList().OrderBy(x => x.Date).ToList();

                string emps = _employeeAppService.GetAll();
                List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

                List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                foreach (var item in mealModelList)
                {
                    var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                    item.EmployeeFullname = emp == null ? "" : emp.FullName;

                    if (!string.IsNullOrEmpty(item.Department))
                    {
                        var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                        item.Department = fd == null ? "" : fd.Text;
                    }
                }

                int recordsTotal = mealModelList.Count();

                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    mealModelList = mealModelList.Where(m => m.Canteen != null && m.Canteen.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "canteen":
                                mealModelList = mealModelList.OrderBy(x => x.Canteen).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "canteen":
                                mealModelList = mealModelList.OrderByDescending(x => x.Canteen).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = mealModelList.Count();

                // Paging     
                var data = mealModelList.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID);

                return Json(new { data = new List<MealRequestModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllDetailReport(string startDate, string endDate, string location, string request)
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

                DateTime startDateTemp = DateTime.Parse(startDate);
                DateTime endDateTemp = DateTime.Parse(endDate);

                List<MealRequestModel> mealModelList = new List<MealRequestModel>();

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                Dictionary<long, string> canteenList = GetCanteenList();

                // Getting all data    			
                if (request == "Form")
                {
                    if (location != "0")
                    {
                        List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                        if (pc != null)
                        {
                            filters.Add(new QueryFilter("ProductionCenter", pc.Text));
                        }
                    }

                    mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                    List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                    foreach (var item in mealModelList)
                    {
                        var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        item.PIC = emp == null ? "" : emp.FullName;
                        item.CostCenter = emp == null ? "" : emp.CostCenter;

                        if (!string.IsNullOrEmpty(item.Department))
                        {
                            var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                            item.Department = fd == null ? "" : fd.Text;
                        }
                    }
                }
                else if (request == "MPP")
                {
                    List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList().Where(x => x.SupervisorID != null).ToList();

                    List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
                        mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                        List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                        List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                        List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                    }
                    else
                    {
                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        foreach (var pc in pcList)
                        {
                            if (pc.Value != "0")
                            {
                                List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
                                var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                                List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                                List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                                List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                            }
                        }
                    }

                    mealModelList = mealModelList.OrderBy(x => x.Date).ToList();
                }
                else
                {
                    List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                        if (pc != null)
                        {
                            filters.Add(new QueryFilter("ProductionCenter", pc.Text));
                        }
                    }

                    mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                    List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
                    List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                    foreach (var item in mealModelList)
                    {
                        var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        item.PIC = emp == null ? "" : emp.FullName;
                        item.CostCenter = emp == null ? "" : emp.CostCenter;

                        if (!string.IsNullOrEmpty(item.Department))
                        {
                            var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                            item.Department = fd == null ? "" : fd.Text;
                        }
                        else
                        {
                            item.Department = emp == null ? "" : emp.DepartmentDesc;
                        }
                    }

                    filters.Add(new QueryFilter("IsDeleted", "0"));
                    filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                    filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                    List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList().Where(x => x.SupervisorID != null).ToList();

                    if (location != "0")
                    {
                        var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
                        mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                        List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                        List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                        List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
                        GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                    }
                    else
                    {
                        List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

                        foreach (var pc in pcList)
                        {
                            if (pc.Value != "0")
                            {
                                List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
                                var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                                List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift1, "1", userList, empModelList, canteenList);

                                List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift2, "2", userList, empModelList, canteenList);

                                List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
                                GetMppDataList(startDateTemp, endDateTemp, mealModelList, pc.Text, mppListShift3, "3", userList, empModelList, canteenList);
                            }
                        }
                    }

                    mealModelList = mealModelList.OrderBy(x => x.Date).ToList();
                }

                int recordsTotal = mealModelList.Count();

                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    mealModelList = mealModelList.Where(m => m.Canteen != null && m.Canteen.ToLower().Contains(searchValue.ToLower()) ||
                                                        m.Guest != null && m.Guest.ToLower().Contains(searchValue.ToLower()) ||
                                                        m.CostCenter != null && m.CostCenter.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "productioncenter":
                                mealModelList = mealModelList.OrderBy(x => x.ProductionCenter).ToList();
                                break;
                            case "department":
                                mealModelList = mealModelList.OrderBy(x => x.Department).ToList();
                                break;
                            case "costcenter":
                                mealModelList = mealModelList.OrderBy(x => x.CostCenter).ToList();
                                break;
                            case "startdate":
                                mealModelList = mealModelList.OrderBy(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                mealModelList = mealModelList.OrderBy(x => x.EndDate).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "productioncenter":
                                mealModelList = mealModelList.OrderByDescending(x => x.ProductionCenter).ToList();
                                break;
                            case "department":
                                mealModelList = mealModelList.OrderByDescending(x => x.Department).ToList();
                                break;
                            case "costcenter":
                                mealModelList = mealModelList.OrderByDescending(x => x.CostCenter).ToList();
                                break;
                            case "startdate":
                                mealModelList = mealModelList.OrderByDescending(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                mealModelList = mealModelList.OrderByDescending(x => x.EndDate).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = mealModelList.Count();

                // Paging     
                var data = mealModelList.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID);

                return Json(new { data = new List<MealRequestModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllSummaryReport(string startDate, string endDate, string location)
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

                DateTime startDateTemp = DateTime.Parse(startDate);
                DateTime endDateTemp = DateTime.Parse(endDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                // Getting all data    			
                List<MealRequestModel> mealModelList = _mealRequestAppService.Find(filters).DeserializeToMealRequestList();

                List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

                #region ::Form::
                if (location != "0")
                {
                    var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
                    if (pc != null)
                    {
                        mealModelList = mealModelList.Where(x => x.ProductionCenter == pc.Text).ToList();
                    }
                }

                string emps = _employeeAppService.GetAll();
                List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

                List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

                foreach (var item in mealModelList)
                {
                    var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                    item.EmployeeFullname = emp == null ? "" : emp.FullName;

                    if (!string.IsNullOrEmpty(item.Department))
                    {
                        var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
                        item.Department = fd == null ? "" : fd.Text;
                    }
                }

                List<MealRequestModel> result = new List<MealRequestModel>();

                foreach (var item in mealModelList)
                {
                    var temp = result.Where(x => x.ProductionCenter == item.ProductionCenter && x.Department == item.Department && x.CostCenter == item.CostCenter && x.Date.Date == item.Date.Date).FirstOrDefault();
                    if (temp == null)
                    {
                        int tempShift;
                        if (item.Shift == "NS")
                        {
                            item.NS = item.TotalGuest;
                        }
                        else if (int.TryParse(item.Shift, out tempShift))
                        {
                            if (tempShift == 1)
                            {
                                item.Shift1 = item.TotalGuest;
                            }
                            else if (tempShift == 2)
                            {
                                item.Shift2 = item.TotalGuest;
                            }
                            else
                            {
                                item.Shift3 = item.TotalGuest;
                            }
                        }

                        result.Add(item);
                    }
                    else
                    {
                        int tempShift;
                        if (item.Shift == "NS")
                        {
                            item.NS += item.TotalGuest;
                        }
                        else if (int.TryParse(item.Shift, out tempShift))
                        {
                            if (tempShift == 1)
                            {
                                temp.Shift1 += item.TotalGuest;
                            }
                            else if (tempShift == 2)
                            {
                                temp.Shift2 += item.TotalGuest;
                            }
                            else
                            {
                                temp.Shift3 += item.TotalGuest;
                            }
                        }
                    }
                }
                #endregion

                #region ::MPP::
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

                List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();
                Dictionary<long, string> canteenList = GetCanteenList();
                List<UserModel> userList = _userAppService.GetAll(true).DeserializeToUserList();

                if (location == "0")
                {
                    foreach (var pc in pcList)
                    {
                        if (pc.Value == "0") continue;

                        long pcID = long.Parse(pc.Value);
                        List<long> pcLocIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                        List<MppModel> pcMppList = mppList.Where(x => pcLocIDList.Any(y => y == x.LocationID)).ToList();

                        foreach (var item in pcMppList)
                        {
                            GetCanteenAndCostCenter(canteenList, userList, empModelList, item);
                        }

                        List<string> mppCanteenList = pcMppList.Select(x => x.Canteen).Distinct().ToList();

                        foreach (var canteen in mppCanteenList)
                        {
                            var canteenMppList = pcMppList.Where(x => x.Canteen == canteen).ToList();

                            List<string> mppCostCenterList = canteenMppList.Select(x => x.CostCenter).Distinct().ToList();
                            foreach (var cc in mppCostCenterList)
                            {
                                var ccMppList = canteenMppList.Where(x => x.CostCenter == cc).ToList();
                                for (var day = startDateTemp.Date; day.Date <= endDateTemp.Date; day = day.AddDays(1))
                                {
                                    var tempList1 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "1").ToList();
                                    var tempList2 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "2").ToList();
                                    var tempList3 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "3").ToList();
                                    var tempList4 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "NS").ToList();

                                    if (tempList1.Count > 0 || tempList2.Count > 0 || tempList3.Count > 0 || tempList4.Count > 0)
                                    {
                                        MealRequestModel newSummaryModel = new MealRequestModel
                                        {
                                            ProductionCenter = pc.Text,
                                            Canteen = canteen,
                                            CostCenter = cc,
                                            Date = day,
                                            Shift1 = tempList1.Count,
                                            Shift2 = tempList2.Count,
                                            Shift3 = tempList3.Count,
                                            NS = tempList4.Count
                                        };

                                        result.Add(newSummaryModel);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    long pcID = long.Parse(location);
                    var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

                    List<long> pcLocIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    List<MppModel> pcMppList = mppList.Where(x => pcLocIDList.Any(y => y == x.LocationID)).ToList();

                    foreach (var item in pcMppList)
                    {
                        GetCanteenAndCostCenter(canteenList, userList, empModelList, item);
                    }

                    List<string> mppCanteenList = pcMppList.Select(x => x.Canteen).Distinct().ToList();

                    foreach (var canteen in mppCanteenList)
                    {
                        var canteenMppList = pcMppList.Where(x => x.Canteen == canteen).ToList();

                        List<string> mppCostCenterList = canteenMppList.Select(x => x.CostCenter).Distinct().ToList();
                        foreach (var cc in mppCostCenterList)
                        {
                            var ccMppList = canteenMppList.Where(x => x.CostCenter == cc).ToList();
                            for (var day = startDateTemp.Date; day.Date <= endDateTemp.Date; day = day.AddDays(1))
                            {
                                var tempList1 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "1").ToList();
                                var tempList2 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "2").ToList();
                                var tempList3 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "3").ToList();
                                var tempList4 = ccMppList.Where(x => x.Date.Date == day.Date && x.Shift.Trim() == "NS").ToList();

                                if (tempList1.Count > 0 || tempList2.Count > 0 || tempList3.Count > 0 || tempList4.Count > 0)
                                {
                                    MealRequestModel newSummaryModel = new MealRequestModel
                                    {
                                        ProductionCenter = pc.Text,
                                        Canteen = canteen,
                                        CostCenter = cc,
                                        Date = day,
                                        Shift1 = tempList1.Count,
                                        Shift2 = tempList2.Count,
                                        Shift3 = tempList3.Count,
                                        NS = tempList4.Count
                                    };

                                    result.Add(newSummaryModel);
                                }
                            }
                        }
                    }
                }
                #endregion

                int recordsTotal = result.Count();

                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => m.Canteen != null && m.Canteen.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "productioncenter":
                                result = result.OrderBy(x => x.ProductionCenter).ToList();
                                break;
                            case "department":
                                result = result.OrderBy(x => x.Department).ToList();
                                break;
                            case "costcenter":
                                result = result.OrderBy(x => x.CostCenter).ToList();
                                break;
                            case "startdate":
                                result = result.OrderBy(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                result = result.OrderBy(x => x.EndDate).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "productioncenter":
                                result = result.OrderByDescending(x => x.ProductionCenter).ToList();
                                break;
                            case "department":
                                result = result.OrderByDescending(x => x.Department).ToList();
                                break;
                            case "costcenter":
                                result = result.OrderByDescending(x => x.CostCenter).ToList();
                                break;
                            case "startdate":
                                result = result.OrderByDescending(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                result = result.OrderByDescending(x => x.EndDate).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = result.Count();

                // Paging     
                var data = result.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID);

                return Json(new { data = new List<MealRequestModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region ::Private Methods::
        private void GetCanteenAndCostCenter(Dictionary<long, string> canteenList, List<UserModel> userList, List<EmployeeModel> empList, MppModel item)
        {
            var user = userList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
            var emp = empList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();

            string canteenCode = "-";
            if (user != null && user.CanteenID.HasValue && canteenList.ContainsKey(user.CanteenID.Value))
            {
                canteenList.TryGetValue(user.CanteenID.Value, out canteenCode);
            }

            item.Canteen = canteenCode;
            item.CostCenter = emp == null ? "-" : emp.CostCenter;
        }

        private Dictionary<long, string> GetCanteenList()
        {
            string reference = _referenceAppService.GetBy("Name", "Canteen", true);
            ReferenceModel refModel = reference.DeserializeToReference();

            string canteens = _referenceAppService.FindDetailBy("ReferenceID", refModel.ID, true);
            List<ReferenceDetailModel> canteenModelList = canteens.DeserializeToRefDetailList();

            Dictionary<long, string> result = new Dictionary<long, string>();

            foreach (var item in canteenModelList)
            {
                result.Add(item.ID, item.Code);
            }

            return result;
        }

        private void GetMppDataList(DateTime startDateTemp, DateTime endDateTemp, List<MealRequestModel> mealModelList, string prodcenter, List<MppModel> mppList, string shift, List<UserModel> userList, List<EmployeeModel> empList, Dictionary<long, string> canteenList)
        {
            foreach (var item in mppList)
            {
                if (mealModelList.Any(x => x.EmployeeID == item.EmployeeID && x.Date == item.Date))
                    continue;

                var user = userList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                var emp = empList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();

                string canteenCode = "-";
                if (user != null && user.CanteenID.HasValue && canteenList.ContainsKey(user.CanteenID.Value))
                {
                    canteenList.TryGetValue(user.CanteenID.Value, out canteenCode);
                }

                MealRequestModel newModel = new MealRequestModel
                {
                    EmployeeID = item.EmployeeID,
                    StartDate = startDateTemp,
                    EndDate = endDateTemp,
                    Date = item.Date,
                    Company = "HMS",
                    Canteen = canteenCode,
                    Guest = item.EmployeeName,
                    GuestType = "Internal-MPP",
                    Purpose = "Daily Meal",
                    Shift = shift,
                    ProductionCenter = prodcenter,
                    TotalGuest = 1,
                    PIC = user == null ? "-" : user.SupervisorName,
                    CostCenter = emp == null ? "-" : emp.CostCenter,
                    Department = emp == null ? "-" : emp.DepartmentDesc
                };

                mealModelList.Add(newModel);
            }
        }
        #endregion
    }
}
