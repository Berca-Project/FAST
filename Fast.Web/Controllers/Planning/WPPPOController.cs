using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Mvc;

namespace Fast.Web.Controllers
{
    [CustomAuthorize("wpp")]
    public class WppPOController : BaseController<WppModel>
    {
        private readonly IWppAppService _wppAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBlendAppService _blendAppService;
        private readonly IBrandAppService _brandAppService;
        private readonly IMaterialCodeAppService _materialAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IMenuAppService _menuAppService;
        private readonly IWppChangesAppService _wppChangesAppService;
        private readonly IWeeksAppService _weeksAppService;
        private readonly ICalendarHolidayAppService _calendarHolidayAppService;

        public WppPOController(
            IBlendAppService blendAppService,
            IBrandAppService brandAppService,
            IMaterialCodeAppService materialAppService,
            IWppAppService wppAppService,
            IReferenceAppService referenceAppService,
            ILocationAppService locationAppService,
            IMachineAppService machineAppService,
            IMenuAppService menuAppService,
            ILoggerAppService logger,
            IWppChangesAppService wppChangesAppService,
            IWeeksAppService weeksAppService,
            ICalendarHolidayAppService calendarHolidayAppService)
        {
            _referenceAppService = referenceAppService;
            _wppAppService = wppAppService;
            _logger = logger;
            _menuAppService = menuAppService;
            _locationAppService = locationAppService;
            _machineAppService = machineAppService;
            _wppChangesAppService = wppChangesAppService;
            _weeksAppService = weeksAppService;
            _calendarHolidayAppService = calendarHolidayAppService;
            _blendAppService = blendAppService;
            _brandAppService = brandAppService;
            _materialAppService = materialAppService;
        }

        [HttpPost]
        public ActionResult GetProductionCenterByCountryID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, id);

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetDepartmentByProdCenterID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, id);

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetSubDepartmentByDepartmentID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, id);
            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Index()
        {
            GetTempData();

            ViewBag.MachineList = DropDownHelper.BindDropDownMachineCode(_machineAppService);
            ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
            ViewBag.BrandList = DropDownHelper.BindDropDownBrand(_referenceAppService);
            ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

            WppModel model = new WppModel();
            model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(WppModel model)
        {
            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

                ViewBag.MachineList = DropDownHelper.BindDropDownMachineCode(_machineAppService);
                ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
                ViewBag.BrandList = DropDownHelper.BindDropDownBrand(_referenceAppService);
                ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

                ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
                ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(ModelState.GetModelStateErrors());
                    return RedirectToAction("Index");
                }

                if (model.PostedFilename != null && model.PostedFilename.ContentLength > 0)
                {
                    Stream stream = model.PostedFilename.InputStream;

                    IExcelDataReader reader = null;

                    if (model.PostedFilename.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (model.PostedFilename.FileName.EndsWith(".xlsx"))
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

                    for (int i = 0; i < dt_.Columns.Count; i++)
                    {
                        string temp = dt_.Rows[0][i].ToString();
                        if (i > 6)
                        {
                            DateTime tempDate;
                            if (DateTime.TryParseExact(temp, "yyyyMMdd", null, DateTimeStyles.None, out tempDate))
                            {
                                if (tempDate <= DateTime.Now.Date)
                                {
                                    SetFalseTempData(string.Format(UIResources.DateWppLessThanEqualToday, tempDate.ToString("yyyyMMdd"), DateTime.Now.ToString("yyyyMMdd")));
                                    return RedirectToAction("Index");
                                }
                            }
                            else
                            {
                                SetFalseTempData(string.Format(UIResources.DateWppInvalid, temp));
                                return RedirectToAction("Index");
                            }

                            temp += "-" + dt_.Rows[1][i].ToString();
                        }

                        dt.Columns.Add(temp);
                    }

                    List<WppModel> result = new List<WppModel>();
                    string location = _locationAppService.GetLocationFullCode(AccountLocationID);

                    if (dt.Columns.Count > 7)
                    {
                        for (int index = 2; index < dt_.Rows.Count; index++)
                        {
                            for (int col = 7; col < dt_.Columns.Count; col++)
                            {
                                WppModel newWpp = new WppModel();
                                newWpp.Location = location;
                                newWpp.Brand = dt_.Rows[index][0].ToString();

                                if (!string.IsNullOrEmpty(newWpp.Brand))
                                {
                                    bool brandCheck = IsBrandExist(newWpp.Brand);
                                    if (brandCheck == false)
                                    {
                                        bool blendCheck = IsBlendExist(newWpp.Brand);
                                        if (blendCheck == false)
                                        {
                                            bool materialCheck = IsMaterialCodeExist(newWpp.Brand);
                                            if (materialCheck == false)
                                            {
                                                SetFalseTempData(UIResources.BrandBlendNotExist + " Please check on cell " + "[" + index + "]" + "[0]");
                                                return RedirectToAction("Index");
                                            }
                                        }
                                    }
                                }

                                newWpp.Description = dt_.Rows[index][1].ToString();
                                newWpp.Maker = dt_.Rows[index][2].ToString();
                                newWpp.Packer = dt_.Rows[index][3].ToString();
                                newWpp.Activity = dt_.Rows[index][4].ToString();
                                newWpp.PONumber = dt_.Rows[index][5].ToString();
                                newWpp.OPSNumber = dt_.Rows[index][6].ToString();

                                if (string.IsNullOrEmpty(newWpp.Maker))
                                {
                                    if (string.IsNullOrEmpty(newWpp.Packer))
                                    {
                                        if (string.IsNullOrEmpty(newWpp.Activity))
                                        {
                                            SetFalseTempData(UIResources.FieldEmpty + " Please check on these field.(Maker/Packer/Activity)");
                                            return RedirectToAction("Index");
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(newWpp.Maker))
                                {
                                    bool checkMaker = IsMachineExist(newWpp.Maker);
                                    if (checkMaker == false)
                                    {
                                        SetFalseTempData(string.Format(UIResources.MachineCodeNotExist, newWpp.Packer) + " Please check on cell " + "[" + index + "]" + "[2]");
                                        return RedirectToAction("Index");
                                    }
                                }

                                if (!string.IsNullOrEmpty(newWpp.Packer))
                                {
                                    bool checkPacker = IsMachineExist(newWpp.Packer);
                                    if (checkPacker == false)
                                    {
                                        SetFalseTempData(string.Format(UIResources.MachineCodeNotExist, newWpp.Packer) + " Please check on cell " + "[" + index + "]" + "[3]");
                                        return RedirectToAction("Index");
                                    }
                                }

                                newWpp.LocationID = AccountLocationID;
                                newWpp.ModifiedBy = AccountName;
                                newWpp.ModifiedDate = DateTime.Now;

                                for (int shift = 1; shift <= 3; shift++, col++)
                                {
                                    string[] headers = dt.Columns[col].ToString().Split('-');
                                    newWpp.Date = DateTime.ParseExact(headers[0], "yyyyMMdd", CultureInfo.InvariantCulture);
                                    string amount = dt_.Rows[index][col].ToString();
                                    decimal amountWpp;

                                    if (decimal.TryParse(amount, out amountWpp))
                                    {
                                        if (shift == 1)
                                        {
                                            if (amountWpp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWpp.Shift1 = amountWpp;
                                        }
                                        else if (shift == 2)
                                        {
                                            if (amountWpp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWpp.Shift2 = amountWpp;
                                        }
                                        else
                                        {
                                            if (amountWpp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWpp.Shift3 = amountWpp;
                                        }
                                    }
                                }

                                // set the column pointer
                                col--;

                                if (!string.IsNullOrEmpty(newWpp.Activity) && newWpp.Shift1 == 0 && newWpp.Shift2 == 0 && newWpp.Shift3 == 0)
                                {
                                    continue;
                                }
                                else
                                {
                                    result.Add(newWpp);
                                }
                            }
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    foreach (var item in result)
                    {
                        // get the existing record
                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("Location", item.Location));
                        filters.Add(new QueryFilter("Brand", item.Brand));
                        filters.Add(new QueryFilter("Date", item.Date.ToString()));

                        string oldData = _wppAppService.Get(filters, true);
                        if (string.IsNullOrEmpty(oldData))
                        {
                            string wpp = JsonHelper<WppModel>.Serialize(item);
                            _wppAppService.Add(wpp);
                        }
                        else
                        {
                            WppModel oldDataModel = oldData.DeserializeToWpp();
                            if (oldDataModel != null)
                            {
                                oldDataModel.Shift1 = item.Shift1;
                                oldDataModel.Shift2 = item.Shift2;
                                oldDataModel.Shift3 = item.Shift3;

                                string newData = JsonHelper<WppModel>.Serialize(oldDataModel);

                                _wppAppService.Update(newData);
                            }
                        }
                    }

                    SetTrueTempData(UIResources.CreateSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.CreateFailed);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Report()
        {
            return View();
        }

        private string GetShiftData(WppDBModel newData)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", newData.Date.ToString()));
            filters.Add(new QueryFilter("Location", newData.Location));
            filters.Add(new QueryFilter("Packer", newData.Packer));
            filters.Add(new QueryFilter("Maker", newData.Maker));
            filters.Add(new QueryFilter("Brand", newData.Brand));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string exist = _wppAppService.Get(filters, true);

            return exist;
        }

        // GET: Wpp/Edit/5
        public ActionResult Edit(int id)
        {
            ViewBag.MachineList = DropDownHelper.BindDropDownMachineMakerCode(_machineAppService);
            ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
            ViewBag.BrandList = DropDownHelper.BindDropDownBrandBlend(_brandAppService, _blendAppService);
            ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

            WppModel model = GetWpp(id);
            if (!string.IsNullOrEmpty(model.Location))
            {
                int index = 1;
                string[] temp = model.Location.Split('-');
                foreach (var code in temp)
                {
                    if (index == 1)
                    {
                        string locs = _locationAppService.FindByNoTracking("Code", code, true);
                        List<LocationModel> locModelList = locs.DeserializeToLocationList();
                        if (locModelList.Count > 0)
                        {
                            var locModel = locModelList.Where(x => x.ParentCode == null).FirstOrDefault();
                            model.CountryID = locModel.ID;
                        }
                    }
                    else if (index == 2)
                    {
                        string locs = _locationAppService.FindByNoTracking("Code", code, true);
                        List<LocationModel> locModelList = locs.DeserializeToLocationList();
                        if (locModelList.Count > 0)
                        {
                            var locModel = locModelList.Where(x => x.ParentCode == temp[0]).FirstOrDefault();
                            model.ProdCenterID = locModel.ID;
                        }
                    }
                    else if (index == 3)
                    {
                        string locs = _locationAppService.FindByNoTracking("Code", code, true);
                        List<LocationModel> locModelList = locs.DeserializeToLocationList();
                        if (locModelList.Count > 0)
                        {
                            var locModel = locModelList.Where(x => x.ParentCode == temp[1]).FirstOrDefault();
                            model.DepartmentID = locModel.ID;
                        }
                    }
                    else if (index == 4)
                    {
                        string locs = _locationAppService.FindByNoTracking("Code", code, true);
                        List<LocationModel> locModelList = locs.DeserializeToLocationList();
                        if (locModelList.Count > 0)
                        {
                            var locModel = locModelList.Where(x => x.ParentCode == temp[2]).FirstOrDefault();
                            model.SubDepartmentID = locModel.ID;
                        }
                    }

                    index++;
                }
            }

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);

            if (model.CountryID == 0)
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            else
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, model.CountryID);

            if (model.ProdCenterID == 0)
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            else
                ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, model.ProdCenterID);

            if (model.DepartmentID == 0)
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
            else
                ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, model.DepartmentID);

            return PartialView(model);
        }

        private WppModel GetWpp(int id)
        {
            string wpp = _wppAppService.GetById(id, true);
            WppModel wppModel = wpp.DeserializeToWpp();

            return wppModel;
        }

        [HttpPost]
        public ActionResult Edit(WppModel model)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BuildEmptyList();
                ViewBag.LinkUpList = DropDownHelper.BuildEmptyList();
                ViewBag.BrandList = DropDownHelper.BuildEmptyList();
                ViewBag.WppTypeList = DropDownHelper.BuildEmptyList();

                ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

                model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(ModelState.GetModelStateErrors());

                    return RedirectToAction("Index");
                }

                WppModel oldModel = GetWpp(Convert.ToInt32(model.ID));

                // put everything updatable here
                oldModel.Date = model.Date;
                oldModel.PONumber = model.PONumber;
                oldModel.ModifiedBy = AccountName;
                oldModel.ModifiedDate = DateTime.Now;

                string data = JsonHelper<WppModel>.Serialize(oldModel);
                _wppAppService.Update(data);

                SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.GetAllMessages());

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult Delete(int id)
        {
            try
            {
                _wppAppService.Remove(id);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
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
                string wppList = _wppAppService.GetAll(true);
                List<WppModel> wpps = wppList.DeserializeToWppList().OrderBy(x => x.Date).ToList();
                int recordsTotal = wpps.Count();

                // Search    - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    wpps = wpps.Where(m => (m.PONumber != null ? m.PONumber.ToLower().Contains(searchValue.ToLower()) : false) || (m.Brand != null ? m.Brand.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Maker != null ? m.Maker.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Packer != null ? m.Packer.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Activity != null ? m.Activity.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Description != null ? m.Description.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "brand":
                                wpps = wpps.OrderBy(x => x.Brand).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderBy(x => x.Location).ToList();
                                break;
                            case "description":
                                wpps = wpps.OrderBy(x => x.Description).ToList();
                                break;
                            case "packer":
                                wpps = wpps.OrderBy(x => x.Packer).ToList();
                                break;
                            case "maker":
                                wpps = wpps.OrderBy(x => x.Maker).ToList();
                                break;
                            case "date":
                                wpps = wpps.OrderBy(x => x.Date).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "brand":
                                wpps = wpps.OrderByDescending(x => x.Brand).ToList();
                                break;
                            case "description":
                                wpps = wpps.OrderByDescending(x => x.Description).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderByDescending(x => x.Location).ToList();
                                break;
                            case "packer":
                                wpps = wpps.OrderByDescending(x => x.Packer).ToList();
                                break;
                            case "maker":
                                wpps = wpps.OrderByDescending(x => x.Maker).ToList();
                                break;
                            case "date":
                                wpps = wpps.OrderByDescending(x => x.Date).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = wpps.Count();

                // Paging     
                var data = wpps.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<WppModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllWPPWithParam(string dateFilter, long locID, string locType)
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

            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);

            // Getting all data    			
            string wppList = _wppAppService.GetAll(true);
            List<WppModel> wpps = wppList.DeserializeToWppList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
            int recordsTotal = wpps.Count();

            // Filter Search
            if (dateFilter != "")
            {
                DateTime dateFL = DateTime.Parse(dateFilter);
                wpps = wpps.Where(x => x.Date == dateFL.Date).ToList();
            }

            // Search    - Correction 231019
            if (!string.IsNullOrEmpty(searchValue))
            {
                wpps = wpps.Where(m => (m.PONumber != null ? m.PONumber.ToLower().Contains(searchValue.ToLower()) : false) || (m.Brand != null ? m.Brand.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Maker != null ? m.Maker.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Packer != null ? m.Packer.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Activity != null ? m.Activity.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Description != null ? m.Description.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "brand":
                            wpps = wpps.OrderBy(x => x.Brand).ToList();
                            break;
                        case "location":
                            wpps = wpps.OrderBy(x => x.Location).ToList();
                            break;
                        case "description":
                            wpps = wpps.OrderBy(x => x.Description).ToList();
                            break;
                        case "packer":
                            wpps = wpps.OrderBy(x => x.Packer).ToList();
                            break;
                        case "maker":
                            wpps = wpps.OrderBy(x => x.Maker).ToList();
                            break;
                        case "date":
                            wpps = wpps.OrderBy(x => x.Date).ToList();
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (sortColumn.ToLower())
                    {
                        case "brand":
                            wpps = wpps.OrderByDescending(x => x.Brand).ToList();
                            break;
                        case "description":
                            wpps = wpps.OrderByDescending(x => x.Description).ToList();
                            break;
                        case "location":
                            wpps = wpps.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "packer":
                            wpps = wpps.OrderByDescending(x => x.Packer).ToList();
                            break;
                        case "maker":
                            wpps = wpps.OrderByDescending(x => x.Maker).ToList();
                            break;
                        case "date":
                            wpps = wpps.OrderByDescending(x => x.Date).ToList();
                            break;
                        default:
                            break;
                    }
                }
            }

            // total number of rows count     
            int recordsFiltered = wpps.Count();

            // Paging     
            var data = wpps.Skip(skip).Take(pageSize).ToList();

            // Returning Json Data    
            return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DownloadTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateWpp.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateWpp.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        #region Helper
        private byte[] GetFile(string filepath)
        {
            FileStream fs = System.IO.File.OpenRead(filepath);
            byte[] data = new byte[fs.Length];
            int br = fs.Read(data, 0, data.Length);
            if (br != fs.Length)
                throw new System.IO.IOException(filepath);
            return data;
        }

        public long GetSubDepartmentID(string name)
        {
            var data = DropDownHelper.BindDropDownSubDepartmentCode(_referenceAppService, _locationAppService);
            long idSubDept = 0;
            foreach (var item in data)
            {
                if (item.Text == name)
                {
                    idSubDept = Convert.ToInt64(item.Value);
                    break;
                }
            }
            return idSubDept;
        }

        public string GetSubDepartmentNameByID(long subDeptID)
        {
            var data = DropDownHelper.BindDropDownSubDepartmentCode(_referenceAppService, _locationAppService);
            string resName = "";

            foreach (var item in data)
            {
                if (item.Value == subDeptID.ToString())
                {
                    resName = item.Text;
                    break;
                }
            }
            return resName;
        }

        public bool IsBrandExist(string brand)
        {
            bool isExist = true;
            string dataBrand = _referenceAppService.GetDetailAll(ReferenceEnum.Brand);
            List<ReferenceDetailModel> brandList = dataBrand.DeserializeToRefDetailList();
            var result = brandList.Where(x => x.Code == brand);
            if (result.Count() > 0)
            {
                isExist = true;
            }
            else
            {
                string brands = _brandAppService.GetBy("Code", brand);
                isExist = !string.IsNullOrEmpty(brands);
            }
            return isExist;
        }

        public bool IsBlendExist(string blend)
        {
            bool isExist = true;
            string dataBlend = _referenceAppService.GetDetailAll(ReferenceEnum.Blend);
            List<ReferenceDetailModel> blendList = dataBlend.DeserializeToRefDetailList();
            var result = blendList.Where(x => x.Code == blend);
            if (result.Count() > 0)
            {
                isExist = true;
            }
            else
            {
                string blends = _blendAppService.GetBy("Code", blend);
                isExist = !string.IsNullOrEmpty(blends);
            }

            return isExist;
        }

        public bool IsMaterialCodeExist(string material)
        {
            string blends = _materialAppService.GetBy("Code", material);
            return !string.IsNullOrEmpty(blends);
        }

        public bool IsMachineExist(string machineCode)
        {
            bool isExist = true;
            long locationId = AccountLocationID;

            string machineList = _machineAppService.GetAll(true);
            List<MachineModel> machineModelList = machineList.DeserializeToMachineList();
            var result = machineModelList.Where(x => x.Code.ToLower() == machineCode.ToLower());

            if (result.Count() > 0)
            {
                isExist = true;
            }
            else
            {
                isExist = false;
            }

            return isExist;

        }

        #endregion
    }
}
