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
    public class WppController : BaseController<WppStpModel>
    {
        private readonly IWppStpAppService _wppStpAppService;
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

        public WppController(
            IBlendAppService blendAppService,
            IBrandAppService brandAppService,
            IMaterialCodeAppService materialAppService,
            IWppStpAppService wppStpAppService,
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
            _wppStpAppService = wppStpAppService;
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

            WppStpModel model = new WppStpModel();
            model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(WppStpModel model)
        {

            IExcelDataReader reader = null;

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

                    int dayNumber = 0;
                    int weekNumber = 0;
                    for (int i = 0; i < dt_.Columns.Count; i++)
                    {
                        string temp = dt_.Rows[0][i].ToString();
                        if (i > 8)
                        {
                            dayNumber++;
                            if (weekNumber > 0)
                                temp = model.StartDate.Value.AddDays(weekNumber).ToString("yyyyMMdd");
                            else
                                temp = model.StartDate.Value.ToString("yyyyMMdd");

                            if (dayNumber > 3)
                            {
                                dayNumber = 1;
                                weekNumber++;
                                temp = model.StartDate.Value.AddDays(weekNumber).ToString("yyyyMMdd");
                            }

                            temp += "-" + dt_.Rows[1][i].ToString();
                        }

                        dt.Columns.Add(temp);
                    }

                    List<WppStpModel> result = new List<WppStpModel>();
                    string description = "";

                    if (dt.Columns.Count > 9)
                    {
                        long locationID = model.ProdCenterID == 0 ? AccountLocationID : model.ProdCenterID;
                        string location = _locationAppService.GetLocationFullCode(locationID);
                        string machineList = _machineAppService.GetAll(true);
                        List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

                        for (int index = 2; index < dt_.Rows.Count; index++)
                        {
                            for (int col = 9; col < dt_.Columns.Count; col++)
                            {
                                WppStpModel newWppStp = new WppStpModel();
                                newWppStp.Location = location;
                                newWppStp.Brand = dt_.Rows[index][0].ToString();
                                newWppStp.Description = dt_.Rows[index][1].ToString();

                                if (!string.IsNullOrEmpty(newWppStp.Brand))
                                {
                                    bool brandCheck = IsBrandExist(newWppStp.Brand, out description);
                                    if (brandCheck == false)
                                    {
                                        bool blendCheck = IsBlendExist(newWppStp.Brand, out description);
                                        if (blendCheck == false)
                                        {
                                            bool materialCheck = IsMaterialCodeExist(newWppStp.Brand);
                                            if (materialCheck == false)
                                            {
                                                BrandModel newBrand = new BrandModel();
                                                newBrand.Code = newWppStp.Brand;
                                                newBrand.Description = newWppStp.Description;
                                                newBrand.IsActive = true;
                                                newBrand.LocationID = locationID;
                                                newBrand.ModifiedBy = AccountName;
                                                newBrand.ModifiedDate = DateTime.Now;

                                                string brand = JsonHelper<BrandModel>.Serialize(newBrand);
                                                _brandAppService.Add(brand);
                                            }
                                        }
                                    }

                                    if (string.IsNullOrEmpty(newWppStp.Description))
                                    {
                                        newWppStp.Description = description;
                                    }
                                }

                                newWppStp.Maker = dt_.Rows[index][2].ToString();
                                newWppStp.Packer = dt_.Rows[index][3].ToString();
                                newWppStp.Activity = dt_.Rows[index][4].ToString();
                                newWppStp.PONumber = dt_.Rows[index][5].ToString();
                                newWppStp.OPSNumber = dt_.Rows[index][6].ToString();
                                newWppStp.BatchSAP = dt_.Rows[index][7].ToString();
                                newWppStp.Others = dt_.Rows[index][8].ToString();

                                if (string.IsNullOrEmpty(newWppStp.Maker))
                                {
                                    if (string.IsNullOrEmpty(newWppStp.Packer))
                                    {
                                        if (string.IsNullOrEmpty(newWppStp.Activity))
                                        {
                                            SetFalseTempData(UIResources.FieldEmpty + " Please check on these field.(Maker/Packer/Activity)");
                                            return RedirectToAction("Index");
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(newWppStp.Maker))
                                {
                                    bool checkMaker = IsMachineExist(newWppStp.Maker, machineModelList);
                                    if (checkMaker == false)
                                    {
                                        SetFalseTempData(string.Format(UIResources.MachineCodeNotExist, newWppStp.Maker) + " Please check on cell " + "[C" + (index + 1) + "]");
                                        return RedirectToAction("Index");
                                    }
                                }

                                if (!string.IsNullOrEmpty(newWppStp.Packer))
                                {
                                    bool checkPacker = IsMachineExist(newWppStp.Packer, machineModelList);
                                    if (checkPacker == false)
                                    {
                                        SetFalseTempData(string.Format(UIResources.MachineCodeNotExist, newWppStp.Packer) + " Please check on cell " + "[D" + (index + 1) + "]");
                                        return RedirectToAction("Index");
                                    }
                                }

                                newWppStp.LocationID = locationID;
                                newWppStp.ModifiedBy = AccountName;
                                newWppStp.ModifiedDate = DateTime.Now;

                                for (int shift = 1; shift <= 3; shift++, col++)
                                {
                                    string[] headers = dt.Columns[col].ToString().Split('-');
                                    newWppStp.Date = DateTime.ParseExact(headers[0], "yyyyMMdd", CultureInfo.InvariantCulture);
                                    string amount = dt_.Rows[index][col].ToString();
                                    decimal amountWppStp;

                                    if (decimal.TryParse(amount, out amountWppStp))
                                    {
                                        if (shift == 1)
                                        {
                                            if (amountWppStp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWppStp.Shift1 = amountWppStp;
                                        }
                                        else if (shift == 2)
                                        {
                                            if (amountWppStp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWppStp.Shift2 = amountWppStp;
                                        }
                                        else
                                        {
                                            if (amountWppStp < 0)
                                            {
                                                SetFalseTempData(UIResources.AmountNegative + " Please check on cell " + "[" + index + "]" + "[" + col + "]");
                                                return RedirectToAction("Index");
                                            }
                                            newWppStp.Shift3 = amountWppStp;
                                        }
                                    }
                                }

                                // set the column pointer
                                col--;

                                if (!string.IsNullOrEmpty(newWppStp.Activity) && newWppStp.Shift1 == 0 && newWppStp.Shift2 == 0 && newWppStp.Shift3 == 0)
                                {
                                    continue;
                                }
                                else
                                {
                                    var temp = result.Where(x => x.Brand == newWppStp.Brand && x.Description == newWppStp.Description && x.Date == newWppStp.Date && x.Maker == newWppStp.Maker && x.Packer == newWppStp.Packer).FirstOrDefault();
                                    if (temp == null)
                                    {
                                        result.Add(newWppStp);
                                    }
                                    else
                                    {
                                        temp.Shift1 += newWppStp.Shift1;
                                        temp.Shift2 += newWppStp.Shift2;
                                        temp.Shift3 += newWppStp.Shift3;
                                        if (!string.IsNullOrEmpty(newWppStp.PONumber))
                                        {
                                            temp.PONumber += ", " + newWppStp.PONumber;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    if (model.GoodsType == "Finish")
                    {
                        string sql = "DELETE FROM WPPSTPS WHERE DATE >= '" + model.StartDate.Value.ToString("yyyy-MM-dd") + "' AND DATE <= '" + model.StartDate.Value.AddDays(6).ToString("yyyy-MM-dd") + "' AND BRAND LIKE 'F%' AND LOCATIONID=" + model.ProdCenterID;
                        ExecuteQuery(sql);
                    }
                    else
                    {
                        string sql = "DELETE FROM WPPSTPS WHERE DATE >= '" + model.StartDate.Value.ToString("yyyy-MM-dd") + "' AND DATE <= '" + model.StartDate.Value.AddDays(6).ToString("yyyy-MM-dd") + "' AND BRAND NOT LIKE 'F%' AND LOCATIONID=" + model.ProdCenterID;
                        ExecuteQuery(sql);
                    }

                    foreach (var item in result)
                    {
                        string wppStp = JsonHelper<WppStpModel>.Serialize(item);
                        _wppStpAppService.Add(wppStp);
                    }

                    SetTrueTempData(UIResources.UploadSucceed);
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
                reader = null;
            }

            return RedirectToAction("Index");
        }

        public ActionResult Report()
        {
            return View();
        }

        public ActionResult ExportExcel(string startDate, string endDate, long locID, string locType)
        {
            try
            {
                if (string.IsNullOrEmpty(startDate))
                {
                    SetFalseTempData("start date is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endDate))
                {
                    SetFalseTempData("end date is not selected");
                    return RedirectToAction("Index");
                }


                // Filter Search
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);

                // Getting all data    			
                string wppStpList = _wppStpAppService.GetAll(true);
                List<WppStpModel> wppStps = wppStpList.DeserializeToWppStpList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
                int recordsTotal = wppStps.Count();

                // Filter Search
                DateTime startDateFL = DateTime.Parse(startDate);
                DateTime endDateFL = DateTime.Parse(endDate);

                if ((endDateFL - startDateFL).TotalDays > 31)
                {
                    SetFalseTempData("Maximum range is 31 days");
                    return RedirectToAction("Index");
                }

                wppStps = wppStps.Where(x => x.Date >= startDateFL.Date && x.Date <= endDateFL.Date).ToList();

                byte[] excelData = ExcelGenerator.ExportWppSP(wppStps, startDateFL, endDateFL);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Wpp-SP.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult ExportExcelView(string startDate, string endDate, long locID, string locType)
        {
            try
            {
                if (string.IsNullOrEmpty(startDate))
                {
                    SetFalseTempData("start date is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endDate))
                {
                    SetFalseTempData("end date is not selected");
                    return RedirectToAction("Index");
                }

                // Filter Search
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);

                // Getting all data    			
                string wppStpList = _wppStpAppService.GetAll(true);
                List<WppStpModel> wppStps = wppStpList.DeserializeToWppStpList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
                int recordsTotal = wppStps.Count();

                // Filter Search
                DateTime startDateFL = DateTime.Parse(startDate);
                DateTime endDateFL = DateTime.Parse(endDate);

                if ((endDateFL - startDateFL).TotalDays > 31)
                {
                    SetFalseTempData("Maximum range is 31 days");
                    return RedirectToAction("Index");
                }

                wppStps = wppStps.Where(x => x.Date >= startDateFL.Date && x.Date <= endDateFL.Date).ToList();

                byte[] excelData = ExcelGenerator.ExportWppSPView(wppStps, startDateFL, endDateFL, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Wpp-SP-View.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        private string GetShiftData(WppStpDBModel newData)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", newData.Date.ToString()));
            filters.Add(new QueryFilter("Location", newData.Location));
            filters.Add(new QueryFilter("Packer", newData.Packer));
            filters.Add(new QueryFilter("Maker", newData.Maker));
            filters.Add(new QueryFilter("Brand", newData.Brand));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string exist = _wppStpAppService.Get(filters, true);

            return exist;
        }

        // GET: WppStp/Edit/5
        public ActionResult Edit(int id)
        {
            ViewBag.MachineList = DropDownHelper.BindDropDownMachineMakerCode(_machineAppService);
            ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
            ViewBag.BrandList = DropDownHelper.BindDropDownBrandBlend(_brandAppService, _blendAppService);
            ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

            WppStpModel model = GetWppStp(id);
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

            string brand = _brandAppService.GetBy("Code", model.Brand);
            BrandModel brandModel = brand.DeserializeToBrand();
            model.BrandID = brandModel.ID;

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

        private WppStpModel GetWppStp(int id)
        {
            string wppStp = _wppStpAppService.GetById(id, true);
            WppStpModel wppStpModel = wppStp.DeserializeToWppStp();

            return wppStpModel;
        }

        [HttpPost]
        public ActionResult Edit(WppStpModel model)
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

                WppStpModel oldModel = GetWppStp(Convert.ToInt32(model.ID));

                if (model.LinkUpID > 0)
                {
                    // get Link Up
                    string lu = _referenceAppService.GetDetailById(oldModel.LinkUpID);
                    ReferenceDetailModel luModel = lu.DeserializeToRefDetail();
                    string machineList = _machineAppService.FindBy("LinkUp", luModel.Code, true);
                    List<MachineModel> machineModelList = machineList.DeserializeToMachineList();
                    MachineModel machine = machineModelList.Where(x => x.SubProcess == "Maker").FirstOrDefault();
                    MachineModel machinePacker = machineModelList.Where(x => x.SubProcess == "Packer").FirstOrDefault();

                    model.Maker = machine == null ? string.Empty : machine.Code;
                    model.Packer = machinePacker == null ? string.Empty : machinePacker.Code;
                }
                else
                {
                    model.Maker = oldModel.Maker;
                }

                // get the activity
                string activity = _referenceAppService.GetDetailById(model.ActivityID);
                ReferenceDetailModel activityModel = activity.DeserializeToRefDetail();
                model.Activity = activityModel.Code;

                // get the brand
                string brand = _brandAppService.GetById(model.BrandID);
                BrandModel brandModel = brand.DeserializeToBrand();
                model.Brand = brandModel.Code;
                model.Description = brandModel.Description;

                if (model.SubDepartmentID != 0)
                {
                    model.LocationID = model.SubDepartmentID;
                }
                else if (model.DepartmentID != 0)
                {
                    model.LocationID = model.DepartmentID;
                }
                else
                {
                    model.LocationID = model.ProdCenterID;
                }

                model.Location = _locationAppService.GetLocationFullCode(model.LocationID);

                WppChangesModel changesModel = new WppChangesModel();

                #region Check Changes
                if (model.Shift1 != oldModel.Shift1)
                {
                    changesModel.WPPID = model.ID;
                    changesModel.FieldName = "Shift1";
                    changesModel.OldValue = oldModel.Shift1.ToString();
                    changesModel.NewValue = model.Shift1.ToString();
                    changesModel.DataType = "Numeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;
                    _wppChangesAppService.Add(JsonHelper<WppChangesModel>.Serialize(changesModel));
                }

                if (model.Shift2 != oldModel.Shift2)
                {
                    changesModel.WPPID = model.ID;
                    changesModel.FieldName = "Shift2";
                    changesModel.OldValue = oldModel.Shift2.ToString();
                    changesModel.NewValue = model.Shift2.ToString();
                    changesModel.DataType = "Numeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;
                    _wppChangesAppService.Add(JsonHelper<WppChangesModel>.Serialize(changesModel));
                }

                if (model.Shift3 != oldModel.Shift3)
                {
                    changesModel.WPPID = model.ID;
                    changesModel.FieldName = "Shift3";
                    changesModel.OldValue = oldModel.Shift3.ToString();
                    changesModel.NewValue = model.Shift3.ToString();
                    changesModel.DataType = "Numeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;
                    _wppChangesAppService.Add(JsonHelper<WppChangesModel>.Serialize(changesModel));
                }

                if (model.Maker != oldModel.Maker)
                {
                    changesModel.WPPID = model.ID;
                    changesModel.FieldName = "Maker";
                    changesModel.OldValue = oldModel.Maker;
                    changesModel.NewValue = model.Maker;
                    changesModel.DataType = "NonNumeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;
                    _wppChangesAppService.Add(JsonHelper<WppChangesModel>.Serialize(changesModel));
                }

                if (model.LocationID != oldModel.LocationID)
                {
                    changesModel.WPPID = model.ID;
                    changesModel.FieldName = "LocationID";
                    changesModel.OldValue = oldModel.LocationID.ToString();
                    changesModel.NewValue = model.LocationID.ToString();
                    changesModel.DataType = "Numeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;
                    _wppChangesAppService.Add(JsonHelper<WppChangesModel>.Serialize(changesModel));
                }
                #endregion

                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                string data = JsonHelper<WppStpModel>.Serialize(model);
                _wppStpAppService.Update(data);

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
                _wppStpAppService.Remove(id);

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
                string wppStpList = _wppStpAppService.GetAll(true);
                List<WppStpModel> wppStps = wppStpList.DeserializeToWppStpList().OrderBy(x => x.Date).ToList();
                int recordsTotal = wppStps.Count();

                // Search    - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    wppStps = wppStps.Where(m => (m.Brand != null ? m.Brand.ToLower().Contains(searchValue.ToLower()) : false) ||
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
                                wppStps = wppStps.OrderBy(x => x.Brand).ToList();
                                break;
                            case "location":
                                wppStps = wppStps.OrderBy(x => x.Location).ToList();
                                break;
                            case "description":
                                wppStps = wppStps.OrderBy(x => x.Description).ToList();
                                break;
                            case "packer":
                                wppStps = wppStps.OrderBy(x => x.Packer).ToList();
                                break;
                            case "maker":
                                wppStps = wppStps.OrderBy(x => x.Maker).ToList();
                                break;
                            case "date":
                                wppStps = wppStps.OrderBy(x => x.Date).ToList();
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
                                wppStps = wppStps.OrderByDescending(x => x.Brand).ToList();
                                break;
                            case "description":
                                wppStps = wppStps.OrderByDescending(x => x.Description).ToList();
                                break;
                            case "location":
                                wppStps = wppStps.OrderByDescending(x => x.Location).ToList();
                                break;
                            case "packer":
                                wppStps = wppStps.OrderByDescending(x => x.Packer).ToList();
                                break;
                            case "maker":
                                wppStps = wppStps.OrderByDescending(x => x.Maker).ToList();
                                break;
                            case "date":
                                wppStps = wppStps.OrderByDescending(x => x.Date).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = wppStps.Count();

                // Paging     
                var data = wppStps.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<WppStpModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllWPPWithParam(string startDateFilter, string endDateFilter, long locID, string locType)
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
            string wppStpList = _wppStpAppService.GetAll(true);
            List<WppStpModel> wppStps = wppStpList.DeserializeToWppStpList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
            int recordsTotal = wppStps.Count();

            // Filter Search
            if (startDateFilter != "" && endDateFilter != "")
            {
                DateTime startDateFL = DateTime.Parse(startDateFilter);
                DateTime endDateFL = DateTime.Parse(endDateFilter);
                wppStps = wppStps.Where(x => x.Date >= startDateFL.Date && x.Date <= endDateFL.Date).ToList();
            }
            else if (startDateFilter != "")
            {
                DateTime dateFL = DateTime.Parse(startDateFilter);
                wppStps = wppStps.Where(x => x.Date >= dateFL.Date).ToList();
            }
            else if (endDateFilter != "")
            {
                DateTime endDateFl = DateTime.Parse(endDateFilter);
                wppStps = wppStps.Where(x => x.Date <= endDateFl.Date).ToList();
            }

            // Search    - Correction 231019
            if (!string.IsNullOrEmpty(searchValue))
            {
                wppStps = wppStps.Where(m => (m.Brand != null ? m.Brand.ToLower().Contains(searchValue.ToLower()) : false) ||
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
                            wppStps = wppStps.OrderBy(x => x.Brand).ToList();
                            break;
                        case "location":
                            wppStps = wppStps.OrderBy(x => x.Location).ToList();
                            break;
                        case "description":
                            wppStps = wppStps.OrderBy(x => x.Description).ToList();
                            break;
                        case "packer":
                            wppStps = wppStps.OrderBy(x => x.Packer).ToList();
                            break;
                        case "maker":
                            wppStps = wppStps.OrderBy(x => x.Maker).ToList();
                            break;
                        case "date":
                            wppStps = wppStps.OrderBy(x => x.Date).ToList();
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
                            wppStps = wppStps.OrderByDescending(x => x.Brand).ToList();
                            break;
                        case "description":
                            wppStps = wppStps.OrderByDescending(x => x.Description).ToList();
                            break;
                        case "location":
                            wppStps = wppStps.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "packer":
                            wppStps = wppStps.OrderByDescending(x => x.Packer).ToList();
                            break;
                        case "maker":
                            wppStps = wppStps.OrderByDescending(x => x.Maker).ToList();
                            break;
                        case "date":
                            wppStps = wppStps.OrderByDescending(x => x.Date).ToList();
                            break;
                        default:
                            break;
                    }
                }
            }

            // total number of rows count     
            int recordsFiltered = wppStps.Count();

            // Paging     
            var data = wppStps.Skip(skip).Take(pageSize).ToList();

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

        public bool IsBrandExist(string brand, out string description)
        {
            description = "";
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
                BrandModel brandModel = brands.DeserializeToBrand();
                description = brandModel.Description;
            }
            return isExist;
        }

        public bool IsBlendExist(string blend, out string description)
        {
            bool isExist = true;
            description = "";
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
                BlendModel blendModel = blends.DeserializeToBlend();
                description = blendModel.Description;
            }

            return isExist;
        }

        public bool IsMaterialCodeExist(string material)
        {
            string blends = _materialAppService.GetBy("Code", material);
            return !string.IsNullOrEmpty(blends);
        }

        public bool IsMachineExist(string machineCode, List<MachineModel> machineModelList)
        {
            bool isExist = true;
            long locationId = AccountLocationID;

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
