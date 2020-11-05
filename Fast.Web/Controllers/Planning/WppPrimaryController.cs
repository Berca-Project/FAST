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
    [CustomAuthorize("wppprimary")]
    public class WppPrimaryController : BaseController<WppPrimModel>
    {
        private readonly IWppPrimAppService _wppPrimAppService;
        private readonly IWppPrimaryAppService _wppPrimaryAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBlendAppService _blendAppService;
        private readonly IBrandAppService _brandAppService;
        private readonly IMaterialCodeAppService _materialAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IMenuAppService _menuAppService;
        private readonly IWeeksAppService _weeksAppService;
        private readonly ICalendarHolidayAppService _calendarHolidayAppService;

        public WppPrimaryController(
            IBlendAppService blendAppService,
            IBrandAppService brandAppService,
            IMaterialCodeAppService materialAppService,
            IWppPrimAppService wppPrimAppService,
            IReferenceAppService referenceAppService,
            ILocationAppService locationAppService,
            IMachineAppService machineAppService,
            IMenuAppService menuAppService,
            ILoggerAppService logger,
            IWeeksAppService weeksAppService,
            IWppPrimaryAppService wppPrimaryAppService,
            ICalendarHolidayAppService calendarHolidayAppService)
        {
            _wppPrimaryAppService = wppPrimaryAppService;
            _referenceAppService = referenceAppService;
            _wppPrimAppService = wppPrimAppService;
            _logger = logger;
            _menuAppService = menuAppService;
            _locationAppService = locationAppService;
            _machineAppService = machineAppService;
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
            ViewBag.BlendList = DropDownHelper.BindDropDownBrand(_referenceAppService);
            ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.YearList = DropDownHelper.BindDropDownYearWpp();
            ViewBag.WeekList = DropDownHelper.BuildEmptyList();

            WppPrimModel model = new WppPrimModel();
            model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

            return View(model);
        }

        public ActionResult Detail()
        {
            GetTempData();

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.YearList = DropDownHelper.BindDropDownYearWpp();
            ViewBag.WeekList = DropDownHelper.BuildEmptyList();

            WppPrimModel model = new WppPrimModel();
            model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(WppPrimModel model)
        {

            IExcelDataReader reader = null;

            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);
                model.Date = DateTime.Now;

                ViewBag.MachineList = DropDownHelper.BindDropDownMachineCode(_machineAppService);
                ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
                ViewBag.BlendList = DropDownHelper.BindDropDownBrandBlend(_brandAppService, _blendAppService);
                ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);

                ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
                ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.YearList = DropDownHelper.BindDropDownYearWpp();
                ViewBag.WeekList = DropDownHelper.BuildEmptyList();

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(ModelState.GetModelStateErrors());
                    return RedirectToAction("Detail");
                }

                if (model.ProdCenterID == 0)
                {
                    SetFalseTempData(UIResources.ProdCenterIsMissing);
                    return RedirectToAction("Detail");
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
                        return RedirectToAction("Detail");
                    }

                    string blends = _blendAppService.GetAll();
                    List<BlendModel> blendList = blends.DeserializeToBlendList().Where(x => x.LocationID == model.ProdCenterID).ToList();

                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;
                    DataTable dt = new DataTable();
                    DataTable dt_ = reader.AsDataSet().Tables[0];

                    List<WppPrimModel> result = new List<WppPrimModel>();
                    List<string> header = new List<string>();

                    string username = AccountName;
                    string blendLocation = _locationAppService.GetLocationFullCode(model.ProdCenterID);
                    long blendLocationID = model.ProdCenterID;

                    for (int i = 0; i < dt_.Columns.Count; i++)
                    {
                        int dateIndex = i == 0 ? 3 : (3 + (i * 5));
                        if (dateIndex < dt_.Columns.Count)
                        {
                            string temp = dt_.Rows[0][dateIndex].ToString();
                            if (!string.IsNullOrEmpty(temp))
                            {
                                DateTime wppDate;
                                if (DateTime.TryParseExact(temp, "yyyyMMdd", null, DateTimeStyles.None, out wppDate))
                                {
                                    if (wppDate.Date <= DateTime.Now.AddDays(-7).Date)
                                    {
                                        SetFalseTempData(string.Format(UIResources.DateWppLessThanEqualToday, wppDate.ToString("yyyyMMdd"), DateTime.Now.ToString("yyyyMMdd")));
                                        return RedirectToAction("Detail");
                                    }
                                }
                                else
                                {
                                    SetFalseTempData(string.Format(UIResources.DateWppInvalid, temp));
                                    return RedirectToAction("Detail");
                                }

                                for (int index = 2; index < dt_.Rows.Count; index++)
                                {
                                    string blend = dt_.Rows[index][dateIndex - 2].ToString();
                                    string batchLama = dt_.Rows[index][dateIndex - 1].ToString();
                                    string po = dt_.Rows[index][dateIndex].ToString();
                                    string batchsap = dt_.Rows[index][dateIndex + 1].ToString();

                                    if (string.IsNullOrEmpty(blend))
                                        continue;

                                    if (!blendList.Any(x => x.Code == blend.Trim()))
                                    {
                                        BlendModel newBlend = new BlendModel();
                                        newBlend.Code = blend;
                                        newBlend.Description = blend;
                                        newBlend.IsActive = true;
                                        newBlend.LocationID = blendLocationID;
                                        newBlend.ModifiedBy = AccountName;
                                        newBlend.ModifiedDate = DateTime.Now;

                                        string blendStr = JsonHelper<BlendModel>.Serialize(newBlend);
                                        _blendAppService.Add(blendStr);
                                    }

                                    WppPrimModel wppPrim = new WppPrimModel();
                                    wppPrim.Date = wppDate;
                                    wppPrim.Blend = blend;
                                    wppPrim.PONumber = po;
                                    wppPrim.BatchLama = batchLama;
                                    wppPrim.BatchSAP = batchsap;
                                    wppPrim.Location = blendLocation;
                                    wppPrim.LocationID = blendLocationID;
                                    wppPrim.ModifiedBy = username;
                                    wppPrim.ModifiedDate = DateTime.Now;

                                    result.Add(wppPrim);
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    foreach (var item in result)
                    {
                        // get the existing record
                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("Location", item.Location));
                        filters.Add(new QueryFilter("Blend", item.Blend));
                        filters.Add(new QueryFilter("PONumber", item.PONumber));
                        filters.Add(new QueryFilter("BatchLama", item.BatchLama));
                        filters.Add(new QueryFilter("BatchSAP", item.BatchSAP));
                        filters.Add(new QueryFilter("Date", item.Date.ToString()));

                        string oldData = _wppPrimAppService.Get(filters, true);
                        if (string.IsNullOrEmpty(oldData))
                        {
                            string wpp = JsonHelper<WppPrimModel>.Serialize(item);
                            _wppPrimAppService.Add(wpp);
                        }
                    }

                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    if (model.PostedFilename == null)
                        SetFalseTempData(UIResources.NoSelectedFile);
                    else
                        SetFalseTempData(UIResources.FileCorrupted);
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.CreateFailed);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                reader = null;
            }

            return RedirectToAction("Detail");
        }

        [HttpPost]
        public ActionResult Upload(WppPrimModel model)
        {

            IExcelDataReader reader = null;

            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.WPP, _menuAppService);

                ViewBag.MachineList = DropDownHelper.BindDropDownMachineCode(_machineAppService);
                ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
                ViewBag.BlendList = DropDownHelper.BindDropDownBrandBlend(_brandAppService, _blendAppService);
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

                if (model.Year == 0)
                {
                    SetFalseTempData("No selected year");
                    return RedirectToAction("Index");
                }

                if (model.Week == 0)
                {
                    SetFalseTempData("no selected week");
                    return RedirectToAction("Index");
                }

                model.Date = FirstDateOfWeek(model.Year, model.Week);

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

                    List<WppPrimaryModel> result = new List<WppPrimaryModel>();
                    List<string> header = new List<string>();

                    string username = AccountName;
                    Dictionary<string, BlendModel> blendCollection = new Dictionary<string, BlendModel>();
                    string blendLocation = "";
                    long blendLocationID = 0;

                    string blendHeader = dt_.Rows[4][12].ToString();
                    if (blendHeader != "BLEND")
                    {
                        SetFalseTempData("No BLEND header detected. Please make sure it is the right file");
                        return RedirectToAction("Index");
                    }

                    for (int index = 5; index < dt_.Rows.Count; index++)
                    {
                        string blend = dt_.Rows[index][12].ToString();

                        if (!string.IsNullOrEmpty(blend) && blend != "Total" && blend != "0")
                        {
                            #region ::Blend::
                            if (!blendCollection.ContainsKey(blend))
                            {
                                BlendModel blendModel = IsBlendExist(blend);
                                if (blendModel == null)
                                {
                                    BlendModel newBlend = new BlendModel();
                                    newBlend.Code = blend;
                                    newBlend.Description = blend;
                                    newBlend.IsActive = true;
                                    newBlend.LocationID = blendLocationID == 0 ? AccountLocationID : blendLocationID;
                                    newBlend.ModifiedBy = AccountName;
                                    newBlend.ModifiedDate = model.Date;

                                    string blendStr = JsonHelper<BlendModel>.Serialize(newBlend);
                                    _blendAppService.Add(blendStr);

                                    newBlend.Location = _locationAppService.GetLocationFullCode(newBlend.LocationID.Value);
                                    blendLocationID = newBlend.LocationID.Value;
                                    blendLocation = newBlend.Location;
                                }
                                else
                                {
                                    if (blendModel.LocationID.HasValue)
                                    {
                                        blendModel.Location = _locationAppService.GetLocationFullCode(blendModel.LocationID.Value);
                                        blendLocationID = blendModel.LocationID.Value;
                                        blendLocation = blendModel.Location;
                                    }

                                    blendCollection.Add(blend, blendModel);
                                }
                            }
                            else
                            {
                                BlendModel blendModelData;
                                blendCollection.TryGetValue(blend, out blendModelData);

                                if (blendModelData.LocationID.HasValue)
                                {
                                    blendLocation = blendModelData.Location;
                                    blendLocationID = blendModelData.LocationID.Value;
                                }
                            }
                            #endregion

                            string volPerOps = dt_.Rows[index][13].ToString();
                            string monday = dt_.Rows[index][14].ToString();
                            string tuesday = dt_.Rows[index][15].ToString();
                            string wednesday = dt_.Rows[index][16].ToString();
                            string thursday = dt_.Rows[index][17].ToString();
                            string friday = dt_.Rows[index][18].ToString();
                            string saturday = dt_.Rows[index][19].ToString();
                            string sunday = dt_.Rows[index][20].ToString();

                            WppPrimaryModel newWpp = new WppPrimaryModel();
                            newWpp.Blend = blend;
                            //newWpp.LocationID = AccountLocationID;
                            //newWpp.Location = _locationAppService.GetLocationFullCode(AccountLocationID);
                            long locationID = model.ProdCenterID == 0 ? AccountLocationID : model.ProdCenterID;

                            newWpp.LocationID = locationID;
                            newWpp.Location = _locationAppService.GetLocationFullCode(locationID);

                            newWpp.VolPerOps = ConvertToInt(volPerOps);
                            newWpp.Monday = ConvertToInt(monday);
                            newWpp.Tuesday = ConvertToInt(tuesday);
                            newWpp.Wednesday = ConvertToInt(wednesday);
                            newWpp.Thursday = ConvertToInt(thursday);
                            newWpp.Friday = ConvertToInt(friday);
                            newWpp.Saturday = ConvertToInt(saturday);
                            newWpp.Sunday = ConvertToInt(sunday);
                            newWpp.StartDate = model.Date;
                            newWpp.Week = model.Week;

                            result.Add(newWpp);
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("StartDate", model.Date.ToString()));
                    filters.Add(new QueryFilter("LocationID", model.ProdCenterID.ToString()));
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    string wppprims = _wppPrimaryAppService.FindNoTracking(filters);
                    List<WppPrimaryModel> wppList = wppprims.DeserializeToWppPrimaryList();

                    foreach (var item in result)
                    {
                        var temp = wppList.Where(x => x.Blend == item.Blend).FirstOrDefault();
                        if (temp != null)
                        {
                            item.ID = temp.ID;
                            string wpp = JsonHelper<WppPrimaryModel>.Serialize(item);
                            _wppPrimaryAppService.Update(wpp);
                        }
                        else
                        {
                            string wpp = JsonHelper<WppPrimaryModel>.Serialize(item);
                            _wppPrimaryAppService.Add(wpp);
                        }
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

        private int ConvertToInt(string volPerOps)
        {
            int result = 0;
            if (int.TryParse(volPerOps, out result))
            {
                return result;
            }

            return 0;
        }

        public ActionResult Report()
        {
            return View();
        }

        public ActionResult Sim()
        {
            GetTempData();

            WppPrimaryModel model = new WppPrimaryModel();

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            ViewBag.YearList = DropDownHelper.BindDropDownYearWpp();
            ViewBag.WeekList = DropDownHelper.BuildEmptyList();

            return View(model);
        }

        public List<SelectListItem> BindDropDownWeeks(string year)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (string.IsNullOrEmpty(year))
            {
                return _menuList;
            }

            int yearVal = Convert.ToInt32(year);

            //var startMonth = yearVal == DateTime.Now.Year ? new DateTime(yearVal, DateTime.Now.Month, 1) : new DateTime(yearVal, 1, 1);
            var startMonth = new DateTime(yearVal, 1, 1);
            var endMonth = new DateTime(yearVal, 12, 31);
            var currentCulture = CultureInfo.CurrentCulture;
            var weeks = new List<int>();

            for (var dt = startMonth; dt < endMonth; dt = dt.AddDays(1))
            {
                var weekNo = currentCulture.Calendar.GetWeekOfYear(dt, currentCulture.DateTimeFormat.CalendarWeekRule, currentCulture.DateTimeFormat.FirstDayOfWeek);
                if (!weeks.Contains(weekNo))
                    weeks.Add(weekNo);
            }

            foreach (var item in weeks)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.ToString(),
                    Value = item.ToString()
                });
            }

            return _menuList;
        }

        [HttpPost]
        public ActionResult GetWeekByYear(string year)
        {
            List<SelectListItem> _menuList = BindDropDownWeeks(year);
            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        private string GetShiftData(WppPrimDBModel newData)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", newData.Date.ToString()));
            filters.Add(new QueryFilter("Location", newData.Location));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string exist = _wppPrimAppService.Get(filters, true);

            return exist;
        }

        public ActionResult Edit(int id)
        {
            ViewBag.MachineList = DropDownHelper.BindDropDownMachineMakerCode(_machineAppService);
            ViewBag.LinkUpList = DropDownHelper.BindDropDownLinkUp(_referenceAppService, _machineAppService);
            ViewBag.BlendList = DropDownHelper.BindDropDownBrandBlend(_brandAppService, _blendAppService);
            ViewBag.WppTypeList = DropDownHelper.BindDropDownWppType(_referenceAppService);
            ViewBag.YearList = DropDownHelper.BindDropDownYearWpp();
            ViewBag.WeekList = DropDownHelper.BuildEmptyList();

            WppPrimModel model = GetWppPrim(id);
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

        private WppPrimModel GetWppPrim(int id)
        {
            string wpp = _wppPrimAppService.GetById(id, true);
            WppPrimModel wppModel = wpp.DeserializeToWppPrim();

            return wppModel;
        }

        [HttpPost]
        public ActionResult Edit(WppPrimModel model)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BuildEmptyList();
                ViewBag.LinkUpList = DropDownHelper.BuildEmptyList();
                ViewBag.BlendList = DropDownHelper.BuildEmptyList();
                ViewBag.WppPrimTypeList = DropDownHelper.BuildEmptyList();

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

                WppPrimModel oldModel = GetWppPrim(Convert.ToInt32(model.ID));

                if (model.LinkUpID > 0)
                {
                    // get Link Up
                    string lu = _referenceAppService.GetDetailById(oldModel.LinkUpID);
                    ReferenceDetailModel luModel = lu.DeserializeToRefDetail();
                    string machineList = _machineAppService.FindBy("LinkUp", luModel.Code, true);
                    List<MachineModel> machineModelList = machineList.DeserializeToMachineList();
                    MachineModel machine = machineModelList.Where(x => x.SubProcess == "Maker").FirstOrDefault();
                    MachineModel machinePacker = machineModelList.Where(x => x.SubProcess == "Packer").FirstOrDefault();

                }

                // get the activity
                string activity = _referenceAppService.GetDetailById(model.ActivityID);
                ReferenceDetailModel activityModel = activity.DeserializeToRefDetail();
                model.Activity = activityModel.Code;

                // get the blend
                string blend = _referenceAppService.GetDetailById(model.BlendID);
                ReferenceDetailModel blendModel = blend.DeserializeToRefDetail();
                model.Blend = blendModel.Code;

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

                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                string data = JsonHelper<WppPrimModel>.Serialize(model);
                _wppPrimAppService.Update(data);

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
                _wppPrimAppService.Remove(id);

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
                string wppList = _wppPrimAppService.GetAll(true);
                List<WppPrimModel> wpps = wppList.DeserializeToWppPrimList().OrderBy(x => x.Date).ToList();
                int recordsTotal = wpps.Count();

                // Search    - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    wpps = wpps.Where(m => (m.Blend != null ? m.Blend.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.PONumber != null ? m.PONumber.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BatchSAP != null ? m.BatchSAP.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Others != null ? m.Others.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BatchLama != null ? m.BatchLama.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
                }


                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderBy(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderBy(x => x.Location).ToList();
                                break;
                            case "batchlama":
                                wpps = wpps.OrderBy(x => x.BatchLama).ToList();
                                break;
                            case "batchsap":
                                wpps = wpps.OrderBy(x => x.BatchSAP).ToList();
                                break;
                            case "ponumber":
                                wpps = wpps.OrderBy(x => x.PONumber).ToList();
                                break;
                            case "date":
                                wpps = wpps.OrderBy(x => x.Date).ToList();
                                break;
                            case "others":
                                wpps = wpps.OrderBy(x => x.Others).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderByDescending(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderByDescending(x => x.Location).ToList();
                                break;
                            case "batchlama":
                                wpps = wpps.OrderByDescending(x => x.BatchLama).ToList();
                                break;
                            case "batchsap":
                                wpps = wpps.OrderByDescending(x => x.BatchSAP).ToList();
                                break;
                            case "ponumber":
                                wpps = wpps.OrderByDescending(x => x.PONumber).ToList();
                                break;
                            case "date":
                                wpps = wpps.OrderByDescending(x => x.Date).ToList();
                                break;
                            case "others":
                                wpps = wpps.OrderByDescending(x => x.Others).ToList();
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
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<WppPrimModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllWPPWithParam(string year, string week, long locID, string locType)
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
            string wppList = _wppPrimAppService.GetAll(true);
            List<WppPrimModel> wpps = wppList.DeserializeToWppPrimList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
            int recordsTotal = wpps.Count();

            // Filter Search
            if (!string.IsNullOrEmpty(year) && !string.IsNullOrEmpty(week))
            {
                DateTime startDate = FirstDateOfWeek(int.Parse(year), int.Parse(week));
                DateTime endDate = startDate.AddDays(6);
                wpps = wpps.Where(x => x.Date >= startDate && x.Date <= endDate).ToList();
            }

            // Search    - Correction 231019
            if (!string.IsNullOrEmpty(searchValue))
            {
                wpps = wpps.Where(m => (m.Blend != null ? m.Blend.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.Activity != null ? m.Activity.ToLower().Contains(searchValue.ToLower()) : false) ||
                                        (m.BatchLama != null ? m.BatchLama.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "blend":
                            wpps = wpps.OrderBy(x => x.Blend).ToList();
                            break;
                        case "location":
                            wpps = wpps.OrderBy(x => x.Location).ToList();
                            break;
                        case "batchlama":
                            wpps = wpps.OrderBy(x => x.BatchLama).ToList();
                            break;
                        case "batchsap":
                            wpps = wpps.OrderBy(x => x.BatchSAP).ToList();
                            break;
                        case "ponumber":
                            wpps = wpps.OrderBy(x => x.PONumber).ToList();
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
                        case "blend":
                            wpps = wpps.OrderByDescending(x => x.Blend).ToList();
                            break;
                        case "location":
                            wpps = wpps.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "batchlama":
                            wpps = wpps.OrderByDescending(x => x.BatchLama).ToList();
                            break;
                        case "batchsap":
                            wpps = wpps.OrderByDescending(x => x.BatchSAP).ToList();
                            break;
                        case "ponumber":
                            wpps = wpps.OrderByDescending(x => x.PONumber).ToList();
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

        [HttpPost]
        public ActionResult GetAllSimulation()
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
                string wppList = _wppPrimaryAppService.GetAll(true);
                List<WppPrimaryModel> wpps = wppList.DeserializeToWppPrimaryList().OrderBy(x => x.StartDate).ToList();

                foreach (var item in wpps)
                {
                    item.Total = item.Monday + item.Tuesday + item.Wednesday + item.Thursday + item.Friday + item.Saturday + item.Sunday;
                }

                int recordsTotal = wpps.Count();

                if (!string.IsNullOrEmpty(searchValue))
                {
                    wpps = wpps.Where(m => (m.Blend != null ? m.Blend.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderBy(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderBy(x => x.Location).ToList();
                                break;
                            case "volperops":
                                wpps = wpps.OrderBy(x => x.VolPerOps).ToList();
                                break;
                            case "monday":
                                wpps = wpps.OrderBy(x => x.Monday).ToList();
                                break;
                            case "tuesday":
                                wpps = wpps.OrderBy(x => x.Tuesday).ToList();
                                break;
                            case "wednesday":
                                wpps = wpps.OrderBy(x => x.Wednesday).ToList();
                                break;
                            case "thursday":
                                wpps = wpps.OrderBy(x => x.Thursday).ToList();
                                break;
                            case "friday":
                                wpps = wpps.OrderBy(x => x.Friday).ToList();
                                break;
                            case "saturday":
                                wpps = wpps.OrderBy(x => x.Saturday).ToList();
                                break;
                            case "sunday":
                                wpps = wpps.OrderBy(x => x.Sunday).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderByDescending(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderByDescending(x => x.Location).ToList();
                                break;
                            case "volperops":
                                wpps = wpps.OrderByDescending(x => x.VolPerOps).ToList();
                                break;
                            case "monday":
                                wpps = wpps.OrderByDescending(x => x.Monday).ToList();
                                break;
                            case "tuesday":
                                wpps = wpps.OrderByDescending(x => x.Tuesday).ToList();
                                break;
                            case "wednesday":
                                wpps = wpps.OrderByDescending(x => x.Wednesday).ToList();
                                break;
                            case "thursday":
                                wpps = wpps.OrderByDescending(x => x.Thursday).ToList();
                                break;
                            case "friday":
                                wpps = wpps.OrderByDescending(x => x.Friday).ToList();
                                break;
                            case "saturday":
                                wpps = wpps.OrderByDescending(x => x.Saturday).ToList();
                                break;
                            case "sunday":
                                wpps = wpps.OrderByDescending(x => x.Sunday).ToList();
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

                if (data.Count > 0)
                {
                    // add total 
                    WppPrimaryModel total = new WppPrimaryModel();
                    total.Blend = "Total";
                    total.VolPerOps = data.Sum(x => x.VolPerOps);
                    total.Monday = data.Sum(x => x.Monday);
                    total.Tuesday = data.Sum(x => x.Tuesday);
                    total.Wednesday = data.Sum(x => x.Wednesday);
                    total.Thursday = data.Sum(x => x.Thursday);
                    total.Friday = data.Sum(x => x.Friday);
                    total.Saturday = data.Sum(x => x.Saturday);
                    total.Sunday = data.Sum(x => x.Sunday);
                    total.Total = data.Sum(x => x.Total);

                    data.Add(total);
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<WppPrimModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllSimulationWithParam(string year, string week, string endYear, string endWeek, long locID, string locType)
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

                List<long> locIDList = _locationAppService.GetLocIDListByLocType(locID, locType);

                // Getting all data    			
                string wppList = _wppPrimaryAppService.GetAll(true);
                List<WppPrimaryModel> wpps = wppList.DeserializeToWppPrimaryList().Where(x => locIDList.Any(y => y == x.LocationID)).OrderBy(x => x.StartDate).ToList();

                DateTime startDate = DateTime.MinValue;
                DateTime endDate = DateTime.MaxValue;

                if (!string.IsNullOrEmpty(year) && !string.IsNullOrEmpty(week) && !string.IsNullOrEmpty(endYear) && !string.IsNullOrEmpty(endWeek))
                {
                    startDate = FirstDateOfWeek(int.Parse(year), int.Parse(week));
                    endDate = FirstDateOfWeek(int.Parse(endYear), int.Parse(endWeek));

                    // Filter Search
                    wpps = wpps.Where(x => x.StartDate >= startDate && x.StartDate <= endDate).ToList();

                }

                /*
				else if (!string.IsNullOrEmpty(year) && !string.IsNullOrEmpty(week))
				{
					startDate = FirstDateOfWeek(int.Parse(year), int.Parse(week));

					// Filter Search
					wpps = wpps.Where(x => x.StartDate == startDate).ToList();
				}
				*/

                foreach (var item in wpps)
                {
                    item.Total = item.Monday + item.Tuesday + item.Wednesday + item.Thursday + item.Friday + item.Saturday + item.Sunday;
                }

                int recordsTotal = wpps.Count();

                if (!string.IsNullOrEmpty(searchValue))
                {
                    wpps = wpps.Where(m => (m.Blend != null ? m.Blend.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderBy(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderBy(x => x.Location).ToList();
                                break;
                            case "volperops":
                                wpps = wpps.OrderBy(x => x.VolPerOps).ToList();
                                break;
                            case "monday":
                                wpps = wpps.OrderBy(x => x.Monday).ToList();
                                break;
                            case "tuesday":
                                wpps = wpps.OrderBy(x => x.Tuesday).ToList();
                                break;
                            case "wednesday":
                                wpps = wpps.OrderBy(x => x.Wednesday).ToList();
                                break;
                            case "thursday":
                                wpps = wpps.OrderBy(x => x.Thursday).ToList();
                                break;
                            case "friday":
                                wpps = wpps.OrderBy(x => x.Friday).ToList();
                                break;
                            case "saturday":
                                wpps = wpps.OrderBy(x => x.Saturday).ToList();
                                break;
                            case "sunday":
                                wpps = wpps.OrderBy(x => x.Sunday).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "blend":
                                wpps = wpps.OrderByDescending(x => x.Blend).ToList();
                                break;
                            case "location":
                                wpps = wpps.OrderByDescending(x => x.Location).ToList();
                                break;
                            case "volperops":
                                wpps = wpps.OrderByDescending(x => x.VolPerOps).ToList();
                                break;
                            case "monday":
                                wpps = wpps.OrderByDescending(x => x.Monday).ToList();
                                break;
                            case "tuesday":
                                wpps = wpps.OrderByDescending(x => x.Tuesday).ToList();
                                break;
                            case "wednesday":
                                wpps = wpps.OrderByDescending(x => x.Wednesday).ToList();
                                break;
                            case "thursday":
                                wpps = wpps.OrderByDescending(x => x.Thursday).ToList();
                                break;
                            case "friday":
                                wpps = wpps.OrderByDescending(x => x.Friday).ToList();
                                break;
                            case "saturday":
                                wpps = wpps.OrderByDescending(x => x.Saturday).ToList();
                                break;
                            case "sunday":
                                wpps = wpps.OrderByDescending(x => x.Sunday).ToList();
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

                if (data.Count > 0)
                {
                    // add total 
                    WppPrimaryModel total = new WppPrimaryModel();
                    total.Blend = "Total";
                    total.VolPerOps = data.Sum(x => x.VolPerOps);
                    total.Monday = data.Sum(x => x.Monday);
                    total.Tuesday = data.Sum(x => x.Tuesday);
                    total.Wednesday = data.Sum(x => x.Wednesday);
                    total.Thursday = data.Sum(x => x.Thursday);
                    total.Friday = data.Sum(x => x.Friday);
                    total.Saturday = data.Sum(x => x.Saturday);
                    total.Sunday = data.Sum(x => x.Sunday);
                    total.Total = data.Sum(x => x.Total);

                    data.Add(total);
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<WppPrimModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GenerateSimulationExcel(string year, string week, string endYear, string endWeek, long locID, string locType)
        {
            try
            {
                List<long> locIDList = _locationAppService.GetLocIDListByLocType(locID, locType);

                // Getting all data    			
                string wppList = _wppPrimaryAppService.GetAll(true);
                List<WppPrimaryModel> wpps = wppList.DeserializeToWppPrimaryList().Where(x => locIDList.Any(y => y == x.LocationID)).OrderBy(x => x.StartDate).ToList();

                if (string.IsNullOrEmpty(year))
                {
                    SetFalseTempData("No selected start year");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(week))
                {
                    SetFalseTempData("no selected start week");
                    return RedirectToAction("Index");
                }
                if (string.IsNullOrEmpty(endYear))
                {
                    SetFalseTempData("No selected end year");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endWeek))
                {
                    SetFalseTempData("no selected end week");
                    return RedirectToAction("Index");
                }

                // Filter Search
                DateTime startDate = FirstDateOfWeek(int.Parse(year), int.Parse(week));
                DateTime endDate = FirstDateOfWeek(int.Parse(endYear), int.Parse(endWeek));

                wpps = wpps.Where(x => x.StartDate >= startDate && x.StartDate <= endDate).ToList();

                foreach (var item in wpps)
                {
                    item.Total = item.Sunday + item.Tuesday + item.Wednesday + item.Thursday + item.Friday + item.Saturday + item.Sunday;
                }

                if (wpps.Count > 0)
                {
                    // add total 
                    WppPrimaryModel total = new WppPrimaryModel();
                    total.Blend = "Total";
                    total.VolPerOps = wpps.Sum(x => x.VolPerOps);
                    total.Monday = wpps.Sum(x => x.Monday);
                    total.Tuesday = wpps.Sum(x => x.Tuesday);
                    total.Wednesday = wpps.Sum(x => x.Wednesday);
                    total.Thursday = wpps.Sum(x => x.Thursday);
                    total.Friday = wpps.Sum(x => x.Friday);
                    total.Saturday = wpps.Sum(x => x.Saturday);
                    total.Sunday = wpps.Sum(x => x.Sunday);
                    total.Total = wpps.Sum(x => x.Total);

                    wpps.Add(total);
                }

                byte[] excelData = ExcelGenerator.ExportWppSimulation(wpps, AccountName, startDate, endDate);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Wpp-Primary.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult GenerateOPSExcel(string year, string week, long locID, string locType)
        {
            try
            {
                if (string.IsNullOrEmpty(year))
                {
                    SetFalseTempData("No selected year");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(week))
                {
                    SetFalseTempData("no selected week");
                    return RedirectToAction("Index");
                }

                // Getting all data    			
                string wppList = _wppPrimAppService.GetAll(true);
                List<WppPrimModel> wpps = wppList.DeserializeToWppPrimList().Where(x => x.LocationID == locID).OrderBy(x => x.Date).ToList();

                int recordsTotal = wpps.Count();

                // Filter Search
                DateTime startDate = DateTime.MinValue;
                if (!string.IsNullOrEmpty(year) && !string.IsNullOrEmpty(week))
                {
                    startDate = FirstDateOfWeek(int.Parse(year), int.Parse(week));
                    DateTime endDate = startDate.AddDays(7);
                    wpps = wpps.Where(x => x.Date >= startDate && x.Date <= endDate).ToList();
                }

                byte[] excelData = ExcelGenerator.ExportWppOps(wpps, startDate);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Wpp-OPS.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult DownloadTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateWpp_Primary.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateWpp_Primary.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult DownloadTemplatePrimary()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateWpp_Primary_Sim.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateWpp_Primary_Sim.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        #region Helper
        public static DateTime FirstDateOfWeek(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);

            int daysOffset = (int)DayOfWeek.Monday - (int)jan1.DayOfWeek;

            DateTime firstMonday = jan1.AddDays(daysOffset);

            int firstWeek = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(jan1, CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, DayOfWeek.Monday);

            if (firstWeek <= 1)
            {
                weekOfYear -= 1;
            }

            return firstMonday.AddDays(weekOfYear * 7);
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

        public BlendModel IsBlendExist(string blend)
        {
            string blends = _blendAppService.GetBy("Code", blend);
            BlendModel blendModel = blends.DeserializeToBlend();
            return blendModel.ID == 0 ? null : blendModel;
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
