using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("reportrsd")]
    public class ReportInputTargetController : BaseController<InputTargetModel>
    {
        private readonly ILoggerAppService _logger;
        private readonly IInputTargetAppService _inputTargetAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;

        public ReportInputTargetController(
               ILoggerAppService logger,
               IInputTargetAppService inputTargetAppService,
               ILocationAppService locationAppService,
               IReferenceAppService referenceAppService,
               IReferenceDetailAppService referenceDetailAppService
           )
        {
            _logger = logger;
            _inputTargetAppService = inputTargetAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _referenceDetailAppService = referenceDetailAppService;
        }
        public ActionResult Index()
        {
            GetTempData();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.KPIList = BindDropDownKPI(true);
            ViewBag.MonthList = BindDropDownMonth();
            ViewBag.VersionList = BindDropDownVersion(true);
            ViewBag.ProductionCenterList = GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, true);
            return View();
        }
        public ActionResult Create()
        {
            ViewBag.KPIList = BindDropDownKPI(true);
            ViewBag.MonthList = BindDropDownMonth();
            ViewBag.VersionList = BindDropDownVersion(true);
            ViewBag.ProductionCenterList = GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, true);

            return View();
        }
        [HttpPost]
        public ActionResult Create(InputTargetModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Index");
                }

                model.LocationID = AccountLocationID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                string data = JsonHelper<InputTargetModel>.Serialize(model);
                _inputTargetAppService.Add(data);

                ViewBag.Result = true;
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }
        public ActionResult Edit(long id)
        {
            ViewBag.KPIList = BindDropDownKPI(true);
            ViewBag.MonthList = BindDropDownMonth();
            ViewBag.ProductionCenterList = GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, true);
            ViewBag.VersionList = BindDropDownVersion(true);
            InputTargetModel model = GetInputTarget(id);

            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(InputTargetModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    ViewBag.Result = false;
                    ViewBag.ErrorMessage = UIResources.InvalidData;

                    return View("Index");
                }
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                string data = JsonHelper<InputTargetModel>.Serialize(model);
                _inputTargetAppService.Update(data);
                ViewBag.Result = true;

            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                InputTargetModel inputTarget = GetInputTarget(id);
                inputTarget.IsDeleted = true;

                string inputTargetData = JsonHelper<InputTargetModel>.Serialize(inputTarget);
                _inputTargetAppService.Update(inputTargetData);

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
                string inputTargetList = _inputTargetAppService.GetAll(true);
                List<InputTargetModel> result = inputTargetList.DeserializeToInputTargetList().OrderBy(x => x.ID).ToList();
                foreach (var item in result)
                {
                    ReferenceDetailModel refDetail = GetRefDetail(item.ProdCenterID);
                    string pc = refDetail.Code + "-" + refDetail.Description;
                    item.ProductionCenter = pc;
                }
                int recordsTotal = result.Count();

                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => (m.Version != null ? m.Version.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.KPI != null ? m.KPI.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.ProductionCenter != null ? m.ProductionCenter.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.Month != null ? m.Month.ToLower().Contains(searchValue.ToLower()) : false)
                                                ).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "version":
                                result = result.OrderBy(x => x.Version).ToList();
                                break;
                            case "kpi":
                                result = result.OrderBy(x => x.KPI).ToList();
                                break;
                            case "month":
                                result = result.OrderBy(x => x.Month).ToList();
                                break;                            
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "version":
                                result = result.OrderByDescending(x => x.Version).ToList();
                                break;
                            case "kpi":
                                result = result.OrderByDescending(x => x.KPI).ToList();
                                break;
                            case "month":
                                result = result.OrderByDescending(x => x.Month).ToList();
                                break;                            
                            default:
                                break;
                        }
                    }
                }
                foreach (var item in result)
                {
                    ReferenceDetailModel refDetail = GetRefDetail(item.ProdCenterID);
                    string pc = refDetail.Code + "-" + refDetail.Description;
                    item.ProductionCenter = pc;
                }
                result = result.OrderByDescending(x => x.ID).ToList();
                // total number of rows count     
                int recordsFiltered = result.Count();

                // Paging     
                List<InputTargetModel> data = result.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<InputDailyModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult Upload(InputTargetModel model)
        {
            IExcelDataReader reader = null;
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
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

                    int countRows = reader.RowCount;
                    DataTable dt = new DataTable();
                    DataTable dt_ = reader.AsDataSet().Tables[0];
                    List<string> listVersion = new List<string>()
                    {"OB",
                    "Internal Target",
                    "RF02",
                    "RF06",
                    "RF09",
                    "RF11"
                    };
                    List<string> listMonth = new List<string>();

                    for (int i = 1; i <= 12; i++)
                    {
                        listMonth.Add(getMonthByNumber(i));
                    }
                  
                    string refDetail = _referenceDetailAppService.GetAll(true);
                    List<ReferenceDetailModel> refModelList = refDetail.DeserializeToRefDetailList();

                    List<InputTargetModel> listTarget = new List<InputTargetModel>();
                    for (int i = 7; i < countRows; i++)
                    {

                        InputTargetModel newTarget = new InputTargetModel();
                        newTarget.ProdCenterID = model.ProdCenterID;
                        newTarget.KPI = dt_.Rows[i][3].ToString();

                        if (listVersion.Contains(dt_.Rows[i][1].ToString()))
                        {
                            newTarget.Version = dt_.Rows[i][1].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        if (listMonth.Contains(dt_.Rows[i][2].ToString()))
                        {
                            newTarget.Month = dt_.Rows[i][2].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        if (decimal.TryParse(dt_.Rows[i][4].ToString(), out decimal value))
                        {
                            newTarget.Value = value;
                            newTarget.ModifiedDate = DateTime.Now;
                            newTarget.ModifiedBy = AccountName;
                            listTarget.Add(newTarget);
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                    }
                    foreach (InputTargetModel itm in listTarget)
                    {

                        //Check target already exist
                        List<QueryFilter> targetFilter = new List<QueryFilter>();
                        targetFilter.Add(new QueryFilter("ProdCenterID", itm.ProdCenterID));
                        targetFilter.Add(new QueryFilter("KPI", itm.KPI));
                        targetFilter.Add(new QueryFilter("Version", itm.Version));
                        targetFilter.Add(new QueryFilter("Month", itm.Month));
                        targetFilter.Add(new QueryFilter("IsDeleted", "0"));

                        InputTargetModel dataTarget = _inputTargetAppService.Get(targetFilter,true).DeserializeToInputTarget();

                        if (dataTarget.ID != 0)
                        {
                            dataTarget.Value = itm.Value;
                            dataTarget.ModifiedDate = DateTime.Now;
                            dataTarget.ModifiedBy = AccountName;
                            string input = JsonHelper<InputTargetModel>.Serialize(dataTarget);
                            _inputTargetAppService.Update(input);

                        }
                        else
                        {
                            string input = JsonHelper<InputTargetModel>.Serialize(itm);
                            _inputTargetAppService.Add(input);
                        }
                    }

                    reader.Close();
                    reader.Dispose();
                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
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
        public static string getMonthByNumber(int i)
        {
            string fullMonthName = new DateTime(2020, i, 1).ToString("MMM", CultureInfo.InvariantCulture);
            return fullMonthName.ToUpper();
        }
        [HttpPost]
        public ActionResult GenerateExcel(InputTargetModel model)
        {
            try
            {
                UserModel user = (UserModel)Session["UserLogon"];
                //if (!user.LocationID.HasValue)
                //{
                //    SetFalseTempData("Location for the logged user is invalid");
                //    return RedirectToAction("Index");
                //}
                long locationID = user.LocationID.Value;
                string version = model.Version;
                string kpi = model.KPI;
                string month = model.Month;

                List<SelectListItem> kpiList = BindDropDownKPI(false);

                ReferenceDetailModel refDetail = GetRefDetail(model.ProdCenterID);
                byte[] excelData = ExcelGenerator.ExportInputTarget(AccountName, version, kpi, month, kpiList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=InputTarget.xlsx");
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

        private List<SelectListItem> BindDropDownVersion(bool isWithSelect)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            if (isWithSelect == true)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Please Select -",
                    Value = ""
                });
            }
            _menuList.Add(new SelectListItem
            {
                Text = "OB",
                Value = "OB"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Internal Target",
                Value = "Internal Target"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF02",
                Value = "RF02"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF06",
                Value = "RF06"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF09",
                Value = "RF09"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF11",
                Value = "RF11"
            });
          
            return _menuList;

        }
        private List<SelectListItem> BindDropDownKPI(bool isWithSelect)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            if (isWithSelect == true)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "All",
                    Value = "All"
                });
            }
            
            _menuList.Add(new SelectListItem
            {
                Text = "MTBF",
                Value = "MTBF"
            });           
            _menuList.Add(new SelectListItem
            {
                Text = "CPQI",
                Value = "CPQI"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "VQI",
                Value = "VQI"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Working Time",
                Value = "Working Time"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Uptime",
                Value = "Uptime"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "STRS",
                Value = "STRS"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Production Volume",
                Value = "Production Volume"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "CRR",
                Value = "CRR"
            });

            return _menuList;

        }
        private List<SelectListItem> BindDropDownMonth()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "All",
                Value = "All"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "JANUARI",
                Value = getMonthByNumber(1)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "FEBRUARI",
                Value = getMonthByNumber(2)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "MARET",
                Value = getMonthByNumber(3)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "APRIL",
                Value = getMonthByNumber(4)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "MEI",
                Value = getMonthByNumber(5)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "JUNI",
                Value = getMonthByNumber(6)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "JULI",
                Value = getMonthByNumber(7)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "AGUSTUS",
                Value = getMonthByNumber(8)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "SEPTEMBER",
                Value = getMonthByNumber(9)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "OKTOBER",
                Value = getMonthByNumber(10)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "NOVEMBER",
                Value = getMonthByNumber(11)
            });
            _menuList.Add(new SelectListItem
            {
                Text = "DESEMBER",
                Value = getMonthByNumber(12)
            });
            //_menuList.Add(new SelectListItem
            //{
            //    Text = "Year",
            //    Value = "Year"
            //});

            return _menuList;

        }
        private InputTargetModel GetInputTarget(long inputTargetId)
        {
            string inputTarget = _inputTargetAppService.GetById(inputTargetId, true);

            return inputTarget.DeserializeToInputTarget();
        }

      
        public static List<SelectListItem> GetProductionCenterInIndonesia(ILocationAppService _locationAppService, IReferenceAppService _referenceAppService, bool isIncludePleaseSelect)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            if (isIncludePleaseSelect)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Please Select -",
                    Value = ""
                });
            }

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = text.Code + " - " + text.Description,
                        Value = item.ID.ToString()
                    });
                }
            }

            return _menuList;
        }
        private ReferenceDetailModel GetRefDetail(long idRef)
        {

            string pcs = _locationAppService.GetById(idRef, true);
            LocationModel locModel = pcs.DeserializeToLocation();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            ReferenceDetailModel refLocModelList = refPC.DeserializeToRefDetailList().Where(x => x.Code == locModel.Code).FirstOrDefault();

            return refLocModelList;
        }
        private long GetRefID(string value, List<ReferenceDetailModel> refModelList)
        {
            var result = refModelList.Where(x => x.Code == value).FirstOrDefault();
            return result.ID;
        }
        private List<SelectListItem> BindDropDownVersionForUpload()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
          
            _menuList.Add(new SelectListItem
            {
                Text = "OB",
                Value = "OB"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Internal Target",
                Value = "Internal Target"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF02",
                Value = "RF02"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF06",
                Value = "RF06"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF09",
                Value = "RF09"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "RF11",
                Value = "RF11"
            });

            return _menuList;

        }
    }
}
