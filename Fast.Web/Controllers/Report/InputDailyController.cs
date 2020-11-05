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
using System.Threading;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("reportrsd")]
    public class InputDailyController : BaseController<InputDailyModel>
    {
        private readonly ILoggerAppService _logger;
        private readonly IInputDailyAppService _inputDailyAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly IMachineAppService _machineAppService;
        public InputDailyController(
                ILoggerAppService logger,
                IInputDailyAppService inputDailyAppService,
                ILocationAppService locationAppService,
                IReferenceAppService referenceAppService,
                IReferenceDetailAppService referenceDetailAppService,
                IMachineAppService machineAppService
            )
        {
            _logger = logger;
            _inputDailyAppService = inputDailyAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _machineAppService = machineAppService;
        }
        // GET: InputDaily
        public ActionResult Index()
        {
            GetTempData();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            //ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            ViewBag.ProductionCenterList = GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, true);
            ViewBag.ShiftList = BindDropDownShift();
            return View();
        }
        public ActionResult Create()
        {
            ViewBag.ShiftList = BindDropDownShift();
          
            return View();
        }
        [HttpPost]
        public ActionResult Create(InputDailyModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Index");
                }
                model.ProdCenterID = AccountProdCenterID;
                model.LocationID = AccountLocationID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                model.Week = GetCurrentWeekNumber(model.Date);                

                string data = JsonHelper<InputDailyModel>.Serialize(model);
                _inputDailyAppService.Add(data);

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
            ViewBag.ShiftList = BindDropDownShift();
         
            InputDailyModel model = GetInputDaily(id);

            return View(model);
        }
        [HttpPost]
        public ActionResult Edit(InputDailyModel model)
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

                string data = JsonHelper<InputDailyModel>.Serialize(model);
                _inputDailyAppService.Update(data);

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
                InputDailyModel inputDaily = GetInputDaily(id);
                inputDaily.IsDeleted = true;
                inputDaily.ModifiedBy = AccountName;
                inputDaily.ModifiedDate = DateTime.Now;

                string inputDailyData = JsonHelper<InputDailyModel>.Serialize(inputDaily);
                _inputDailyAppService.Update(inputDailyData);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
    
        [HttpPost]
        public ActionResult Upload(InputDailyModel model)
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

                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;
                    DataTable dt = new DataTable();
                    DataTable dt_ = reader.AsDataSet().Tables[0];

                    int countRows = dt_.Rows.Count;
                    int countColumn = dt_.Columns.Count;

                   
                    string refDetail = _referenceDetailAppService.GetAll(true);
                    List<ReferenceDetailModel> refModelList = refDetail.DeserializeToRefDetailList();

                    string checkDate = dt_.Rows[5][1].ToString();
                    string checkPC = dt_.Rows[6][0].ToString();
                    DateTime dtVal = model.Date;
                    string sub = checkPC.Substring(0, 2);
                    
                    string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
                    List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();
                    long pcID = locModelList.Where(x => x.Code == sub).FirstOrDefault().ID;




                    List<InputDailyModel> listInput = new List<InputDailyModel>();
                    string LinkUp = "";
                    for (int i = 7; i < countRows; i = i + 4)
                    {
                        InputDailyModel newInput = new InputDailyModel();
                        if (dt_.Rows[i][0].ToString() == "Link Up")
                        {
                            LinkUp = dt_.Rows[i++][1].ToString();
                            newInput.LinkUp = LinkUp;
                        }
                        else
                        {
                            newInput.LinkUp = LinkUp;
                        }
                        newInput.ProdCenterID = pcID;
                        newInput.Date = dtVal;
                        newInput.Week = GetCurrentWeekNumber(dtVal);
                        newInput.LocationID = AccountLocationID;
                        newInput.ModifiedBy = AccountName;
                        newInput.ModifiedDate = DateTime.Now;

                        string shiftString = dt_.Rows[i][0].ToString();
                        if (int.TryParse(dt_.Rows[i][0].ToString().Trim().LastOrDefault().ToString(), out int shift))
                        {
                            newInput.Shift = shift.ToString();

                            //Get value
                            int colIndex = 1;
                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal mtbfVal))
                            {
                                newInput.MTBFValue = mtbfVal;
                            }
                            newInput.MTBFFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.MTBFActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal cpqifVal))
                            {
                                newInput.CPQIValue = cpqifVal;
                            }
                            newInput.CPQIFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.CPQIActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal vqifVal))
                            {
                                newInput.VQIValue = vqifVal;
                            }
                            newInput.VQIFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.VQIActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal workingVal))
                            {
                                newInput.WorkingValue = workingVal;
                            }
                            newInput.WorkingFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.WorkingActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal uptimeVal))
                            {
                                newInput.UptimeValue = uptimeVal;
                            }
                            newInput.UptimeFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.UptimeActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal strsVal))
                            {
                                newInput.STRSValue = strsVal;
                            }
                            newInput.STRSFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.STRSActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal prodVolVal))
                            {
                                newInput.ProdVolumeValue = prodVolVal;
                            }
                            newInput.ProdVolumeFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.ProdVolumeActPlan = dt_.Rows[i + 3][colIndex++].ToString();

                            string crrString = dt_.Rows[i + 1][colIndex].ToString();
                            if (decimal.TryParse(dt_.Rows[i + 1][colIndex].ToString(), out decimal crrVal))
                            {
                                newInput.CRRValue = crrVal;
                            }
                            newInput.CRRFocus = dt_.Rows[i + 2][colIndex].ToString();
                            newInput.CRRActPlan = dt_.Rows[i + 3][colIndex++].ToString();
                            listInput.Add(newInput);
                        }
                    }


                    foreach (InputDailyModel idm in listInput)
                    {
                        //Check Input already exist
                        List<QueryFilter> inputFilter = new List<QueryFilter>();
                        inputFilter.Add(new QueryFilter("ProdCenterID", idm.ProdCenterID));
                        inputFilter.Add(new QueryFilter("Shift", idm.Shift));
                        inputFilter.Add(new QueryFilter("LinkUp", idm.LinkUp));
                        inputFilter.Add(new QueryFilter("Date", idm.Date.ToString()));
                        inputFilter.Add(new QueryFilter("IsDeleted", "0"));

                        InputDailyModel dataDaily = _inputDailyAppService.Get(inputFilter, true).DeserializeToInputDaily();

                        if (dataDaily.ID != 0)
                        {
                            dataDaily.MTBFValue = idm.MTBFValue > 0 ? idm.MTBFValue : dataDaily.MTBFValue;
                            dataDaily.MTBFFocus = idm.MTBFFocus != null && idm.MTBFFocus != "" ? idm.MTBFFocus : dataDaily.MTBFFocus;
                            dataDaily.MTBFActPlan = idm.MTBFActPlan != null && idm.MTBFActPlan != "" ? idm.MTBFActPlan : dataDaily.MTBFActPlan;

                            dataDaily.CPQIValue = idm.CPQIValue > 0 ? idm.CPQIValue : dataDaily.CPQIValue;
                            dataDaily.CPQIFocus = idm.CPQIFocus != null && idm.CPQIFocus != "" ? idm.CPQIFocus : dataDaily.CPQIFocus;
                            dataDaily.CPQIActPlan = idm.CPQIActPlan != null && idm.CPQIActPlan != "" ? idm.CPQIActPlan : dataDaily.CPQIActPlan;

                            dataDaily.VQIValue = idm.VQIValue > 0 ? idm.VQIValue : dataDaily.VQIValue;
                            dataDaily.VQIFocus = idm.VQIFocus != null && idm.VQIFocus != "" ? idm.VQIFocus : dataDaily.VQIFocus;
                            dataDaily.VQIActPlan = idm.VQIActPlan != null && idm.VQIActPlan != "" ? idm.VQIActPlan : dataDaily.VQIActPlan;

                            dataDaily.WorkingValue = idm.WorkingValue > 0 ? idm.WorkingValue : dataDaily.WorkingValue;
                            dataDaily.WorkingFocus = idm.WorkingFocus != null && idm.WorkingFocus != "" ? idm.WorkingFocus : dataDaily.WorkingFocus;
                            dataDaily.WorkingActPlan = idm.WorkingActPlan != null && idm.WorkingActPlan != "" ? idm.WorkingActPlan : dataDaily.WorkingActPlan;

                            dataDaily.UptimeValue = idm.UptimeValue > 0 ? idm.UptimeValue : dataDaily.UptimeValue;
                            dataDaily.UptimeFocus = idm.UptimeFocus != null && idm.UptimeFocus != "" ? idm.UptimeFocus : dataDaily.UptimeFocus;
                            dataDaily.UptimeActPlan = idm.UptimeActPlan != null && idm.UptimeActPlan != "" ? idm.UptimeActPlan : dataDaily.UptimeActPlan;

                            dataDaily.STRSValue = idm.STRSValue > 0 ? idm.STRSValue : dataDaily.STRSValue;
                            dataDaily.STRSFocus = idm.STRSFocus != null && idm.STRSFocus != "" ? idm.STRSFocus : dataDaily.STRSFocus;
                            dataDaily.STRSActPlan = idm.STRSActPlan != null && idm.STRSActPlan != "" ? idm.STRSActPlan : dataDaily.STRSActPlan;

                            dataDaily.ProdVolumeValue = idm.ProdVolumeValue > 0 ? idm.ProdVolumeValue : dataDaily.ProdVolumeValue;
                            dataDaily.ProdVolumeFocus = idm.ProdVolumeFocus != null && idm.ProdVolumeFocus != "" ? idm.ProdVolumeFocus : dataDaily.ProdVolumeFocus;
                            dataDaily.ProdVolumeActPlan = idm.ProdVolumeActPlan != null && idm.ProdVolumeActPlan != "" ? idm.ProdVolumeActPlan : dataDaily.CPQIActPlan;

                            dataDaily.CRRValue = idm.CRRValue > 0 ? idm.CRRValue : dataDaily.CRRValue;
                            dataDaily.CRRFocus = idm.CRRFocus != null && idm.CRRFocus != "" ? idm.CRRFocus : dataDaily.CRRFocus;
                            dataDaily.CRRActPlan = idm.CRRActPlan != null && idm.CRRActPlan != "" ? idm.CRRActPlan : dataDaily.CRRActPlan;

                            dataDaily.ModifiedDate = DateTime.Now;
                            dataDaily.ModifiedBy = AccountName;
                            string input = JsonHelper<InputDailyModel>.Serialize(dataDaily);
                            _inputDailyAppService.Update(input);
                        }
                        else
                        {
                            string input = JsonHelper<InputDailyModel>.Serialize(idm);
                            _inputDailyAppService.Add(input);
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
        [HttpPost]
        public ActionResult GenerateExcel(InputDailyModel model)
        {
            try
            {
                UserModel user = (UserModel)Session["UserLogon"];

                long locationID = user.LocationID.Value;

                //string dateVal = model.Date.ToString("dd-MMM-yy");
                ReferenceDetailModel refDetail = GetRefDetail(model.ProdCenterID);
                string pc = refDetail.Code + "-" + refDetail.Description;

                //List<QueryFilter> filters = new List<QueryFilter>();
                //filters.Add(new QueryFilter("LinkUp", null, Operator.NotEqual));
                //filters.Add(new QueryFilter("LinkUp", "", Operator.NotEqual));
                //filters.Add(new QueryFilter("IsDeleted", "0"));

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(model.ProdCenterID, "productioncenter");
                List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList().Where(x => x.LinkUp != null && x.LinkUp != "").ToList();
                List<string> listLinkUp = machineList.Where(x =>
                        locationIdList.Any(y => y == x.LocationID)
                    ).Select(x => x.LinkUp).Distinct().ToList();

                byte[] excelData = ExcelGenerator.ExportInputDaily(AccountName, pc,listLinkUp);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=InputDaily.xlsx");
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

        [HttpPost]
        public ActionResult GetAll(DateTime dtFilter, string prodCenter, string machine)
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
                filters.Add(new QueryFilter("Date", dtFilter.ToString()));
                filters.Add(new QueryFilter("ProdCenterID", prodCenter));
                if (!machine.Equals("All"))
                {
                    filters.Add(new QueryFilter("LinkUp", machine));
                }
                filters.Add(new QueryFilter("IsDeleted", "0"));
                string inputDailyList = _inputDailyAppService.Find(filters);
                List<InputDailyModel> result = inputDailyList.DeserializeToInputDailyList().OrderBy(x => x.ID).ToList();

           

                int recordsTotal = result.Count();

                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||                                               
                                               (m.LinkUp != null ? m.LinkUp.ToLower().Contains(searchValue.ToLower()) : false)
                                                ).ToList();
                }

                //if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                //{
                //    if (sortColumnDir == "asc")
                //    {
                //        switch (sortColumn.ToLower())
                //        {
                //            case "shift":
                //                result = result.OrderBy(x => x.Shift).ToList();
                //                break;
                //            case "kpi":
                //                result = result.OrderBy(x => x.KPI).ToList();
                //                break;
                //            case "date":
                //                result = result.OrderBy(x => x.Date).ToList();
                //                break;
                //            case "linkup":
                //                result = result.OrderBy(x => x.LinkUp).ToList();
                //                break;                                                        
                //            default:
                //                break;
                //        }
                //    }
                //    else
                //    {
                //        switch (sortColumn.ToLower())
                //        {
                //            case "shift":
                //                result = result.OrderByDescending(x => x.Shift).ToList();
                //                break;
                //            case "kpi":
                //                result = result.OrderByDescending(x => x.KPI).ToList();
                //                break;
                //            case "date":
                //                result = result.OrderByDescending(x => x.Date).ToList();
                //                break;
                //            case "linkup":
                //                result = result.OrderByDescending(x => x.LinkUp).ToList();
                //                break;                                                        
                //            default:
                //                break;
                //        }
                //    }
                //}
                foreach (var item in result)
                {
                    ReferenceDetailModel refDetail = GetRefDetail(item.ProdCenterID);
                    string pc = refDetail.Code + "-" + refDetail.Description;
                    item.ProductionCenter = pc;
                }
                result = result.OrderByDescending(x => x.Date).ToList();
                // total number of rows count     
                int recordsFiltered = result.Count();

                // Paging     
                List<InputDailyModel> data = result.Skip(skip).Take(pageSize).ToList();
           
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

        #region Helper 
   
        private List<SelectListItem> BindDropDownShift()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "1",
                Value = "1"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "2",
                Value = "2"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "3",
                Value = "3"
            });

            return _menuList;
        }
        public string GetCurrentWeekNumber(DateTime dateTime)
        {
            var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dateTime, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weeknum.ToString();/*dateTime.ToString($"{weeknum}")*/
        }
        private InputDailyModel GetInputDaily(long inputDailyID)
        {
            string inputDaily = _inputDailyAppService.GetById(inputDailyID, true);
            return inputDaily.DeserializeToInputDaily();
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
        #endregion
    }
}