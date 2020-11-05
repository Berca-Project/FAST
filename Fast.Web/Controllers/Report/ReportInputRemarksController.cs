using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportInputRemarksController : BaseController<ReportRemarksModel>
    {
        private readonly IReportRemarksAppService _reportRemarksAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;

        public ReportInputRemarksController(
        IReportRemarksAppService reportRemarksService,
        ILocationAppService locationAppService,
        IReferenceAppService referenceAppService,        
        ILoggerAppService logger)
        {
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _reportRemarksAppService = reportRemarksService;            
            _logger = logger;
        }
        public ActionResult Index()
        {
            GetTempData();
            return View();
        }
		 public ActionResult Create()
        {
            GetTempData();
            ViewBag.RSAList = BindDropDownRSA();
            ViewBag.ShiftList = BindDropDownShift();
            return View();
        }

        [HttpPost]
        public ActionResult Create(ReportRemarksModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Index");
                }
                model.UserID = AccountID;
                model.LocationID = AccountLocationID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                string data = JsonHelper<ReportRemarksModel>.Serialize(model);
                _reportRemarksAppService.Add(data);

                SetTrueTempData(UIResources.DataHasBeenSaved);
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
            GetTempData();
            ViewBag.RSAList = BindDropDownRSA();
            ViewBag.ShiftList = BindDropDownShift();

            ReportRemarksModel model = GetReportRemarks(id);

            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(ReportRemarksModel model)
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

                string data = JsonHelper<ReportRemarksModel>.Serialize(model);

                _reportRemarksAppService.Update(data);

                SetTrueTempData(UIResources.DataHasBeenSaved);

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
                ReportRemarksModel reportRemarks = GetReportRemarks(id);
                reportRemarks.IsDeleted = true;

                string reportRemData = JsonHelper<ReportRemarksModel>.Serialize(reportRemarks);
                _reportRemarksAppService.Update(reportRemData);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        private ReportRemarksModel GetReportRemarks(long reportRemarksID)
        {
            string reportRemarks = _reportRemarksAppService.GetById(reportRemarksID, true);

            return reportRemarks.DeserializeToReportRemarks();
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

                // Getting all data report remark   			
                string remarksList = _reportRemarksAppService.GetAll(true);
                List<ReportRemarksModel> remarks = remarksList.DeserializeToReportRemarksList().OrderByDescending(x => x.ModifiedDate).ToList();
                          
                int recordsTotal = remarks.Count();

                // Search    - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    remarks = remarks.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.RSA != null ? m.RSA.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Focus != null ? m.Focus.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.ActionPlan != null ? m.ActionPlan.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Date != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false)).ToList();

                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "id":
                                remarks = remarks.OrderBy(x => x.ID).ToList();
                                break;
                            case "shift":
                                remarks = remarks.OrderBy(x => x.Shift).ToList();
                                break;
                            case "rsa":
                                remarks = remarks.OrderBy(x => x.RSA).ToList();
                                break;
                            case "date":
                                remarks = remarks.OrderBy(x => x.Date).ToList();
                                break;
                            case "focus":
                                remarks = remarks.OrderBy(x => x.Focus).ToList();
                                break;
                            case "actionplan":
                                remarks = remarks.OrderBy(x => x.ActionPlan).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "id":
                                remarks = remarks.OrderBy(x => x.ID).ToList();
                                break;
                            case "shift":
                                remarks = remarks.OrderBy(x => x.Shift).ToList();
                                break;
                            case "rsa":
                                remarks = remarks.OrderBy(x => x.RSA).ToList();
                                break;
                            case "date":
                                remarks = remarks.OrderBy(x => x.Date).ToList();
                                break;
                            case "focus":
                                remarks = remarks.OrderBy(x => x.Focus).ToList();
                                break;
                            case "actionplan":
                                remarks = remarks.OrderBy(x => x.ActionPlan).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = remarks.Count();

                // Paging     
                var data = remarks.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
          
                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<ReportRemarksModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
      
        private List<SelectListItem> BindDropDownRSA()
        {          
            List<SelectListItem> _menuList = new List<SelectListItem>();
                      
            _menuList.Add(new SelectListItem
            {
                Text = "MTBF",
                Value = "MTBF"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "VQI",
                Value = "VQI"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "CPQI",
                Value = "CPQI"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Working Time",
                Value = "Working Time"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "STRS",
                Value = "STRS"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Uptime",
                Value = "Uptime"
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
    }
}
