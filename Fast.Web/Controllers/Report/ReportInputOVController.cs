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
    public class ReportInputOVController : BaseController<InputOVModel>
    {
        private readonly IInputOVAppService _inputOVAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        public ReportInputOVController(
           IInputOVAppService inputOVService,
           ILocationAppService locationAppService,
           IReferenceAppService referenceAppService,
           ILoggerAppService logger)
        {
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _inputOVAppService = inputOVService;
            _logger = logger;
        }
        public ActionResult Index()
        {          
            return View();
        }
        public ActionResult Create()
        {
            ViewBag.WasteCategoryList = BindDropDownWasteCategory();
            return View();
        }

        [HttpPost]
        public ActionResult Create(InputOVModel model)
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

                string data = JsonHelper<InputOVModel>.Serialize(model);
                _inputOVAppService.Add(data);

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
            ViewBag.WasteCategoryList = BindDropDownWasteCategory();
            InputOVModel model = GetInputOV(id);

            return View(model);
        }
        [HttpPost]
        public ActionResult Edit(InputOVModel model)
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

                string data = JsonHelper<InputOVModel>.Serialize(model);

                _inputOVAppService.Update(data);

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
                InputOVModel inputOV = GetInputOV(id);
                inputOV.IsDeleted = true;

                string inputOVData = JsonHelper<InputOVModel>.Serialize(inputOV);
                _inputOVAppService.Update(inputOVData);

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

                // Getting all data report remark   			
                string ovList = _inputOVAppService.GetAll(true);
                List<InputOVModel> ov = ovList.DeserializeToInputOVList().OrderByDescending(x => x.ID).ToList();

                int recordsTotal = ov.Count();

                // Search    - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    ov = ov.Where(m => (m.WasteCategory != null ? m.WasteCategory.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.Week != null ? m.Week.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.OV != null ? m.OV.ToLower().Contains(searchValue.ToLower()) : false)).ToList();                    
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "id":
                                ov = ov.OrderBy(x => x.ID).ToList();
                                break;
                            case "wastecategory":
                                ov = ov.OrderBy(x => x.WasteCategory).ToList();
                                break;
                            case "week":
                                ov = ov.OrderBy(x => x.Week).ToList();
                                break;
                            case "ov":
                                ov = ov.OrderBy(x => x.OV).ToList();
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
                                ov = ov.OrderBy(x => x.ID).ToList();
                                break;
                            case "wastecategory":
                                ov = ov.OrderBy(x => x.WasteCategory).ToList();
                                break;
                            case "week":
                                ov = ov.OrderBy(x => x.Week).ToList();
                                break;
                            case "ov":
                                ov = ov.OrderBy(x => x.OV).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = ov.Count();

                // Paging     
                var data = ov.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<InputOVModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        private InputOVModel GetInputOV(long inputOVID)
        {
            string inputOV = _inputOVAppService.GetById(inputOVID, true);
            return inputOV.DeserializeToInputOV();
        }

        private List<SelectListItem> BindDropDownWasteCategory()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Wet waste JCC",
                Value = "Wet waste JCC"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Wet waste Conveyor/Separator",
                Value = "Wet waste Conveyor/Separator"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Heavies after thresser",
                Value = "Heavies after thresser"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Debu cleaning Rajangan line",
                Value = "Debu cleaning Rajangan line"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Dust Filter Krs-Rjg",
                Value = "Dust Filter Krs-Rjg"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "Vibro Sieving Rajangan",
                Value = "Vibro Sieving Rajangan"
            });

            return _menuList;
        }
    }
}
