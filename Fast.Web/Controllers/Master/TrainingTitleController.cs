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
    [CustomAuthorize("training")]
    public class TrainingTitleController : BaseController<TrainingTitleModel>
    {
        #region ::Init::
        private readonly ITrainingTitleAppService _trainingTitleAppService;
        private readonly IMenuAppService _menuService;
        private readonly ITrainingAppService _trainingAppService;
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ITrainingTitleMachineTypeAppService _trainingTitleMachineTypeAppService;
        #endregion

        #region ::Constructor::
        public TrainingTitleController(
            ILoggerAppService logger,
            IReferenceAppService referenceAppService,
            ITrainingTitleMachineTypeAppService trainingTitleMachineTypeAppService,
            ITrainingTitleAppService trainingTitleAppService,
            ITrainingAppService trainingAppService,
            IMenuAppService menuService)
        {
            _referenceAppService = referenceAppService;
            _trainingTitleMachineTypeAppService = trainingTitleMachineTypeAppService;
            _trainingAppService = trainingAppService;
            _trainingTitleAppService = trainingTitleAppService;
            _menuService = menuService;
            _logger = logger;
        }
        #endregion

        #region ::Public Methods::		
        public ActionResult Index()
        {
            GetTempData();

            TrainingTitleModel model = new TrainingTitleModel();
            model.Access = GetAccess(WebConstants.MenuSlug.TRAINING, _menuService);

            return View(model);
        }

        [HttpPost]
        public JsonResult AutoComplete(string prefix)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Title", prefix, Operator.Contains));

            string trainingTitles1 = _trainingTitleAppService.Find(filters);
            List<TrainingTitleModel> trainingTitleModelList = trainingTitles1.DeserializeToTrainingTitleList();

            string trainingTitles = _trainingTitleAppService.GetAll();
            List<TrainingTitleModel> trainingTitleAllModelList = trainingTitles.DeserializeToTrainingTitleList();

            string trainingMTs = _trainingTitleMachineTypeAppService.GetAll();
            List<TrainingTitleMachineTypeModel> trainingMTList = trainingMTs.DeserializeToTrainingTitleMachineTypeList();

            foreach (var item in trainingMTList)
            {
                var temp = trainingTitleAllModelList.Where(x => x.ID == item.TrainingTitleID).FirstOrDefault();
                if (temp != null)
                    item.TrainingTitle = temp.Title;
            }

            List<string> trainingListed = trainingMTList.Select(x => x.TrainingTitle).ToList();

            trainingTitleModelList = trainingTitleModelList.OrderBy(x => x.Title).ToList();

            trainingTitleModelList = trainingTitleModelList.Where(x => !trainingListed.Any(y => y == x.Title)).ToList();

            return Json(trainingTitleModelList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetTrainingMembers(string title)
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

            List<TrainingModel> data = new List<TrainingModel>();
            if (string.IsNullOrEmpty(title))
                return Json(new { data, recordsFiltered = 0, recordsTotal = 0, draw = "1" }, JsonRequestBehavior.AllowGet);

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("TrainingTitle", title, Operator.Contains));

            string trainings = _trainingAppService.Find(filters);
            data = trainings.DeserializeToTrainingList();

            // total number of rows count     
            int recordsFiltered = data.Count();

            int recordsTotal = data.Count();

            // Paging     
            data = data.Skip(skip).Take(pageSize).ToList();

            return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Create()
        {
            ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService);
            ViewBag.TrainingTitleList = DropDownHelper.BindDropDownTrainingTitle(_trainingTitleAppService);

            TrainingTitleMachineTypeModel model = new TrainingTitleMachineTypeModel();
            model.Access = GetAccess(WebConstants.MenuSlug.TRAINING, _menuService);

            return PartialView(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(TrainingTitleMachineTypeModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                if (model.TrainingTitleID == 0)
                {
                    SetFalseTempData("No training title selected");
                    return RedirectToAction("Index");
                }

                if (model.MachineTypeIDs == null || model.MachineTypeIDs.Count() == 0)
                {
                    SetFalseTempData("Please select machine type first");
                    return RedirectToAction("Index");
                }

                foreach (var item in model.MachineTypeIDs)
                {
                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("TrainingTitleID", model.TrainingTitleID.ToString()));
                    filters.Add(new QueryFilter("MachineTypeID", item.ToString()));
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    string exist = _trainingTitleMachineTypeAppService.Get(filters, true);
                    if (!string.IsNullOrEmpty(exist))
                    {
                        SetFalseTempData(string.Format(UIResources.DataExist, model.TrainingTitleID, item));
                        return RedirectToAction("Index");
                    }
                }

                foreach (var machineId in model.MachineTypeIDs)
                {
                    TrainingTitleMachineTypeModel newEntity = new TrainingTitleMachineTypeModel();
                    newEntity.TrainingTitleID = model.TrainingTitleID;
                    newEntity.MachineTypeID = machineId;
                    newEntity.ModifiedBy = AccountName;
                    newEntity.ModifiedDate = DateTime.Now;

                    string data = JsonHelper<TrainingTitleMachineTypeModel>.Serialize(newEntity);

                    _trainingTitleMachineTypeAppService.Add(data);
                }

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.InvalidModelState);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Edit(int id)
        {
            TrainingTitleMachineTypeModel model = GetTrainingTitleMachineType(id);

            string machines = _trainingTitleMachineTypeAppService.FindByNoTracking("TrainingTitleID", model.TrainingTitleID.ToString(), true);
            List<TrainingTitleMachineTypeModel> machineList = machines.DeserializeToTrainingTitleMachineTypeList();
            List<long> machineIDList = machineList.Select(c => c.MachineTypeID).Distinct().ToList();

            ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService, machineIDList);
            ViewBag.TrainingTitleList = DropDownHelper.BindDropDownTrainingTitle(_trainingTitleAppService);

            return PartialView(model);
        }

        [HttpPost]
        public ActionResult Edit(TrainingTitleMachineTypeModel model)
        {
            try
            {
                model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                if (model.TrainingTitleID == 0)
                {
                    SetFalseTempData("No training title selected");
                    return RedirectToAction("Index");
                }

                if (model.MachineTypeIDs == null || model.MachineTypeIDs.Count() == 0)
                {
                    SetFalseTempData("Please select machine type first");
                    return RedirectToAction("Index");
                }

                string trainings = _trainingTitleMachineTypeAppService.FindByNoTracking("TrainingTitleID", model.TrainingTitleID.ToString(), true);
                List<TrainingTitleMachineTypeModel> trainingList = trainings.DeserializeToTrainingTitleMachineTypeList();

                foreach (var item in trainingList)
                {
                    if (!model.MachineTypeIDs.Any(x => x == item.MachineTypeID))
                    {
                        // remove if not selected						
                        _trainingTitleMachineTypeAppService.Remove(item.ID);
                    }
                }

                foreach (var item in model.MachineTypeIDs)
                {
                    if (!trainingList.Any(x => x.MachineTypeID == item))
                    {
                        TrainingTitleMachineTypeModel newEntity = new TrainingTitleMachineTypeModel();
                        newEntity.TrainingTitleID = model.TrainingTitleID;
                        newEntity.MachineTypeID = item;
                        newEntity.ModifiedBy = AccountName;
                        newEntity.ModifiedDate = DateTime.Now;

                        string data = JsonHelper<TrainingTitleMachineTypeModel>.Serialize(newEntity);

                        _trainingTitleMachineTypeAppService.Add(data);
                    }
                }

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
                string tt = _trainingTitleAppService.GetById(id, true);
                TrainingTitleModel ttm = tt.DeserializeToTrainingTitle();
                ttm.IsDeleted = true;

                string data = JsonHelper<TrainingTitleModel>.Serialize(ttm);
                _trainingTitleAppService.Update(data);

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
                string trainingTitles = _trainingTitleAppService.GetAll();
                List<TrainingTitleModel> trainingTitleList = trainingTitles.DeserializeToTrainingTitleList();

                string trainings = _trainingAppService.GetAll();
                List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();

                string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
                List<ReferenceDetailModel> machineTypeReferenceList = mts.DeserializeToRefDetailList();

                foreach (var training in trainingTitleList)
                {
                    var trainingTemp = trainingList.Where(x => x.TrainingTitle == training.Title && x.MachineTypeID != null).FirstOrDefault();
                    if (trainingTemp != null)
                    {
                        if (trainingTemp.MachineTypeID.HasValue)
                        {
                            var mt = machineTypeReferenceList.Where(x => x.ID == trainingTemp.MachineTypeID.Value).FirstOrDefault();
                            if (mt != null)
                            {
                                training.Competency = mt.Code;
                            }
                        }
                    }

                    var trainingTempList = trainingList.Where(x => x.TrainingTitle == training.Title).ToList();

                    training.Trainees = string.Join(",", trainingTempList.Select(x => x.FullName));
                }

                byte[] excelData = ExcelGenerator.ExportTrainingTitle(trainingTitleList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-TrainingTitle.xlsx");
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

                string refMachineTypeList = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
                List<ReferenceDetailModel> refMachineTypeModelList = refMachineTypeList.DeserializeToRefDetailList();

                string trainingTitleList = _trainingTitleAppService.GetAll(true);
                List<TrainingTitleModel> trainingTitleModelList = trainingTitleList.DeserializeToTrainingTitleList();

                // Getting all data    			
                string trainingTitleMTList = _trainingTitleMachineTypeAppService.GetAll(true);
                List<TrainingTitleMachineTypeModel> trainingTitleMTModelList = trainingTitleMTList.DeserializeToTrainingTitleMachineTypeList();


                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                List<TrainingTitleMachineTypeModel> result = new List<TrainingTitleMachineTypeModel>();

                foreach (var item in trainingTitleMTModelList)
                {
                    var exist = result.Where(x => x.TrainingTitleID == item.TrainingTitleID).FirstOrDefault();
                    if (exist == null)
                    {
                        var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
                        item.MachineTypeList = mt == null ? string.Empty : mt.Code;

                        var tt = trainingTitleModelList.Where(x => x.ID == item.TrainingTitleID).FirstOrDefault();
                        item.TrainingTitle = tt == null ? string.Empty : tt.Title;

                        result.Add(item);
                    }
                    else
                    {
                        var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
                        exist.MachineTypeList = exist.MachineTypeList + ", " + mt.Code;
                    }
                }

                int recordsTotal = result.Count();

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => m.TrainingTitle.ToLower().Contains(searchValue.ToLower()) ||
                                               m.MachineTypeList.ToLower().Contains(searchValue.ToLower())).ToList();

                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "trainingtitle":
                                result = result.OrderBy(x => x.TrainingTitle).ToList();
                                break;
                            case "machinetypelist":
                                result = result.OrderBy(x => x.MachineType).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "trainingtitle":
                                result = result.OrderByDescending(x => x.TrainingTitle).ToList();
                                break;
                            case "machinetypelist":
                                result = result.OrderByDescending(x => x.MachineType).ToList();
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

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<TrainingTitleMachineTypeModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllOld()
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
                string trainings = _trainingTitleAppService.GetAll(true);
                List<TrainingTitleModel> trainingList = trainings.DeserializeToTrainingTitleList();

                int recordsTotal = trainingList.Count();

                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    trainingList = trainingList.Where(m => m.Title != null && m.Title.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "title":
                                trainingList = trainingList.OrderBy(x => x.Title).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "title":
                                trainingList = trainingList.OrderByDescending(x => x.Title).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = trainingList.Count();

                // Paging     
                var data = trainingList.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID);

                return Json(new { data = new List<TrainingTitleModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        private TrainingTitleMachineTypeModel GetTrainingTitleMachineType(long trainingTitleMachineTypeID)
        {
            string trainingTitleMachineType = _trainingTitleMachineTypeAppService.GetById(trainingTitleMachineTypeID, true);
            TrainingTitleMachineTypeModel model = trainingTitleMachineType.DeserializeToTrainingTitleMachineType();

            return model;
        }
    }
}
