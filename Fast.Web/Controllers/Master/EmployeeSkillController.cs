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
    [CustomAuthorize("employeeskill")]
    public class EmployeeSkillController : BaseController<UserMachineTypeModel>
    {
        #region ::Init::
        private readonly IUserMachineTypeAppService _userMachineTypeAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IUserAppService _userAppService;
        private readonly IEmployeeAppService _empAppService;
        private readonly ITrainingAppService _trainingAppService;
        private readonly ILoggerAppService _logger;
        private readonly IMenuAppService _menuService;
        private readonly ITrainingTitleAppService _trainingTitleAppService;
        private readonly ITrainingTitleMachineTypeAppService _trainingTitleMachineTypeAppService;
        #endregion

        #region ::Constructor::
        public EmployeeSkillController(
            IUserMachineTypeAppService userMachineTypeAppService,
            IUserAppService userAppService,
            ILoggerAppService logger,
            IMenuAppService menuService,
            ITrainingAppService trainingAppService,
            IEmployeeAppService empAppService,
            ITrainingTitleAppService trainingTitleAppService,
            ITrainingTitleMachineTypeAppService trainingTitleMachineTypeAppService,
            IReferenceAppService referenceAppService)
        {
            _trainingAppService = trainingAppService;
            _userMachineTypeAppService = userMachineTypeAppService;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
            _empAppService = empAppService;
            _menuService = menuService;
            _logger = logger;
            _trainingTitleMachineTypeAppService = trainingTitleMachineTypeAppService;
            _trainingTitleAppService = trainingTitleAppService;
        }
        #endregion

        #region ::Public Methods::
        public ActionResult Index()
        {
            GetTempData();

            ViewBag.MachineTypeList = DropDownHelper.BuildMultiEmpty();

            UserMachineTypeModel model = new UserMachineTypeModel();
            model.Access = GetAccess(WebConstants.MenuSlug.EMP_SKILL, _menuService);

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

            string emplist = _empAppService.Find(filters);
            List<EmployeeModel> empModelList = emplist.DeserializeToEmployeeList();

            if (prefix.All(Char.IsDigit))
            {
                empModelList = empModelList.OrderBy(x => x.EmployeeID).ToList();
            }
            else
            {
                empModelList = empModelList.OrderBy(x => x.FullName).ToList();
            }

            string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
            List<ReferenceDetailModel> machineTypeList = mts.DeserializeToRefDetailList();

            string trainings = _trainingAppService.GetAll();
            List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();

            foreach (var item in empModelList)
            {
                var trainingEmpList = trainingList.Where(x => x.EmployeeID == item.EmployeeID).ToList();

                string machineList = string.Empty;
                foreach (var training in trainingEmpList)
                {
                    if (training.MachineTypeID.HasValue)
                    {
                        var mt = machineTypeList.Where(x => x.ID == training.MachineTypeID.Value).FirstOrDefault();
                        if (mt != null && !machineList.Contains(mt.Code))
                        {
                            if (machineList == string.Empty)
                                machineList = mt.Code;
                            else
                                machineList += ", " + mt.Code;
                        }
                    }
                }

                item.MachineReferenceList = machineList;
            }

            return Json(empModelList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetTrainingList(string empID)
        {
            var draw = Request.Form.GetValues("draw").FirstOrDefault();
            var start = Request.Form.GetValues("start").FirstOrDefault();
            var length = Request.Form.GetValues("length").FirstOrDefault();

            // Paging Size (10,20,50,100)    
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            int skip = start != null ? Convert.ToInt32(start) : 0;

            List<TrainingModel> data = new List<TrainingModel>();
            if (string.IsNullOrEmpty(empID))
                return Json(new { data, recordsFiltered = 0, recordsTotal = 0, draw = "1" }, JsonRequestBehavior.AllowGet);

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("EmployeeID", empID, Operator.Contains));

            string trainings = _trainingAppService.Find(filters);
            data = trainings.DeserializeToTrainingList();

            // total number of rows count     
            int recordsFiltered = data.Count();

            int recordsTotal = data.Count();

            // Paging     
            data = data.Skip(skip).Take(pageSize).ToList();

            return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ExportExcel()
        {
            try
            {
                // Getting all data    			
                string users = _userAppService.GetAll(true);
                List<UserModel> userList = users.DeserializeToUserList();
                userList = userList.Where(x => x.IsActive && x.IsFast).ToList();

                string emps = _empAppService.GetAll(true);
                List<EmployeeModel> empList = emps.DeserializeToEmployeeList();

                string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
                List<ReferenceDetailModel> machineTypeList = mts.DeserializeToRefDetailList();

                string trainingTitles = _trainingTitleAppService.GetAll();
                List<TrainingTitleModel> trainingTitleList = trainingTitles.DeserializeToTrainingTitleList();

                string trainingMachineTypes = _trainingTitleMachineTypeAppService.GetAll();
                List<TrainingTitleMachineTypeModel> trainingMachineTypeList = trainingMachineTypes.DeserializeToTrainingTitleMachineTypeList();

                string trainings = _trainingAppService.GetAll();
                List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();
                List<string> userTrainingList = trainingList.Select(c => c.EmployeeID).Distinct().ToList();

                userList = userList.Where(x => userTrainingList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

                List<UserMachineTypeModel> result = new List<UserMachineTypeModel>();

                Dictionary<long, string> trainingSkillMap = new Dictionary<long, string>();
                foreach (var user in userList)
                {
                    EmployeeModel emp = empList.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                    UserMachineTypeModel newEntity = new UserMachineTypeModel();
                    newEntity.UserName = user.UserName;
                    newEntity.EmployeeID = user.EmployeeID;
                    newEntity.FullName = emp == null ? "" : emp.FullName;
                    newEntity.PositionDesc = emp == null ? "" : emp.PositionDesc;

                    List<TrainingModel> trainingTempList = trainingList.Where(x => x.EmployeeID.Trim() == user.EmployeeID).ToList();

                    foreach (var training in trainingTempList)
                    {
                        TrainingTitleModel trainingTitleTemp = trainingTitleList.Where(x => x.Title == training.TrainingTitle).FirstOrDefault();
                        if (trainingTitleTemp != null)
                        {
                            List<TrainingTitleMachineTypeModel> trainingMTList = trainingMachineTypeList.Where(x => x.TrainingTitleID == trainingTitleTemp.ID).ToList();

                            string machineTypeListRef = "";
                            if (!trainingSkillMap.TryGetValue(trainingTitleTemp.ID, out machineTypeListRef))
                            {
                                machineTypeListRef = GetMachineTypeList(machineTypeList, trainingMTList, machineTypeListRef);
                                trainingSkillMap.Add(trainingTitleTemp.ID, machineTypeListRef);
                            }

                            if (string.IsNullOrEmpty(newEntity.Skills))
                                newEntity.Skills = machineTypeListRef;
                            else
                            {
                                if (!string.IsNullOrEmpty(machineTypeListRef) && !newEntity.Skills.Contains(machineTypeListRef))
                                    newEntity.Skills += ", " + machineTypeListRef;
                            }
                        }
                    }

                    result.Add(newEntity);
                }

                byte[] excelData = ExcelGenerator.ExportMasterEmployeeSkill(result, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Employee-Skill.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Create()
        {
            ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService);

            UserMachineTypeModel model = new UserMachineTypeModel();
            model.Access = GetAccess(WebConstants.MenuSlug.EMP_SKILL, _menuService);

            return PartialView(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(UserMachineTypeModel userMachineTypeModel)
        {
            try
            {
                ViewBag.MachineTypeList = DropDownHelper.BuildEmptyList();

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(userMachineTypeModel.EmployeeID))
                {
                    SetFalseTempData("Please select user first");
                    return RedirectToAction("Index");
                }

                if (userMachineTypeModel.MachineTypeIDs == null || userMachineTypeModel.MachineTypeIDs.Count() == 0)
                {
                    SetFalseTempData("Please select machine type first");
                    return RedirectToAction("Index");
                }

                string user = _userAppService.GetBy("EmployeeID", userMachineTypeModel.EmployeeID, true);
                UserModel userModel = user.DeserializeToUser();

                userMachineTypeModel.UserID = userModel.ID;

                foreach (var item in userMachineTypeModel.MachineTypeIDs)
                {
                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("MachineTypeID", item.ToString()));
                    filters.Add(new QueryFilter("UserID", userMachineTypeModel.UserID.ToString()));
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    string exist = _userMachineTypeAppService.Get(filters, true);
                    if (!string.IsNullOrEmpty(exist))
                    {
                        SetFalseTempData(UIResources.UserMachineTypeExist);
                        return RedirectToAction("Index");
                    }
                }

                foreach (var machineTypeId in userMachineTypeModel.MachineTypeIDs)
                {
                    UserMachineTypeModel newEntity = new UserMachineTypeModel();
                    newEntity.UserID = userMachineTypeModel.UserID;
                    newEntity.MachineTypeID = machineTypeId;
                    newEntity.ModifiedBy = AccountName;
                    newEntity.ModifiedDate = DateTime.Now;

                    string data = JsonHelper<UserMachineTypeModel>.Serialize(newEntity);

                    _userMachineTypeAppService.Add(data);
                }

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
            UserMachineTypeModel userMachineType = GetUserMachineType(id);
            string userData = _userAppService.GetById(userMachineType.UserID);
            UserModel userModel = userData.DeserializeToUser();
            string emp = _empAppService.GetBy("EmployeeID", userModel.EmployeeID);
            userMachineType.FullName = emp.DeserializeToEmployee().FullName;

            string skills = _userMachineTypeAppService.FindByNoTracking("UserID", userMachineType.UserID.ToString(), true);
            List<UserMachineTypeModel> skillList = skills.DeserializeToUserMachineTypeList();
            List<long> machineTypeList = skillList.Select(c => c.MachineTypeID).Distinct().ToList();

            ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService, machineTypeList);

            string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
            List<ReferenceDetailModel> machineTypeReferenceList = mts.DeserializeToRefDetailList();

            string trainings = _trainingAppService.GetAll();
            List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();

            var trainingEmpList = trainingList.Where(x => x.EmployeeID == userModel.EmployeeID).ToList();

            string machineList = string.Empty;
            foreach (var training in trainingEmpList)
            {
                if (training.MachineTypeID.HasValue)
                {
                    var mt = machineTypeReferenceList.Where(x => x.ID == training.MachineTypeID.Value).FirstOrDefault();
                    if (mt != null && !machineList.Contains(mt.Code))
                    {
                        if (machineList == string.Empty)
                            machineList = mt.Code;
                        else
                            machineList += ", " + mt.Code;
                    }
                }
            }

            userMachineType.Competency = machineList;

            return PartialView(userMachineType);
        }

        [HttpPost]
        public ActionResult Edit(UserMachineTypeModel userMachineTypeModel)
        {
            try
            {
                ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                userMachineTypeModel.ModifiedBy = AccountName;
                userMachineTypeModel.ModifiedDate = DateTime.Now;

                string skills = _userMachineTypeAppService.FindByNoTracking("UserID", userMachineTypeModel.UserID.ToString(), true);
                List<UserMachineTypeModel> skillModelList = skills.DeserializeToUserMachineTypeList();


                foreach (var item in skillModelList)
                {
                    if (!userMachineTypeModel.MachineTypeIDs.Any(x => x == item.MachineTypeID))
                    {
                        // remove if not selected						
                        _userMachineTypeAppService.Remove(item.ID);
                    }
                }

                foreach (var item in userMachineTypeModel.MachineTypeIDs)
                {
                    if (!skillModelList.Any(x => x.MachineTypeID == item))
                    {
                        UserMachineTypeModel newEntity = new UserMachineTypeModel();
                        newEntity.UserID = userMachineTypeModel.UserID;
                        newEntity.MachineTypeID = item;
                        newEntity.ModifiedBy = AccountName;
                        newEntity.ModifiedDate = DateTime.Now;

                        string data = JsonHelper<UserMachineTypeModel>.Serialize(newEntity);

                        _userMachineTypeAppService.Add(data);
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

        public ActionResult Delete(int id)
        {
            return View();
        }

        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                List<UserMachineTypeModel> userMachineTypes = GetUserMachineTypeListByUserID(id);
                foreach (var item in userMachineTypes)
                {
                    _userMachineTypeAppService.Remove(item.ID);
                }

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllFromTrainings()
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
                string users = _userAppService.GetAll(true);
                List<UserModel> userList = users.DeserializeToUserList();
                userList = userList.Where(x => x.IsActive && x.IsFast).ToList();

                string emps = _empAppService.GetAll(true);
                List<EmployeeModel> empList = emps.DeserializeToEmployeeList();

                string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
                List<ReferenceDetailModel> machineTypeList = mts.DeserializeToRefDetailList();

                string trainingTitles = _trainingTitleAppService.GetAll();
                List<TrainingTitleModel> trainingTitleList = trainingTitles.DeserializeToTrainingTitleList();

                string trainingMachineTypes = _trainingTitleMachineTypeAppService.GetAll();
                List<TrainingTitleMachineTypeModel> trainingMachineTypeList = trainingMachineTypes.DeserializeToTrainingTitleMachineTypeList();

                string trainings = _trainingAppService.GetAll();
                List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();
                List<string> userTrainingList = trainingList.Select(c => c.EmployeeID).Distinct().ToList();

                userList = userList.Where(x => userTrainingList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

                bool isLoaded = false;

                List<UserMachineTypeModel> result = new List<UserMachineTypeModel>();
                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
                {
                    isLoaded = true;
                    Dictionary<long, string> trainingSkillMap = new Dictionary<long, string>();
                    foreach (var user in userList)
                    {
                        EmployeeModel emp = empList.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                        UserMachineTypeModel newEntity = new UserMachineTypeModel();
                        newEntity.UserName = user.UserName;
                        newEntity.EmployeeID = user.EmployeeID;
                        newEntity.FullName = emp == null ? "" : emp.FullName;
                        newEntity.PositionDesc = emp == null ? "" : emp.PositionDesc;

                        List<TrainingModel> trainingTempList = trainingList.Where(x => x.EmployeeID.Trim() == user.EmployeeID).ToList();

                        foreach (var training in trainingTempList)
                        {
                            TrainingTitleModel trainingTitleTemp = trainingTitleList.Where(x => x.Title == training.TrainingTitle).FirstOrDefault();
                            if (trainingTitleTemp != null)
                            {
                                List<TrainingTitleMachineTypeModel> trainingMTList = trainingMachineTypeList.Where(x => x.TrainingTitleID == trainingTitleTemp.ID).ToList();

                                string machineTypeListRef = "";
                                if (!trainingSkillMap.TryGetValue(trainingTitleTemp.ID, out machineTypeListRef))
                                {
                                    machineTypeListRef = GetMachineTypeList(machineTypeList, trainingMTList, machineTypeListRef);
                                    trainingSkillMap.Add(trainingTitleTemp.ID, machineTypeListRef);
                                }

                                if (string.IsNullOrEmpty(newEntity.Skills))
                                    newEntity.Skills = machineTypeListRef;
                                else
                                {
                                    if (!string.IsNullOrEmpty(machineTypeListRef) && !newEntity.Skills.Contains(machineTypeListRef))
                                        newEntity.Skills += ", " + machineTypeListRef;
                                }
                            }
                        }

                        result.Add(newEntity);
                    }
                }
                else
                {
                    foreach (var user in userList)
                    {
                        EmployeeModel emp = empList.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                        UserMachineTypeModel newEntity = new UserMachineTypeModel();
                        newEntity.UserName = user.UserName;
                        newEntity.EmployeeID = user.EmployeeID;
                        newEntity.FullName = emp == null ? "" : emp.FullName;
                        newEntity.PositionDesc = emp == null ? "" : emp.PositionDesc;

                        result.Add(newEntity);
                    }
                }

                int recordsTotal = result.Count();

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => m.UserName != null && m.UserName.ToLower().Contains(searchValue.ToLower()) ||
                                               m.EmployeeID != null && m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                               m.PositionDesc != null && m.PositionDesc.ToLower().Contains(searchValue.ToLower()) ||
                                               m.FullName != null && m.FullName.ToLower().Contains(searchValue.ToLower()) ||
                                               m.Skills != null && m.Skills.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                result = result.OrderBy(x => x.EmployeeID).ToList();
                                break;
                            case "username":
                                result = result.OrderBy(x => x.UserName).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderBy(x => x.PositionDesc).ToList();
                                break;
                            case "fullname":
                                result = result.OrderBy(x => x.FullName).ToList();
                                break;
                            case "skills":
                                result = result.OrderBy(x => x.Skills).ToList();
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
                                result = result.OrderByDescending(x => x.EmployeeID).ToList();
                                break;
                            case "username":
                                result = result.OrderByDescending(x => x.UserName).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderByDescending(x => x.PositionDesc).ToList();
                                break;
                            case "fullname":
                                result = result.OrderByDescending(x => x.FullName).ToList();
                                break;
                            case "skills":
                                result = result.OrderByDescending(x => x.Skills).ToList();
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

                if (!isLoaded)
                {
                    Dictionary<long, string> trainingSkillMap = new Dictionary<long, string>();
                    foreach (var user in data)
                    {
                        List<TrainingModel> trainingTempList = trainingList.Where(x => x.EmployeeID.Trim() == user.EmployeeID).ToList();

                        foreach (var training in trainingTempList)
                        {
                            TrainingTitleModel trainingTitleTemp = trainingTitleList.Where(x => x.Title == training.TrainingTitle).FirstOrDefault();
                            if (trainingTitleTemp != null)
                            {
                                List<TrainingTitleMachineTypeModel> trainingMTList = trainingMachineTypeList.Where(x => x.TrainingTitleID == trainingTitleTemp.ID).ToList();

                                string machineTypeListRef = "";
                                if (!trainingSkillMap.TryGetValue(trainingTitleTemp.ID, out machineTypeListRef))
                                {
                                    machineTypeListRef = GetMachineTypeList(machineTypeList, trainingMTList, machineTypeListRef);
                                    trainingSkillMap.Add(trainingTitleTemp.ID, machineTypeListRef);
                                }

                                if (string.IsNullOrEmpty(user.Skills))
                                    user.Skills = machineTypeListRef;
                                else
                                {
                                    if (!string.IsNullOrEmpty(machineTypeListRef) && !user.Skills.Contains(machineTypeListRef))
                                        user.Skills += ", " + machineTypeListRef;
                                }
                            }
                        }
                    }
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<UserMachineTypeModel>() }, JsonRequestBehavior.AllowGet);
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
                string userMachineTypeList = _userMachineTypeAppService.GetAll(true);
                List<UserMachineTypeModel> userMachineTypeModelList = userMachineTypeList.DeserializeToUserMachineTypeList();

                string mts = _referenceAppService.GetDetailAll(ReferenceEnum.MachineType, true);
                List<ReferenceDetailModel> machineTypeReferenceList = mts.DeserializeToRefDetailList();

                string trainings = _trainingAppService.GetAll();
                List<TrainingModel> trainingList = trainings.DeserializeToTrainingList();

                string userList = _userAppService.GetAll(true);
                List<UserModel> userModelList = userList.DeserializeToUserList();

                string empList = _empAppService.GetAll();
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                var machineTypeList = DropDownHelper.BindDropDownMachineType(_referenceAppService);

                int recordsTotal = userMachineTypeModelList.Count();
                bool isDataComplete = false;
                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                List<UserMachineTypeModel> result = new List<UserMachineTypeModel>();
                Dictionary<long, string> machineTypeMap = new Dictionary<long, string>();

                foreach (var item in userMachineTypeModelList)
                {
                    UserMachineTypeModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
                    if (exist == null)
                    {
                        UserModel userModel = userModelList.Where(x => x.ID == item.UserID).FirstOrDefault();
                        if (userModel != null)
                        {
                            EmployeeModel empModel = empModelList.Where(x => x.EmployeeID == userModel.EmployeeID).FirstOrDefault();
                            if (empModel != null)
                            {
                                item.UserName = userModel.UserName;
                                item.EmployeeID = empModel.EmployeeID;
                                item.FullName = empModel.FullName;
                                item.PositionDesc = empModel.PositionDesc;
                                var mt = machineTypeList.Where(x => x.Value == item.MachineTypeID.ToString()).FirstOrDefault();
                                item.Skills = mt == null ? string.Empty : mt.Text;

                                result.Add(item);
                            }
                        }
                    }
                    else
                    {
                        var mt = machineTypeList.Where(x => x.Value == item.MachineTypeID.ToString()).FirstOrDefault();
                        if (mt != null)
                            exist.Skills = exist.Skills + ", " + mt.Text;
                    }
                }

                if (!string.IsNullOrEmpty(searchValue) || (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDir)))
                {
                    isDataComplete = true;

                    foreach (var item in result)
                    {
                        var trainingEmpList = trainingList.Where(x => x.EmployeeID == item.EmployeeID).ToList();

                        string machineList = string.Empty;
                        foreach (var training in trainingEmpList)
                        {
                            if (training.MachineTypeID.HasValue)
                            {
                                var mt = machineTypeReferenceList.Where(x => x.ID == training.MachineTypeID.Value).FirstOrDefault();
                                if (mt != null && !machineList.Contains(mt.Code))
                                {
                                    if (machineList == string.Empty)
                                        machineList = mt.Code;
                                    else
                                        machineList += ", " + mt.Code;
                                }
                            }
                        }

                        item.Competency = machineList;
                    }
                }

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => m.UserName.ToLower().Contains(searchValue.ToLower()) ||
                                               m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                               m.PositionDesc.ToLower().Contains(searchValue.ToLower()) ||
                                               m.FullName.ToLower().Contains(searchValue.ToLower()) ||
                                               m.Skills.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!string.IsNullOrEmpty(sortColumn))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "employeeid":
                                result = result.OrderBy(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                result = result.OrderBy(x => x.FullName).ToList();
                                break;
                            case "skills":
                                result = result.OrderBy(x => x.Skills).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderBy(x => x.PositionDesc).ToList();
                                break;
                            case "competency":
                                result = result.OrderBy(x => x.Competency).ToList();
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
                                result = result.OrderByDescending(x => x.EmployeeID).ToList();
                                break;
                            case "fullname":
                                result = result.OrderByDescending(x => x.FullName).ToList();
                                break;
                            case "skills":
                                result = result.OrderByDescending(x => x.Skills).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderByDescending(x => x.PositionDesc).ToList();
                                break;
                            case "competency":
                                result = result.OrderByDescending(x => x.Competency).ToList();
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

                if (!isDataComplete)
                {
                    foreach (var item in data)
                    {
                        var trainingEmpList = trainingList.Where(x => x.EmployeeID == item.EmployeeID).ToList();

                        string machineList = string.Empty;
                        foreach (var training in trainingEmpList)
                        {
                            if (training.MachineTypeID.HasValue)
                            {
                                var mt = machineTypeReferenceList.Where(x => x.ID == training.MachineTypeID.Value).FirstOrDefault();
                                if (mt != null && !machineList.Contains(mt.Code))
                                {
                                    if (machineList == string.Empty)
                                        machineList = mt.Code;
                                    else
                                        machineList += ", " + mt.Code;
                                }
                            }
                        }

                        item.Competency = machineList;
                    }
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<UserMachineTypeModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region ::Private Methods::
        private static string GetMachineTypeList(List<ReferenceDetailModel> machineTypeList, List<TrainingTitleMachineTypeModel> trainingMTList, string machineTypeListRef)
        {
            foreach (var item in trainingMTList)
            {
                var temp = machineTypeList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
                if (temp == null)
                    continue;

                if (string.IsNullOrEmpty(machineTypeListRef))
                    machineTypeListRef = temp.Code;
                else
                {
                    if (!string.IsNullOrEmpty(temp.Code) && !machineTypeListRef.Contains(temp.Code))
                        machineTypeListRef += ", " + temp.Code;
                }
            }

            return machineTypeListRef;
        }

        private UserMachineTypeModel GetUserMachineType(long userMachineTypeID)
        {
            string user = _userMachineTypeAppService.GetById(userMachineTypeID, true);
            UserMachineTypeModel model = user.DeserializeToUserMachineType();

            return model;
        }

        private List<UserMachineTypeModel> GetUserMachineTypeListByUserID(long userID)
        {
            string usermachines = _userMachineTypeAppService.FindByNoTracking("UserID", userID.ToString(), true);
            List<UserMachineTypeModel> models = usermachines.DeserializeToUserMachineTypeList();

            return models;
        }

        private string GetMachineType(long machineTypeID)
        {
            string machineType = _referenceAppService.GetDetailById(machineTypeID, true);
            ReferenceDetailModel machineTypeModel = machineType.DeserializeToRefDetail();

            return machineTypeModel.Code;
        }

        #endregion
    }
}
