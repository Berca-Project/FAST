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
    [CustomAuthorize("usermachines")]
    public class UserMachineController : BaseController<UserMachineModel>
    {
        private readonly IUserMachineAppService _userMachineAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IUserAppService _userAppService;
        private readonly ILoggerAppService _logger;
        private readonly IEmployeeAppService _employeeService;
        private readonly IMenuAppService _menuService;
        private readonly ILocationAppService _locationAppService;

        public UserMachineController(
            ILocationAppService locationAppService,
            IUserMachineAppService userMachineAppService,
            IUserAppService userAppService,
            ILoggerAppService logger,
            IMachineAppService machineAppService,
            IEmployeeAppService empService,
            IMenuAppService menuService,
            IReferenceAppService referenceAppService)
        {
            _locationAppService = locationAppService;
            _userMachineAppService = userMachineAppService;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
            _menuService = menuService;
            _logger = logger;
            _machineAppService = machineAppService;
            _employeeService = empService;
        }

        // GET: UserMachine
        public ActionResult Index()
        {
            GetTempData();

            ViewBag.MachineList = DropDownHelper.BuildEmptyList();
            ViewBag.UserList = DropDownHelper.BuildEmptyList();

            UserMachineModel model = new UserMachineModel();
            model.Access = GetAccess(WebConstants.MenuSlug.USER_MACHINE, _menuService);

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

            string emplist = _employeeService.Find(filters);
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

        public ActionResult ExportExcel()
        {
            try
            {
                // Getting all data    			
                string userMachineList = _userMachineAppService.GetAll(true);
                List<UserMachineModel> userMachineModelList = userMachineList.DeserializeToUserMachineList();

                List<UserMachineModel> result = new List<UserMachineModel>();

                foreach (var item in userMachineModelList)
                {
                    UserMachineModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
                    if (exist == null)
                    {
                        string user = _userAppService.GetById(item.UserID, true);
                        UserModel userModel = user.DeserializeToUser();
                        string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID, true);
                        item.Employee = emp.DeserializeToEmployee();
                        item.MachineList = GetMachine(item.MachineID);

                        result.Add(item);
                    }
                    else
                    {
                        exist.MachineList = exist.MachineList + ", " + GetMachine(item.MachineID);
                    }
                }

                byte[] excelData = ExcelGenerator.ExportMasterUserMachine(result, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-UserMachine.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        // GET: UserMachine/Create
        public ActionResult Create()
        {
            ViewBag.MachineList = DropDownHelper.BindDropDownMultiMachine(_machineAppService);
            ViewBag.UserList = GetEmployeeList();

            UserMachineModel model = new UserMachineModel();
            model.Access = GetAccess(WebConstants.MenuSlug.USER_MACHINE, _menuService);

            return PartialView(model);
        }

        // POST: UserMachine/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(UserMachineModel userMachineModel)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BuildEmptyList();
                ViewBag.UserList = DropDownHelper.BuildEmptyList();

                userMachineModel.Access = GetAccess(WebConstants.MenuSlug.USER_MACHINE, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(userMachineModel.EmployeeID))
                {
                    SetFalseTempData("Please select user first");
                    return RedirectToAction("Index");
                }

                if (userMachineModel.MachineIDs == null || userMachineModel.MachineIDs.Count() == 0)
                {
                    SetFalseTempData("Please select machine first");
                    return RedirectToAction("Index");
                }

                string userT = _userAppService.GetBy("EmployeeID", userMachineModel.EmployeeID, true);
                UserModel userModelT = userT.DeserializeToUser();

                userMachineModel.UserID = userModelT.ID;

                foreach (var item in userMachineModel.MachineIDs)
                {
                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("MachineID", item.ToString()));
                    filters.Add(new QueryFilter("UserID", userMachineModel.UserID.ToString()));
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    string exist = _userMachineAppService.Get(filters, true);
                    if (!string.IsNullOrEmpty(exist))
                    {
                        string user = _userAppService.GetById(userMachineModel.UserID, true);
                        UserModel userModel = user.DeserializeToUser();
                        string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID);
                        EmployeeModel empModel = emp.DeserializeToEmployee();

                        string machine = _machineAppService.GetById(item, true);
                        MachineModel machineModel = machine.DeserializeToMachine();

                        SetFalseTempData(string.Format(UIResources.DataExist, empModel.FullName, machineModel.Code));
                        return RedirectToAction("Index");
                    }
                }

                foreach (var machineId in userMachineModel.MachineIDs)
                {
                    UserMachineModel newEntity = new UserMachineModel();
                    newEntity.UserID = userMachineModel.UserID;
                    newEntity.MachineID = machineId;
                    newEntity.ModifiedBy = AccountName;
                    newEntity.ModifiedDate = DateTime.Now;

                    string data = JsonHelper<UserMachineModel>.Serialize(newEntity);

                    _userMachineAppService.Add(data);
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

        // GET: UserMachine/Edit/5
        public ActionResult Edit(int id)
        {
            UserMachineModel userMachine = GetUserMachine(id);
            string machines = _userMachineAppService.FindByNoTracking("UserID", userMachine.UserID.ToString(), true);
            List<UserMachineModel> machineList = machines.DeserializeToUserMachineList();
            List<long> machineIDList = machineList.Select(c => c.MachineID).Distinct().ToList();

            ViewBag.MachineList = DropDownHelper.BindDropDownMultiMachine(_machineAppService, machineIDList);
            ViewBag.UserList = DropDownHelper.BuildEmptyList();

            return PartialView(userMachine);
        }

        // POST: UserMachine/Edit/5
        [HttpPost]
        public ActionResult Edit(UserMachineModel userMachineModel)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BindDropDownMultiMachine(_machineAppService);
                ViewBag.UserList = GetEmployeeList();

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                userMachineModel.ModifiedBy = AccountName;
                userMachineModel.ModifiedDate = DateTime.Now;

                string machines = _userMachineAppService.FindByNoTracking("UserID", userMachineModel.UserID.ToString(), true);
                List<UserMachineModel> machineList = machines.DeserializeToUserMachineList();


                foreach (var item in machineList)
                {
                    if (!userMachineModel.MachineIDs.Any(x => x == item.MachineID))
                    {
                        // remove if not selected						
                        _userMachineAppService.Remove(item.ID);
                    }
                }

                foreach (var item in userMachineModel.MachineIDs)
                {
                    if (!machineList.Any(x => x.MachineID == item))
                    {
                        UserMachineModel newEntity = new UserMachineModel();
                        newEntity.UserID = userMachineModel.UserID;
                        newEntity.MachineID = item;
                        newEntity.ModifiedBy = AccountName;
                        newEntity.ModifiedDate = DateTime.Now;

                        string data = JsonHelper<UserMachineModel>.Serialize(newEntity);

                        _userMachineAppService.Add(data);
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

        // GET: UserMachine/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: UserMachine/Delete/5
        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                List<UserMachineModel> userMachines = GetUserMachineListByUserID(id);
                foreach (var item in userMachines)
                {
                    _userMachineAppService.Remove(item.ID);
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
                string userMachineList = _userMachineAppService.GetAll(true);
                List<UserMachineModel> userMachineModelList = userMachineList.DeserializeToUserMachineList();

                List<UserMachineModel> result = new List<UserMachineModel>();

                Dictionary<long, string> locationMap = new Dictionary<long, string>();
                foreach (var item in userMachineModelList)
                {
                    UserMachineModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
                    if (exist == null)
                    {
                        string user = _userAppService.GetById(item.UserID, true);
                        UserModel userModel = user.DeserializeToUser();
                        if (userModel.LocationID.HasValue)
                        {
                            if (locationMap.ContainsKey(userModel.LocationID.Value))
                            {
                                string loc;
                                locationMap.TryGetValue(userModel.LocationID.Value, out loc);
                                item.Location = loc;
                            }
                            else
                            {
                                item.Location = _locationAppService.GetLocationFullCode(userModel.LocationID.Value);
                                locationMap.Add(userModel.LocationID.Value, item.Location);
                            }
                        }

                        string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID, true);
                        item.Employee = emp.DeserializeToEmployee();
                        item.MachineList = GetMachine(item.MachineID);

                        result.Add(item);
                    }
                    else
                    {
                        string user = _userAppService.GetById(item.UserID, true);
                        UserModel userModel = user.DeserializeToUser();
                        if (userModel.LocationID.HasValue)
                        {
                            if (locationMap.ContainsKey(userModel.LocationID.Value))
                            {
                                string loc;
                                locationMap.TryGetValue(userModel.LocationID.Value, out loc);
                                exist.Location = loc;
                            }
                            else
                            {
                                exist.Location = _locationAppService.GetLocationFullCode(userModel.LocationID.Value);
                                locationMap.Add(userModel.LocationID.Value, item.Location);
                            }
                        }

                        exist.MachineList = string.IsNullOrEmpty(exist.MachineList) ? GetMachine(item.MachineID) : (exist.MachineList + ", " + GetMachine(item.MachineID));
                    }
                }

                int recordsTotal = result.Count();

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => m.Employee != null && m.Employee.FullName != null && m.Employee.FullName.ToLower().Contains(searchValue) ||
                                               m.EmployeeID != null && m.Employee.EmployeeID != null && m.Employee.EmployeeID.ToLower().Contains(searchValue) ||
                                               m.Employee != null && m.Employee.PositionDesc != null && m.Employee.PositionDesc.ToLower().Contains(searchValue) ||
                                               m.Employee != null && m.Employee.BaseTownLocation != null && m.Employee.BaseTownLocation.ToLower().Contains(searchValue) ||
                                               m.MachineList != null && m.MachineList.ToLower().Contains(searchValue)).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "userfullname":
                                result = result.OrderBy(x => x.Employee.FullName).ToList();
                                break;
                            case "machine":
                                result = result.OrderBy(x => x.Machine).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderBy(x => x.Employee.PositionDesc).ToList();
                                break;
                            case "employeeid":
                                result = result.OrderBy(x => x.Employee.EmployeeID).ToList();
                                break;
                            case "location":
                                result = result.OrderBy(x => x.Location).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "userfullname":
                                result = result.OrderByDescending(x => x.Employee.FullName).ToList();
                                break;
                            case "machine":
                                result = result.OrderByDescending(x => x.Machine).ToList();
                                break;
                            case "positiondesc":
                                result = result.OrderByDescending(x => x.Employee.PositionDesc).ToList();
                                break;
                            case "employeeid":
                                result = result.OrderByDescending(x => x.Employee.EmployeeID).ToList();
                                break;
                            case "location":
                                result = result.OrderByDescending(x => x.Location).ToList();
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

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<UserMachineModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        private UserMachineModel GetUserMachine(long userMachineID)
        {
            string user = _userMachineAppService.GetById(userMachineID, true);
            UserMachineModel model = user.DeserializeToUserMachine();
            string userData = _userAppService.GetById(model.UserID);
            UserModel userModel = userData.DeserializeToUser();
            string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID);
            model.Employee = emp.DeserializeToEmployee();

            return model;
        }

        private List<UserMachineModel> GetUserMachineListByUserID(long userID)
        {
            string usermachines = _userMachineAppService.FindByNoTracking("UserID", userID.ToString(), true);
            List<UserMachineModel> models = usermachines.DeserializeToUserMachineList();

            return models;
        }

        private string GetMachine(long machineID)
        {
            string machine = _machineAppService.GetById(machineID, true);
            MachineModel machineModel = machine.DeserializeToMachine();

            return machineModel.Code;
        }

        private List<SelectListItem> GetEmployeeList()
        {
            List<SelectListItem> result = DropDownHelper.BindDropDownEmployee(_employeeService);
            //string usermachines = _userMachineAppService.GetAll(true);
            //List<UserMachineModel> usermachineList = usermachines.DeserializeToUserMachineList();
            //result = result.Where(x => !usermachineList.Any(y => y.UserID.ToString() == x.Value)).ToList();
            return result;
        }
    }
}
