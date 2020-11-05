using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Web.Mvc;


namespace Fast.Web.Controllers.Master
{
    [CustomAuthorize("user")]
    public class UserController : BaseController<UserModel>
    {
        private readonly IUserAppService _userAppService;
        private readonly IUserRoleAppService _userRoleAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IRoleAppService _roleAppService;
        private readonly IJobTitleAppService _jobTitleAppService;
        private readonly ILoggerAppService _logger;
        private readonly IEmployeeAppService _empService;
        private readonly IEmployeeAllAppService _empAllService;
        private readonly IMenuAppService _menuService;
        private readonly IUserLogAppService _userLogAppService;

        public UserController(IUserAppService userAppService,
            ILocationAppService locationAppService,
            IRoleAppService roleAppService,
            ILoggerAppService logger,
            IJobTitleAppService jobTitleAppService,
            IUserLogAppService userLogAppService,
            IEmployeeAllAppService empAllService,
            IEmployeeAppService empService,
            IMenuAppService menuService,
            IUserRoleAppService userRoleAppService,
            IReferenceAppService referenceAppService)
        {
            _userRoleAppService = userRoleAppService;
            _jobTitleAppService = jobTitleAppService;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
            _locationAppService = locationAppService;
            _roleAppService = roleAppService;
            _logger = logger;
            _userLogAppService = userLogAppService;
            _empAllService = empAllService;
            _empService = empService;
            _menuService = menuService;
        }

        [HttpPost]
        public JsonResult AutoComplete(string prefix, string searchBy)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            if (searchBy.Equals("ID"))
                filters.Add(new QueryFilter("EmployeeID", prefix, Operator.StartsWith));
            else
                filters.Add(new QueryFilter("FullName", prefix, Operator.Contains));

            string emp = _empAllService.Find(filters);
            List<EmployeeModel> empList = emp.DeserializeToEmployeeList();

            List<EmployeeModel> result = new List<EmployeeModel>();

            // exclude if already exist in the user table
            foreach (var item in empList)
            {
                string user = _userAppService.GetBy("EmployeeID", item.EmployeeID, true);
                if (string.IsNullOrEmpty(user))
                {
                    result.Add(item);
                }
            }

            if (searchBy.Equals("ID"))
                result = result.OrderBy(x => x.EmployeeID).ToList();
            else
                result = result.OrderBy(x => x.FullName).ToList();

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult AutoCompleteSpv(string prefix)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            if (prefix.All(Char.IsDigit))
                filters.Add(new QueryFilter("EmployeeID", prefix, Operator.StartsWith));
            else
                filters.Add(new QueryFilter("FullName", prefix, Operator.Contains));

            string emplist = _empAllService.Find(filters);
            List<EmployeeModel> empModelList = emplist.DeserializeToEmployeeList();

            string empAlllist = _empAllService.GetAll();
            List<EmployeeModel> empAllListModel = empAlllist.DeserializeToEmployeeList();
            //fery edit
            empAllListModel = empAllListModel.Where(x => !string.IsNullOrEmpty(x.EmployeeID) && !x.EmployeeID.StartsWith("00")).ToList();

            //List<string> spvList = empModelList.Select(x => x.ReportToID1).Distinct().ToList();
            //spvList = spvList.Where(x => !string.IsNullOrEmpty(x)).ToList();

            List<EmployeeModel> spvModelList = empModelList.Where(x => empAllListModel.Any(y => y.EmployeeID.Trim() == x.EmployeeID.Trim())).ToList();

            if (prefix.All(Char.IsDigit))
            {
                spvModelList = spvModelList.OrderBy(x => x.EmployeeID).ToList();
            }
            else
            {
                spvModelList = spvModelList.OrderBy(x => x.FullName).ToList();
            }

            return Json(spvModelList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ExportExcel()
        {
            try
            {
                // Getting all data    			
                string users = _userAppService.FindBy("IsFast", "true", true);
                List<UserModel> userModelList = users.DeserializeToUserList();

                // Construct custom attributes
                string jobTitles = _jobTitleAppService.GetAll(true);
                List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

                string employees = _empService.GetAll(true);
                List<EmployeeModel> employeeList = employees.DeserializeToEmployeeList();

                List<ReferenceDetailModel> canteenModelList = new List<ReferenceDetailModel>();
                Dictionary<long, string> canteenMap = GetCanteenList(ref canteenModelList);

                // Construct custom attributes
                foreach (var item in userModelList)
                {
                    JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                    item.RoleName = jt == null ? string.Empty : jt.RoleName;
                    item.JobTitle = jt == null ? string.Empty : jt.Title;

                    if (!string.IsNullOrEmpty(item.EmployeeID))
                    {
                        EmployeeModel emp = employeeList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                        item.Employee = emp == null ? item.Employee : emp;

                        if (emp != null)
                        {
                            if (!string.IsNullOrEmpty(item.Employee.ReportToID1))
                            {
                                if (!string.IsNullOrEmpty(item.Employee.ReportToID1.Trim()))
                                {
                                    EmployeeModel spv = employeeList.Where(x => x.EmployeeID.Trim() == item.Employee.ReportToID1.Trim()).FirstOrDefault();
                                    item.SupervisorName = spv == null ? string.Empty : spv.FullName;
                                }
                            }

                            if (!string.IsNullOrEmpty(item.Employee.ReportToID2))
                            {
                                if (!string.IsNullOrEmpty(item.Employee.ReportToID2.Trim()))
                                {
                                    EmployeeModel mgr = employeeList.Where(x => x.EmployeeID.Trim() == item.Employee.ReportToID2.Trim()).FirstOrDefault();
                                    item.ManagerName = mgr == null ? string.Empty : mgr.FullName;
                                }
                            }

                            if (item.CanteenID.HasValue)
                            {
                                item.Canteen = GetCanteen(item.CanteenID, canteenMap);
                            }
                        }
                    }
                }

                byte[] excelData = ExcelGenerator.ExportMasterUser(userModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Users.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        //excel mass upload / change group
        public ActionResult GenerateExcel(string supervisorID, string locationID, string locationType)
        {
            try
            {
                // Getting all data    			
                string empList = _empService.GetAll(true);
                List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsFast", "true"));
                filters.Add(new QueryFilter("IsActive", "true"));

                string userList = _userAppService.Find(filters);
                List<UserModel> userModelList = userList.DeserializeToUserList();

                List<EmployeeModel> result = new List<EmployeeModel>();
                List<ReferenceDetailModel> canteenModelList = new List<ReferenceDetailModel>();
                Dictionary<long, string> canteenMap = GetCanteenList(ref canteenModelList);

                if (!string.IsNullOrEmpty(supervisorID))
                {
                    var empModelListSpv = empModelList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == supervisorID || x.EmployeeID.Trim() == supervisorID).ToList();
                    var userModelListSpv = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == supervisorID || x.EmployeeID.Trim() == supervisorID).ToList();

                    if (!string.IsNullOrEmpty(locationID))
                    {
                        List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(locationID), locationType);
                        userModelListSpv = userModelListSpv.Where(x => locIDList.Any(y => y == x.LocationID.Value) || !x.LocationID.HasValue).ToList();
                    }

                    foreach (var user in userModelListSpv)
                    {
                        EmployeeModel emp = empModelListSpv.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                        if (emp != null)
                        {
                            emp.Location = user.Location;
                            emp.Canteen = GetCanteen(user.CanteenID, canteenMap);
                            emp.HomeTownLocation = user.SupervisorName;
                            if (!result.Any(x => x.EmployeeID.Trim() == emp.EmployeeID.Trim())) result.Add(emp);

                            var userModelListSpv1 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user.EmployeeID && x.SupervisorID.Trim() != supervisorID).ToList();
                            foreach (var user1 in userModelListSpv1)
                            {
                                EmployeeModel emp1 = empModelList.Where(x => x.EmployeeID.Trim() == user1.EmployeeID.Trim()).FirstOrDefault();
                                if (emp1 != null)
                                {
                                    emp1.Location = user1.Location;
                                    emp1.Canteen = GetCanteen(user1.CanteenID, canteenMap);
                                    if (!result.Any(x => x.EmployeeID.Trim() == emp1.EmployeeID.Trim())) result.Add(emp1);

                                    var userModelListSpv2 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user1.EmployeeID).ToList();
                                    foreach (var user2 in userModelListSpv2)
                                    {
                                        EmployeeModel emp2 = empModelList.Where(x => x.EmployeeID.Trim() == user2.EmployeeID.Trim()).FirstOrDefault();
                                        if (emp2 != null)
                                        {
                                            emp2.Location = user2.Location;
                                            emp2.Canteen = GetCanteen(user2.CanteenID, canteenMap);
                                            emp2.HomeTownLocation = user2.SupervisorName;
                                            if (!result.Any(x => x.EmployeeID.Trim() == emp2.EmployeeID.Trim())) result.Add(emp2);

                                            var userModelListSpv3 = userModelList.Where(x => x.SupervisorID != null && x.SupervisorID.Trim() == user2.EmployeeID).ToList();
                                            foreach (var user3 in userModelListSpv3)
                                            {
                                                EmployeeModel emp3 = empModelList.Where(x => x.EmployeeID.Trim() == user3.EmployeeID.Trim()).FirstOrDefault();
                                                if (emp3 != null)
                                                {
                                                    emp3.Location = user3.Location;
                                                    emp3.Canteen = GetCanteen(user3.CanteenID, canteenMap);
                                                    emp3.HomeTownLocation = user3.SupervisorName;
                                                    if (!result.Any(x => x.EmployeeID.Trim() == emp3.EmployeeID.Trim())) result.Add(emp3);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(locationID))
                {
                    List<long> locationIdList = _locationAppService.GetLocIDListByLocType(long.Parse(locationID), locationType);

                    var userModelList2 = userModelList.Where(x => !x.LocationID.HasValue || locationIdList.Any(y => y == x.LocationID)).ToList();

                    if (userModelList2.Count > 0)
                    {
                        foreach (var user in userModelList2)
                        {
                            EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                            if (emp != null)
                            {
                                emp.Location = user.Location;
                                emp.Canteen = GetCanteen(user.CanteenID, canteenMap);
                                emp.HomeTownLocation = user.SupervisorName;
                                result.Add(emp);
                            }
                        }
                    }
                }
                else
                {
                    foreach (var user in userModelList)
                    {
                        EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim() == user.EmployeeID.Trim()).FirstOrDefault();
                        if (emp != null)
                        {
                            emp.Location = user.Location;
                            emp.Canteen = GetCanteen(user.CanteenID, canteenMap);
                            emp.HomeTownLocation = user.SupervisorName;
                            result.Add(emp);
                        }
                    }
                }
                //panggil function excel generator
                byte[] excelData = ExcelGenerator.ExportUserGroupType(result, canteenModelList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Template-User-GroupType.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();

                ViewBag.Result = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        private string GetCanteen(long? canteenID, Dictionary<long, string> canteenMap)
        {
            string result = string.Empty;

            if (canteenID.HasValue)
                canteenMap.TryGetValue(canteenID.Value, out result);

            return result;
        }

        private Dictionary<long, string> GetCanteenList(ref List<ReferenceDetailModel> canteenModelList)
        {
            string reference = _referenceAppService.GetBy("Name", "Canteen", true);
            ReferenceModel refModel = reference.DeserializeToReference();

            string canteens = _referenceAppService.FindDetailBy("ReferenceID", refModel.ID, true);
            canteenModelList = canteens.DeserializeToRefDetailList();

            Dictionary<long, string> result = new Dictionary<long, string>();

            foreach (var item in canteenModelList)
            {
                result.Add(item.ID, item.Code);
            }

            return result;
        }

        private Dictionary<string, long> GetCanteenList()
        {
            string reference = _referenceAppService.GetBy("Name", "Canteen", true);
            ReferenceModel refModel = reference.DeserializeToReference();

            string canteens = _referenceAppService.FindDetailBy("ReferenceID", refModel.ID, true);
            List<ReferenceDetailModel> canteenModelList = canteens.DeserializeToRefDetailList();

            Dictionary<string, long> result = new Dictionary<string, long>();

            foreach (var item in canteenModelList)
            {
                result.Add(item.Code, item.ID);
            }

            return result;
        }

        private List<EmployeeModel> GetEmployeesFromAD(string keyword, string searchBy)
        {
            List<EmployeeModel> empResult = new List<EmployeeModel>();

            try
            {
                DirectoryEntry de = new DirectoryEntry(ConfigurationManager.AppSettings["DirectoryEntryPath"]);
                de.Username = ConfigurationManager.AppSettings["DirectoryEntryUsername"];
                de.Password = ConfigurationManager.AppSettings["DirectoryEntryPassword"];

                DirectorySearcher ds = new DirectorySearcher(de);

                // setup directory searcher
                ds.PageSize = 1000;
                ds.SearchScope = SearchScope.Subtree;
                ds.CacheResults = false;

                if (searchBy.Equals("ID"))
                {
                    if (keyword.Length < 8)
                    {
                        keyword += "*";
                    }
                    ds.Filter = "(&(objectCategory=person)(objectClass=user)(employeeID=" + keyword + "))";
                }
                else
                {
                    ds.Filter = "(&(objectCategory=person)(objectClass=user)(displayName=" + keyword + "*))";
                }

                Helper.LogErrorMessage("AD connection keyword: " + keyword, Server.MapPath("~/Uploads/"));

                SearchResultCollection results = ds.FindAll();

                foreach (SearchResult sr in results)
                {
                    DirectoryEntry user = sr.GetDirectoryEntry();

                    EmployeeModel result = new EmployeeModel();
                    result.UserName = GetUsername(user);
                    result.EmployeeID = user.Properties["employeeid"].Value.ToString().Trim();
                    result.FullName = user.Properties["displayName"].Value.ToString() + " (AD)";
                    //result.DepartmentDesc = user.Properties["department"].Value.ToString();
                    //result.BusinessUnit = user.Properties["company"].Value.ToString();
                    //result.PositionDesc = user.Properties["title"].Value.ToString();
                    //result.Email = user.Properties["mail"].Value.ToString();
                    //result.Phone = user.Properties["telephoneNumber"].Value.ToString();
                    //result.BusinessUnit = user.Properties["department"].Value.ToString();
                    ////result.ReportToID2 = user.Properties["manager"].Value.ToString();
                    ////result.BaseTownLocation = user.Properties["physicalDeliveryOfficeName"].Value.ToString();

                    empResult.Add(result);
                }
            }
            catch (Exception ex)
            {
                Helper.LogErrorMessage("AD connection error: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"));
            }

            return empResult;
        }

        private EmployeeModel GetEmployeesFromADByID(string empID)
        {
            try
            {
                DirectoryEntry de = new DirectoryEntry(ConfigurationManager.AppSettings["DirectoryEntryPath"]);
                de.Username = ConfigurationManager.AppSettings["DirectoryEntryUsername"];
                de.Password = ConfigurationManager.AppSettings["DirectoryEntryPassword"];

                DirectorySearcher ds = new DirectorySearcher(de);

                // setup directory searcher
                ds.PageSize = 1000;
                ds.SearchScope = SearchScope.Subtree;
                ds.CacheResults = false;

                ds.Filter = "(&(objectCategory=person)(objectClass=user)(employeeID=" + empID + "*))";

                Helper.LogErrorMessage("AD connection emp ID: " + empID, Server.MapPath("~/Uploads/"));

                SearchResult sr = ds.FindOne();

                DirectoryEntry user = sr.GetDirectoryEntry();
                if (user != null)
                {
                    EmployeeModel result = new EmployeeModel();
                    result.UserName = GetUsername(user);
                    result.EmployeeID = user.Properties["employeeid"].Value.ToString().Trim();
                    result.FullName = user.Properties["displayName"].Value.ToString();
                    result.DepartmentDesc = user.Properties["department"].Value.ToString();
                    //result.BusinessUnit = user.Properties["company"].Value.ToString();
                    result.PositionDesc = user.Properties["title"].Value.ToString();
                    result.Email = user.Properties["mail"].Value.ToString();
                    result.Phone = user.Properties["telephoneNumber"].Value.ToString();
                    //result.BusinessUnit = user.Properties["department"].Value.ToString();
                    //result.ReportToID2 = user.Properties["manager"].Value.ToString();
                    //result.BaseTownLocation = user.Properties["physicalDeliveryOfficeName"].Value.ToString();

                    return result;
                }
            }
            catch (Exception ex)
            {
                Helper.LogErrorMessage("AD connection error: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"));
            }

            return null;
        }

        private static string GetUsername(DirectoryEntry user)
        {
            string pmiuser = user.Properties["userPrincipalName"].Value.ToString();
            pmiuser = pmiuser.Replace("@PMINTL.NET", "");
            pmiuser = pmiuser.Replace("@pmintl.net", "");
            pmiuser = pmiuser.Replace("@pmi.com", "");
            pmiuser = "PMI\\" + pmiuser;
            return pmiuser;
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

        // GET: User
        public ActionResult Index()
        {
            GetTempData();

            UserModel usermodel = GetIndexModel();

            return View(usermodel);
        }

        private UserModel GetIndexModel()
        {
            ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
            ViewBag.CountryList = DropDownHelper.BuildEmptyList();
            ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
            ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
            ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
            ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
            ViewBag.GroupNameList = DropDownHelper.BuildEmptyList();
            ViewBag.SearchByList = DropDownHelper.BindDropDownSearchBy();

            UserModel usermodel = new UserModel();
            usermodel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

            return usermodel;
        }

        // GET: User/Create
        public ActionResult Create()
        {
            ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
            ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
            ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
            ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);
            ViewBag.SearchByList = DropDownHelper.BindDropDownSearchBy();
            ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService, true);

            UserModel usermodel = new UserModel();
            usermodel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

            return PartialView(usermodel);
        }

        // POST: User/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(UserModel userModel)
        {
            try
            {
                ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
                ViewBag.CountryList = DropDownHelper.BuildEmptyList();
                ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
                ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
                ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
                ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
                ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
                ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);
                ViewBag.SearchByList = DropDownHelper.BindDropDownSearchBy();
                ViewBag.CanteenList = DropDownHelper.BuildEmptyList();

                userModel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(ModelState.GetModelStateErrors());
                    return RedirectToAction("Index");
                }

                string user = _userAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
                if (!string.IsNullOrEmpty(user))
                {
                    SetFalseTempData("User already exist");
                    return RedirectToAction("Index");
                }

                EmployeeModel empModel = new EmployeeModel();
                string emp = _empAllService.GetBy("EmployeeID", userModel.EmployeeID);
                empModel = emp.DeserializeToEmployee();

                //userModel.UserName = "helmy";

                // set username with empid for temporary while waiting username info from AD
                userModel.UserName = userModel.EmployeeID;

                string spv = _empAllService.GetBy("EmployeeID", empModel.ReportToID1);
                if (spv != string.Empty)
                {
                    EmployeeModel spvModel = spv.DeserializeToEmployee();
                    userModel.SupervisorID = spvModel.EmployeeID;
                    userModel.SupervisorName = spvModel.FullName;
                }

                string position = _jobTitleAppService.GetBy("Title", empModel.PositionDesc);
                if (position == string.Empty)
                {
                    JobTitleModel jt = new JobTitleModel();
                    jt.RoleName = null;
                    jt.Code = "NEW";
                    jt.Title = empModel.PositionDesc;

                    string newTitle = JsonHelper<JobTitleModel>.Serialize(jt);

                    long jobTitleID = _jobTitleAppService.Add(newTitle);

                    userModel.JobTitleID = jobTitleID;
                }
                else
                {
                    JobTitleModel jt = position.DeserializeToJobTitle();
                    userModel.JobTitleID = jt.ID;
                }

                userModel.ModifiedBy = AccountName;
                userModel.ModifiedDate = DateTime.Now;
                userModel.IsActive = true;
                userModel.IsFast = true;

                if (userModel.SubDepartmentID != 0)
                {
                    userModel.LocationID = userModel.SubDepartmentID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.SubDepartmentID);
                }
                else if (userModel.DepartmentID != 0)
                {
                    userModel.LocationID = userModel.DepartmentID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.DepartmentID);
                }
                else
                {
                    userModel.LocationID = userModel.ProdCenterID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.ProdCenterID);
                }

                userModel.ID = _userAppService.AddModel(userModel);

                string empOld = _empService.FindByNoTracking("EmployeeID", userModel.EmployeeID);
                EmployeeModel empOldModel = empOld.DeserializeToEmployeeList().FirstOrDefault();
                if (empOldModel != null)
                {
                    empOldModel.GroupType = userModel.GroupType;
                    empOldModel.GroupName = userModel.GroupName;

                    string updateEmp = JsonHelper<EmployeeModel>.Serialize(empOldModel);

                    _empService.Update(updateEmp);
                }
                else
                {
                    empModel.GroupType = userModel.GroupType;
                    empModel.GroupName = userModel.GroupName;
                    empModel.ModifiedBy = AccountName;
                    empModel.ModifiedDate = DateTime.Now;

                    string addEmp = JsonHelper<EmployeeModel>.Serialize(empModel);

                    _empService.Add(addEmp);
                }

                //get username from AD
                EmployeeModel model = GetEmployeesFromADByID(userModel.EmployeeID);
                if (model != null)
                {
                    userModel.UserName = model.UserName;
                    string updateUsername = JsonHelper<UserModel>.Serialize(userModel);

                    _userAppService.Update(updateUsername);
                }
                else
                {
                    SetFalseTempData("Username has been set to EmployeeID due to AD failure");
                    return RedirectToAction("Index");
                }

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.ServerIsBusy);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult CreateOS()
        {
            ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitleOS(_jobTitleAppService);
            ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
            ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
            ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);
            ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService, true);

            UserModel usermodel = new UserModel();
            usermodel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

            return PartialView(usermodel);
        }

        // POST: User/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult CreateOS(UserModel userModel)
        {
            try
            {
                ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
                ViewBag.CountryList = DropDownHelper.BuildEmptyList();
                ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
                ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
                ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
                ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
                ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
                ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);
                ViewBag.CanteenList = DropDownHelper.BuildEmptyList();

                userModel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidData);
                    return RedirectToAction("Index");
                }

                string exist = _userAppService.GetBy("EmployeeID", userModel.EmployeeID);
                if (!string.IsNullOrEmpty(exist))
                {
                    SetFalseTempData(string.Format(UIResources.EmployeeIDExist, userModel.EmployeeID));
                    return RedirectToAction("Index");
                }

                exist = _userAppService.GetBy("Username", userModel.UserName);
                if (!string.IsNullOrEmpty(exist))
                {
                    SetFalseTempData(string.Format(UIResources.UsernameExist, userModel.UserName));
                    return RedirectToAction("Index");
                }

                // create job title
                if (userModel.JobTitleID == 0)
                {
                    JobTitleModel jt = new JobTitleModel();
                    jt.RoleName = null;
                    jt.Code = "OS";
                    jt.Title = "OS - " + userModel.JobTitle;

                    string newTitle = JsonHelper<JobTitleModel>.Serialize(jt);

                    long jobTitleID = _jobTitleAppService.Add(newTitle);

                    userModel.JobTitleID = jobTitleID;
                    userModel.JobTitle = jt.Title;
                }
                else
                {
                    string jt = _jobTitleAppService.GetById(userModel.JobTitleID);
                    JobTitleModel jtModel = jt.DeserializeToJobTitle();
                    userModel.JobTitle = jtModel.Title;
                }

                userModel.ModifiedBy = AccountName;
                userModel.ModifiedDate = DateTime.Now;
                userModel.IsActive = true;

                if (userModel.SubDepartmentID != 0)
                {
                    userModel.LocationID = userModel.SubDepartmentID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.SubDepartmentID);
                }
                else if (userModel.DepartmentID != 0)
                {
                    userModel.LocationID = userModel.DepartmentID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.DepartmentID);
                }
                else
                {
                    userModel.LocationID = userModel.ProdCenterID;
                    userModel.Location = _locationAppService.GetLocationFullCode(userModel.ProdCenterID);
                }

                userModel.IsOS = true;
                userModel.IsFast = true;
                string[] spvNames = userModel.SupervisorName.Split('-');
                userModel.SupervisorName = spvNames.Length > 1 ? spvNames[1].TrimStart() : "";
                userModel.SupervisorID = spvNames.Length > 1 ? spvNames[0].TrimStart() : "";

                _userAppService.AddModel(userModel);

                // create emp profiles
                EmployeeModel empModel = new EmployeeModel();
                empModel.EmployeeID = userModel.EmployeeID;
                empModel.PositionDesc = userModel.JobTitle;
                empModel.Status = "Active";
                empModel.OS = "1";
                empModel.FullName = userModel.FullName;
                empModel.GroupType = userModel.GroupType;
                empModel.GroupName = userModel.GroupType.Trim() == "NS" ? "" : userModel.GroupName;

                if (userModel.SupervisorID != null && userModel.SupervisorID != "")
                    empModel.ReportToID1 = userModel.SupervisorID;


                string newEmp = JsonHelper<EmployeeModel>.Serialize(empModel);

                _empService.Add(newEmp);

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.CreateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Details(long id)
        {
            UserModel model = _userAppService.GetModelById(id);

            return PartialView(model);
        }

        [HttpPost]
        public ActionResult GetExtraRoles(long userId)
        {
            var draw = Request.Form.GetValues("draw").FirstOrDefault();

            string userRoles = _userRoleAppService.FindBy("UserID", userId);
            List<UserRoleModel> userRoleList = userRoles.DeserializeToUserRoleList();

            List<ExtraRoleModel> data = new List<ExtraRoleModel>();

            foreach (var item in userRoleList)
            {
                data.Add(new ExtraRoleModel { Role = item.RoleName, ModifiedBy = item.ModifiedBy, ModifiedDate = item.ModifiedDate.HasValue ? item.ModifiedDate.Value : DateTime.Now });
            }

            // total number of rows count     
            int recordsFiltered = data.Count();

            int recordsTotal = data.Count();

            return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
        }

        // GET: User/Edit/5
        public ActionResult Edit(long id)
        {
            ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
            ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
            ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);

            UserModel model = _userAppService.GetModelById(id);
            model.Access = new AccessRightDBModel();
            model.Access.IsAdmin = AccountIsAdmin;

            if (model.IsOS)
            {
                ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitleOS(_jobTitleAppService, false);
                ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService, false);
            }
            else
            {
                ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService, true);
                ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitle(_jobTitleAppService);
            }
            EmployeeModel emp = _empService.GetModelByEmpId(model.EmployeeID);
            model.GroupType = emp.GroupType == null ? string.Empty : emp.GroupType.Trim();
            model.GroupName = emp.GroupName == null ? string.Empty : emp.GroupName.Trim();

            long countryID = 0;
            long pcID = 0;
            long depID = 0;
            long subDepID = 0;
            string completeLocation = DropDownHelper.ExtractLocation(_locationAppService, model.LocationID, out countryID, out pcID, out depID, out subDepID);
            model.CountryID = countryID;
            model.ProdCenterID = pcID;
            model.DepartmentID = depID;
            model.SubDepartmentID = subDepID;

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);

            if (model.CountryID == 0)
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            else
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, model.CountryID);

            if (model.ProdCenterID == 0)
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            else
                ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, model.ProdCenterID);

            if (model.DepartmentID == 0)
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            else
                ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, model.DepartmentID);

            return PartialView(model);
        }

        // POST: User/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(UserModel userModel)
        {
            try
            {
                ViewBag.EmployeeList = DropDownHelper.BuildEmptyList();
                ViewBag.CountryList = DropDownHelper.BuildEmptyList();
                ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
                ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
                ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
                ViewBag.UserAdminList = DropDownHelper.BindDropDownUserAdmin();
                ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
                ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);
                ViewBag.CanteenList = DropDownHelper.BuildEmptyList();

                userModel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidData);
                    return RedirectToAction("Index");
                }

                if (userModel.ProdCenterID == 0)
                {
                    SetFalseTempData(UIResources.ProdCenterIsMissing);
                    return RedirectToAction("Index");
                }

                UserModel oldUserModel = _userAppService.GetModelById(userModel.ID);
                oldUserModel.IsActive = userModel.IsActive;
                oldUserModel.IsAdmin = userModel.IsAdmin;
                oldUserModel.UserName = userModel.UserName;
                oldUserModel.EmployeeID = userModel.EmployeeID;
                if (oldUserModel.IsOS) //OS tidak perlu canteen
                    oldUserModel.CanteenID = userModel.CanteenID;
                else
                    oldUserModel.CanteenID = userModel.CanteenID;
                oldUserModel.ModifiedBy = AccountName;
                oldUserModel.ModifiedDate = DateTime.Now;

                bool isSpvUpdated = false;

                if (oldUserModel.SupervisorID != userModel.SupervisorID)
                {
                    isSpvUpdated = true;
                    oldUserModel.SupervisorID = userModel.SupervisorID;
                    string[] spvNames = userModel.SupervisorName.Split('-');
                    oldUserModel.SupervisorName = spvNames.Length > 1 ? spvNames[1].TrimStart() : "";
                }

                if (userModel.SubDepartmentID != 0)
                {
                    oldUserModel.LocationID = userModel.SubDepartmentID;
                    oldUserModel.Location = _locationAppService.GetLocationFullCode(userModel.SubDepartmentID);
                }
                else if (userModel.DepartmentID != 0)
                {
                    oldUserModel.LocationID = userModel.DepartmentID;
                    oldUserModel.Location = _locationAppService.GetLocationFullCode(userModel.DepartmentID);
                }
                else
                {
                    oldUserModel.LocationID = userModel.ProdCenterID;
                    oldUserModel.Location = _locationAppService.GetLocationFullCode(userModel.ProdCenterID);
                }

                bool isUpdateJobTitle = false;
                if (userModel.JobTitleID != 0 && userModel.JobTitleID != oldUserModel.JobTitleID)
                {
                    oldUserModel.JobTitleID = userModel.JobTitleID;
                    string jt = _jobTitleAppService.GetById(userModel.JobTitleID);
                    JobTitleModel jtModel = jt.DeserializeToJobTitle();
                    userModel.JobTitle = jtModel.Title;
                    isUpdateJobTitle = true;
                }

                _userAppService.UpdateModel(oldUserModel);

                //update group type
                string emp = _empService.FindByNoTracking("EmployeeID", userModel.EmployeeID);
                if (!string.IsNullOrEmpty(emp))
                {
                    EmployeeModel empModel = emp.DeserializeToEmployeeList().First();
                    empModel.Status = userModel.IsActive ? "Active" : "Inactive";
                    empModel.GroupType = userModel.GroupType;
                    empModel.GroupName = userModel.GroupName;
                    empModel.ReportToID1 = isSpvUpdated ? userModel.SupervisorID : empModel.ReportToID1;
                    if (userModel.SupervisorID != null && userModel.SupervisorID != "")
                        empModel.ReportToID1 = userModel.SupervisorID;
                    if (isUpdateJobTitle)
                        empModel.PositionDesc = userModel.JobTitle;

                    string updatedEmp = JsonHelper<EmployeeModel>.Serialize(empModel);

                    _empService.Update(updatedEmp);
                }

                SetTrueTempData(UIResources.UpdateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UpdateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        // POST: User/Delete/5
        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                ViewBag.Result = string.Empty;

                _userAppService.Remove(id);

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
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsFast", "true"));

                string userList = _userAppService.Find(filters);
                List<UserModel> users = userList.DeserializeToUserList();

                string userRoles = _userRoleAppService.GetAll(true);
                List<UserRoleModel> userExtraRoleList = userRoles.DeserializeToUserRoleList();

                List<ReferenceDetailModel> canteenModelList = new List<ReferenceDetailModel>();
                Dictionary<long, string> canteenMap = GetCanteenList(ref canteenModelList);

                int recordsTotal = users.Count();
                bool isDataComplete = false;
                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                if (!string.IsNullOrEmpty(searchValue) || (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDir)))
                {
                    isDataComplete = true;

                    // Construct custom attributes
                    string jobTitles = _jobTitleAppService.GetAll(true);
                    List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

                    string employees = _empService.GetAll(true);
                    List<EmployeeModel> employeeList = employees.DeserializeToEmployeeList();

                    // Construct custom attributes
                    foreach (var item in users)
                    {
                        JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                        item.RoleName = jt == null ? string.Empty : jt.RoleName;

                        if (!string.IsNullOrEmpty(item.EmployeeID))
                        {
                            EmployeeModel emp = employeeList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                            item.Employee = emp == null ? item.Employee : emp;
                        }

                        if (userExtraRoleList.Any(x => x.UserID == item.ID))
                        {
                            item.IsHasExtraRole = true;
                        }

                        if (item.CanteenID.HasValue)
                        {
                            item.Canteen = GetCanteen(item.CanteenID, canteenMap);
                        }
                    }
                }

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    users = users.Where(m => m.UserName != null && m.UserName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.RoleName != null && m.RoleName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorID != null && m.SupervisorID.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorName != null && m.SupervisorName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.EmployeeID != null && m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.FullName != null && m.Employee.FullName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.PositionDesc != null && m.Employee.PositionDesc.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.DepartmentDesc != null && m.Employee.DepartmentDesc.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.GroupType != null && m.Employee.GroupType.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.GroupName != null && m.Employee.GroupName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.BaseTownLocation != null && m.Employee.BaseTownLocation.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.CostCenter != null && m.Employee.CostCenter.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorID != null && m.SupervisorID.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorName != null && m.SupervisorName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.EmployeeType != null && m.Employee.EmployeeType.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    switch (sortColumn.ToLower())
                    {
                        case "username":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.UserName).ToList() : users.OrderByDescending(x => x.UserName).ToList();
                            break;
                        case "employeeid":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.EmployeeID).ToList() : users.OrderByDescending(x => x.UserName).ToList();
                            break;
                        case "fullname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.FullName).ToList() : users.OrderByDescending(x => x.Employee.FullName).ToList();
                            break;
                        case "rolename":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.RoleName).ToList() : users.OrderByDescending(x => x.RoleName).ToList();
                            break;
                        case "jobtitle":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.JobTitle).ToList() : users.OrderByDescending(x => x.JobTitle).ToList();
                            break;
                        case "positiondesc":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.PositionDesc).ToList() : users.OrderByDescending(x => x.Employee.PositionDesc).ToList();
                            break;
                        case "departmentdesc":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.DepartmentDesc).ToList() : users.OrderByDescending(x => x.Employee.DepartmentDesc).ToList();
                            break;
                        case "grouptype":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.GroupType).ToList() : users.OrderByDescending(x => x.Employee.GroupType).ToList();
                            break;
                        case "groupname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.GroupName).ToList() : users.OrderByDescending(x => x.Employee.GroupName).ToList();
                            break;
                        case "employeetype":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.EmployeeType).ToList() : users.OrderByDescending(x => x.Employee.EmployeeType).ToList();
                            break;
                        case "basetownlocation":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.BaseTownLocation).ToList() : users.OrderByDescending(x => x.Employee.BaseTownLocation).ToList();
                            break;
                        case "costcenter":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.CostCenter).ToList() : users.OrderByDescending(x => x.Employee.CostCenter).ToList();
                            break;
                        case "location":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Location).ToList() : users.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "supervisorid":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.SupervisorID).ToList() : users.OrderByDescending(x => x.SupervisorID).ToList();
                            break;
                        case "supervisorname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.SupervisorName).ToList() : users.OrderByDescending(x => x.SupervisorName).ToList();
                            break;
                        default:
                            break;
                    }
                }

                // total number of rows count     
                int recordsFiltered = users.Count();

                // Paging     
                var data = users.Skip(skip).Take(pageSize).ToList();

                if (!isDataComplete)
                {
                    string jobTitles = _jobTitleAppService.GetAll(true);
                    List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

                    string employees = _empService.GetAll(true);
                    List<EmployeeModel> employeeList = employees.DeserializeToEmployeeList();

                    // Construct custom attributes
                    foreach (var item in data)
                    {
                        JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                        item.RoleName = jt == null ? string.Empty : jt.RoleName;
                        item.JobTitle = jt == null ? string.Empty : jt.Title;

                        if (!string.IsNullOrEmpty(item.EmployeeID))
                        {
                            EmployeeModel emp = employeeList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                            if (emp != null)
                            {
                                item.Employee = emp == null ? item.Employee : emp;
                            }
                        }

                        if (userExtraRoleList.Any(x => x.UserID == item.ID))
                        {
                            item.IsHasExtraRole = true;
                        }

                        if (item.CanteenID.HasValue)
                        {
                            item.Canteen = GetCanteen(item.CanteenID, canteenMap);
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

                return Json(new { data = new List<UserModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllByLocation()
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
                filters.Add(new QueryFilter("IsFast", "true"));

                string userList = _userAppService.Find(filters);
                List<UserModel> users = userList.DeserializeToUserList();
                // Modified
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
                users = users.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

                string userRoles = _userRoleAppService.GetAll(true);
                List<UserRoleModel> userExtraRoleList = userRoles.DeserializeToUserRoleList();

                List<ReferenceDetailModel> canteenModelList = new List<ReferenceDetailModel>();
                Dictionary<long, string> canteenMap = GetCanteenList(ref canteenModelList);

                int recordsTotal = users.Count();
                bool isDataComplete = false;
                sortColumn = sortColumn == "ID" ? "" : sortColumn;

                if (!string.IsNullOrEmpty(searchValue) || (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDir)))
                {
                    isDataComplete = true;

                    // Construct custom attributes
                    string jobTitles = _jobTitleAppService.GetAll(true);
                    List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

                    string employees = _empService.GetAll(true);
                    List<EmployeeModel> employeeList = employees.DeserializeToEmployeeList();

                    // Construct custom attributes
                    foreach (var item in users)
                    {
                        JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                        item.RoleName = jt == null ? string.Empty : jt.RoleName;

                        if (!string.IsNullOrEmpty(item.EmployeeID))
                        {
                            EmployeeModel emp = employeeList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                            item.Employee = emp == null ? item.Employee : emp;
                        }

                        if (userExtraRoleList.Any(x => x.UserID == item.ID))
                        {
                            item.IsHasExtraRole = true;
                        }

                        if (item.CanteenID.HasValue)
                        {
                            item.Canteen = GetCanteen(item.CanteenID, canteenMap);
                        }
                    }
                }

                // Search    
                if (!string.IsNullOrEmpty(searchValue))
                {
                    users = users.Where(m => m.UserName != null && m.UserName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.RoleName != null && m.RoleName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorID != null && m.SupervisorID.ToLower().Contains(searchValue.ToLower()) ||
                                             m.SupervisorName != null && m.SupervisorName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.EmployeeID != null && m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.FullName != null && m.Employee.FullName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.PositionDesc != null && m.Employee.PositionDesc.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.DepartmentDesc != null && m.Employee.DepartmentDesc.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.GroupType != null && m.Employee.GroupType.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.GroupName != null && m.Employee.GroupName.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.BaseTownLocation != null && m.Employee.BaseTownLocation.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.CostCenter != null && m.Employee.CostCenter.ToLower().Contains(searchValue.ToLower()) ||
                                             m.Employee.EmployeeType != null && m.Employee.EmployeeType.ToLower().Contains(searchValue.ToLower())).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    switch (sortColumn.ToLower())
                    {
                        case "username":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.UserName).ToList() : users.OrderByDescending(x => x.UserName).ToList();
                            break;
                        case "employeeid":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.EmployeeID).ToList() : users.OrderByDescending(x => x.UserName).ToList();
                            break;
                        case "fullname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.FullName).ToList() : users.OrderByDescending(x => x.Employee.FullName).ToList();
                            break;
                        case "rolename":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.RoleName).ToList() : users.OrderByDescending(x => x.RoleName).ToList();
                            break;
                        case "jobtitle":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.JobTitle).ToList() : users.OrderByDescending(x => x.JobTitle).ToList();
                            break;
                        case "positiondesc":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.PositionDesc).ToList() : users.OrderByDescending(x => x.Employee.PositionDesc).ToList();
                            break;
                        case "departmentdesc":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.DepartmentDesc).ToList() : users.OrderByDescending(x => x.Employee.DepartmentDesc).ToList();
                            break;
                        case "grouptype":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.GroupType).ToList() : users.OrderByDescending(x => x.Employee.GroupType).ToList();
                            break;
                        case "groupname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.GroupName).ToList() : users.OrderByDescending(x => x.Employee.GroupName).ToList();
                            break;
                        case "employeetype":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.EmployeeType).ToList() : users.OrderByDescending(x => x.Employee.EmployeeType).ToList();
                            break;
                        case "basetownlocation":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.BaseTownLocation).ToList() : users.OrderByDescending(x => x.Employee.BaseTownLocation).ToList();
                            break;
                        case "costcenter":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Employee.CostCenter).ToList() : users.OrderByDescending(x => x.Employee.CostCenter).ToList();
                            break;
                        case "location":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.Location).ToList() : users.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "supervisorid":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.SupervisorID).ToList() : users.OrderByDescending(x => x.SupervisorID).ToList();
                            break;
                        case "supervisorname":
                            users = sortColumnDir == "asc" ? users.OrderBy(x => x.SupervisorName).ToList() : users.OrderByDescending(x => x.SupervisorName).ToList();
                            break;
                        default:
                            break;
                    }
                }

                // total number of rows count     
                int recordsFiltered = users.Count();

                // Paging     
                var data = users.Skip(skip).Take(pageSize).ToList();

                if (!isDataComplete)
                {
                    string jobTitles = _jobTitleAppService.GetAll(true);
                    List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

                    string employees = _empService.GetAll(true);
                    List<EmployeeModel> employeeList = employees.DeserializeToEmployeeList();

                    // Construct custom attributes
                    foreach (var item in data)
                    {
                        JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                        item.RoleName = jt == null ? string.Empty : jt.RoleName;
                        item.JobTitle = jt == null ? string.Empty : jt.Title;

                        if (!string.IsNullOrEmpty(item.EmployeeID))
                        {
                            EmployeeModel emp = employeeList.Where(x => x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
                            if (emp != null)
                            {
                                item.Employee = emp == null ? item.Employee : emp;
                            }
                        }

                        if (userExtraRoleList.Any(x => x.UserID == item.ID))
                        {
                            item.IsHasExtraRole = true;
                        }

                        if (item.CanteenID.HasValue)
                        {
                            item.Canteen = GetCanteen(item.CanteenID, canteenMap);
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

                return Json(new { data = new List<UserModel>() }, JsonRequestBehavior.AllowGet);
            }
        }


        private string GetRoleName(long jobTitleID)
        {
            string jobTitle = _jobTitleAppService.GetById(jobTitleID);
            JobTitleModel model = jobTitle.DeserializeToJobTitle();

            return model.RoleName;
        }

        public ActionResult UpdateGroupType()
        {
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();

            UserModel usermodel = new UserModel();
            usermodel.Access = GetAccess(WebConstants.MenuSlug.USER, _menuService);

            return PartialView(usermodel);
        }

        [HttpPost]
        public ActionResult UpdateGroupType(MppModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidData);
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

                    List<EmployeeModel> empList = new List<EmployeeModel>();
                    string jobtitle = string.Empty;


                    string groupType = _referenceAppService.GetDetailAll(ReferenceEnum.Group, true);
                    List<ReferenceDetailModel> groupTypeList = groupType.DeserializeToRefDetailList();

                    ReferenceModel gnModel = _referenceAppService.GetBy("Name", "GroupName").DeserializeToReference();
                    List<ReferenceDetailModel> groupNameList = _referenceAppService.FindDetailBy("ReferenceID", gnModel.ID).DeserializeToRefDetailList();

                    for (int index = 1; index < dt_.Rows.Count; index++)
                    {
                        string empId = dt_.Rows[index][0].ToString();
                        string gt = dt_.Rows[index][2].ToString();
                        string gn = dt_.Rows[index][3].ToString();

                        if (empId == string.Empty && gt == string.Empty)
                        {
                            break;
                        }

                        if (groupTypeList.Any(x => x.Code == gt.Trim()))
                        {
                            if (groupNameList.Any(x => x.Code == gn.Trim()))
                            {
                                EmployeeModel empModel = new EmployeeModel();
                                empModel.EmployeeID = dt_.Rows[index][0].ToString();
                                empModel.GroupType = gt;
                                empModel.GroupName = gn;
                                empModel.Location = dt_.Rows[index][4].ToString();
                                empModel.Canteen = dt_.Rows[index][5].ToString();

                                empList.Add(empModel);
                            }
                            else
                            {
                                SetFalseTempData(string.Format(UIResources.GroupNameInvalidAt, gn, index + 1));
                                return RedirectToAction("Index");
                            }
                        }
                        else
                        {
                            SetFalseTempData(string.Format(UIResources.GroupTypeInvalidAt, gt, index + 1));
                            return RedirectToAction("Index");
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    if (empList.Count > 0)
                    {
                        long locationID = AccountLocationID;
                        string location = locationID == 0 ? "" : _locationAppService.GetLocationFullCode(AccountLocationID);

                        Dictionary<string, long> canteenMap = GetCanteenList();

                        string emps = _empService.GetAll();
                        List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("IsFast", "true"));

                        string userList = _userAppService.Find(filters);
                        List<UserModel> usersModelList = userList.DeserializeToUserList();

                        List<EmployeeModel> empUpdateList = new List<EmployeeModel>();
                        List<UserModel> userUpdateList = new List<UserModel>();

                        foreach (var item in empList.ToList())
                        {
                            EmployeeModel empModel = empModelList.Where(x => x.EmployeeID == item.EmployeeID).FirstOrDefault();
                            if (empModel != null && (empModel.GroupType != item.GroupType || empModel.GroupName != item.GroupName))
                            {
                                empModel.GroupType = item.GroupType;
                                empModel.GroupName = item.GroupName;
                                empModel.ModifiedBy = AccountName;
                                empModel.ModifiedDate = DateTime.Now;

                                //tambahan fery location
                                empModel.Location = item.Location;

                                empUpdateList.Add(empModel);

                                //string updatedEmp = JsonHelper<EmployeeModel>.Serialize(empModel);
                                //_empService.Update(updatedEmp);
                            }

                            long newCanteenID = 0;
                            if (!string.IsNullOrEmpty(item.Canteen))
                            {
                                if (!canteenMap.TryGetValue(item.Canteen, out newCanteenID))
                                {
                                    SetFalseTempData("Canteen " + item.Canteen + " is invalid");
                                    return RedirectToAction("Index");
                                }
                            }

                            UserModel userModel = usersModelList.Where(x => x.EmployeeID == item.EmployeeID).FirstOrDefault();
                            if (userModel != null &&
                                (userModel.LocationID != locationID && userModel.Location != location) || (userModel.CanteenID != newCanteenID))
                            {
                                if (!userModel.LocationID.HasValue || userModel.LocationID.Value == 0)
                                {
                                    userModel.LocationID = locationID;
                                    userModel.Location = location;
                                }

                                //tambahan fery location baca dari excel nya
                                if (item.Location != "")
                                {
                                    userModel.Location = item.Location;
                                    long newloc = 0;
                                    newloc = _locationAppService.GetLocationID(item.Location);
                                    if (newloc > 0)
                                        userModel.LocationID = newloc;
                                }

                                userModel.CanteenID = newCanteenID;
                                userModel.ModifiedBy = AccountName;
                                userModel.ModifiedDate = DateTime.Now;

                                userUpdateList.Add(userModel);

                                //string updatedUser = JsonHelper<UserModel>.Serialize(userModel);
                                //_userAppService.Update(updatedUser);
                            }
                        }

                        if (UpdateUserAndEmployee(userUpdateList, empUpdateList))
                        {
                            SetTrueTempData(UIResources.UploadSucceed);
                        }
                        else
                        {
                            SetFalseTempData("Upload failed. Please try it again");
                        }
                    }
                    else
                    {
                        SetFalseTempData(UIResources.InvalidData);
                    }
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
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

        private bool UpdateUserAndEmployee(List<UserModel> userUpdateList, List<EmployeeModel> empUpdateList)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            string updateEmp = "UPDATE EmployeeProfiles SET GroupName = @GroupName, GroupType = @GroupType WHERE EmployeeID = @EmployeeID";
            string updateUser = "UPDATE Users SET Location = @Location, LocationID = @LocationID, CanteenID = @CanteenID WHERE EmployeeID = @EmployeeID";

            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    foreach (var emp in empUpdateList)
                    {
                        SqlCommand command = new SqlCommand(updateEmp, connection, transaction);
                        command.Parameters.Add("@GroupName", SqlDbType.Char).Value = emp.GroupName;
                        command.Parameters.Add("@GroupType", SqlDbType.Char).Value = emp.GroupType;
                        command.Parameters.Add("@EmployeeID", SqlDbType.Char).Value = emp.EmployeeID;
                        command.ExecuteNonQuery();
                    }

                    foreach (var user in userUpdateList)
                    {
                        SqlCommand command = new SqlCommand(updateUser, connection, transaction);
                        command.Parameters.Add("@Location", SqlDbType.VarChar).Value = user.Location;
                        command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = user.LocationID;
                        command.Parameters.Add("@EmployeeID", SqlDbType.Char).Value = user.EmployeeID;
                        command.Parameters.Add("@CanteenID", SqlDbType.BigInt).Value = user.CanteenID;
                        command.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();

                    return false;
                }
            }

            return true;
        }

        public ActionResult DownloadTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateUserGroupType.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateUserGroupType.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
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
    }
}
