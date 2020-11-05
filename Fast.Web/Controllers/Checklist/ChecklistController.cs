using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Utils;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Data.SqlClient;
using System.DirectoryServices;
using Fast.Web.Resources;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using System.Web;
using System.Drawing;
using System.Threading.Tasks;
using System.Globalization;
using System.Configuration;

namespace Fast.Web.Controllers.Checklist
{
    public class ChecklistController : BaseController<AccessRightDBModel>
    {
        private static string HEADER_COLOR = "#99ccff";
        private readonly IChecklistAppService _checklistAppService;
        private readonly IChecklistLocationAppService _checklistLocationAppService;
        private readonly IChecklistComponentAppService _checklistComponentAppService;
        private readonly IChecklistSubmitAppService _checklistSubmitAppService;
        private readonly IChecklistValueAppService _checklistValueAppService;
        private readonly IChecklistValueHistoryAppService _checklistValueHistoryAppService;
        private readonly IChecklistApprovalAppService _checklistApprovalAppService;
        private readonly IChecklistApproverAppService _checklistApproverAppService;
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly IRoleAppService _roleAppService;
        private readonly IJobTitleAppService _jobTitleAppService;
        private readonly IUserAppService _userAppService;

        public ChecklistController(
            IChecklistAppService checklistAppService,
            IChecklistLocationAppService checklistLocationAppService,
            IChecklistComponentAppService checklistComponentAppService,
            IChecklistSubmitAppService checklistSubmitAppService,
            IChecklistValueAppService checklistValueAppService,
            IChecklistValueHistoryAppService checklistValueHistoryAppService,
            IChecklistApprovalAppService checklistApprovalAppService,
            IChecklistApproverAppService checklistApproverAppService,
            ILoggerAppService logger,
            IEmployeeAppService employeeAppService,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IReferenceDetailAppService referenceDetailAppService,
            IRoleAppService roleAppService,
            IJobTitleAppService jobTitleAppService,
            IUserAppService userAppService)
        {
            _checklistAppService = checklistAppService;
            _checklistLocationAppService = checklistLocationAppService;
            _checklistComponentAppService = checklistComponentAppService;
            _checklistSubmitAppService = checklistSubmitAppService;
            _checklistValueAppService = checklistValueAppService;
            _checklistValueHistoryAppService = checklistValueHistoryAppService;
            _checklistApprovalAppService = checklistApprovalAppService;
            _checklistApproverAppService = checklistApproverAppService;
            _logger = logger;
            _referenceAppService = referenceAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _locationAppService = locationAppService;
            _employeeAppService = employeeAppService;
            _roleAppService = roleAppService;
            _jobTitleAppService = jobTitleAppService;
            _userAppService = userAppService;
        }

        private List<SelectListItem> BindDropDownLocation()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            LocationTreeModel model = new LocationTreeModel();

            int index = 1;

            string currentCountry = "ID";
            string pcs = _locationAppService.FindBy("ParentCode", currentCountry, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            if (pcList != null)
            {
                foreach (var item in pcList)
                {
                    string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
                    ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
                    model.ProductionCenters.Add(pcModel);
                }

                if (model.ProductionCenters != null)
                {
                    foreach (var pc in model.ProductionCenters)
                    {
                        LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
                        string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
                        List<LocationModel> departmentList = departments.DeserializeToLocationList();

                        if (departmentList != null)
                        {
                            foreach (var d in departmentList)
                            {
                                string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
                                DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

                                string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                                List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                                if (subdepartmentList != null)
                                {
                                    foreach (var subdeb in subdepartmentList)
                                    {
                                        string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
                                        deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
                                    }

                                    pc.Departments.Add(deptModel);
                                }
                            }
                        }
                    }

                    if (model.ProductionCenters != null)
                    {
                        foreach (var pc in model.ProductionCenters)
                        {
                            if (pc.Departments != null)
                            {
                                foreach (var d in pc.Departments)
                                {
                                    if (d.SubDepartments != null)
                                    {
                                        foreach (var sd in d.SubDepartments)
                                        {
                                            _menuList.Add(new SelectListItem
                                            {
                                                Text = pc.Description + " - " + d.Description + " " + sd.Description,
                                                Value = sd.LocationID.ToString()
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownEmployee()
        {
            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (employeeList != null)
            {
                employeeList = employeeList.OrderBy(x => x.FullName).Distinct().ToList();

                _menuList.Add(new SelectListItem
                {
                    Text = "Select employee (optional)",
                    Value = ""
                });

                if (employeeList != null)
                {
                    foreach (var emplo in employeeList)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = emplo.EmployeeID + " " + emplo.FullName,
                            Value = emplo.EmployeeID
                        });
                    }
                }
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownReference()
        {
            string reference = _referenceAppService.GetAll(true);
            List<ReferenceModel> referenceList = reference.DeserializeToReferenceList();
            List<SelectListItem> _menuList = new List<SelectListItem>();
            if (referenceList != null)
            {
                referenceList = referenceList.OrderBy(x => x.Purpose).ToList();

                if (referenceList != null)
                {
                    foreach (var refe in referenceList)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = refe.Purpose,
                            Value = refe.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownRole()
        {
            string role = _roleAppService.GetAll(true);
            List<RoleModel> roleList = role.DeserializeToRoleList();
            List<SelectListItem> _menuList = new List<SelectListItem>();
            if (roleList != null)
            {
                roleList = roleList.OrderBy(x => x.Description).ToList();

                _menuList.Add(new SelectListItem
                {
                    Text = "Select Role",
                    Value = ""
                });

                if (roleList != null)
                {
                    foreach (var emplo in roleList)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = emplo.Description,
                            Value = emplo.Name
                        });
                    }
                }
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownJobTitle()
        {
            string jobTitle = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (jobTitleList != null)
            {
                jobTitleList = jobTitleList.GroupBy(x => x.Title).Select(y => y.First()).ToList();
                jobTitleList = jobTitleList.OrderBy(x => x.Title).ToList();

                _menuList.Add(new SelectListItem
                {
                    Text = "Select Job Title",
                    Value = ""
                });
                _menuList.Add(new SelectListItem
                {
                    Text = "Default First Approver",
                    Value = "Default First Approver"
                });
                _menuList.Add(new SelectListItem
                {
                    Text = "Default Second Approver",
                    Value = "Default Second Approver"
                });

                if (jobTitleList != null)
                {
                    foreach (var emplo in jobTitleList)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = emplo.Title,
                            Value = emplo.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        // GET: Checklist
        [CustomAuthorize("checklist")]
        public ActionResult Index(string data = "", string location1 = "", string location2 = "", string location3 = "")
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string checklistLocation = "";
            if (data == "GetAll")
            {
            }
            else if (data == "Location")
            {
                long locat = 0;
                if (location3 != "" && location3 != "null")
                {
                    locat = Int64.Parse(location3);
                }
                else if (location2 != "" && location2 != "null")
                {
                    locat = Int64.Parse(location2);

                }
                else if (location1 != "" && location1 != "null")
                {
                    locat = Int64.Parse(location1);
                }

                if (locat > 0)
                {
                    checklistLocation = _checklistLocationAppService.FindBy("LocationID", locat, true);
                }
                else
                {
                    data = "GetAll";
                }
            }
            else
            {
                checklistLocation = _checklistLocationAppService.FindBy("LocationID", (long)Account.LocationID, true);
            }

            var CLmodel = checklistLocation.DeserializeToChecklistLocationList();

            List<long> CheckIDs = CLmodel.Select(x => x.ChecklistID).Distinct().ToList();

            string checklists = _checklistAppService.GetAll(true);
            List<ChecklistModel> checklistsList = checklists.DeserializeToChecklistList();

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            if (data == "GetAll")
            {
                ViewBag.FilterMode = 1;
            }
            else if (data == "Location")
            {
                ViewBag.FilterMode = 2;
                checklistsList = checklistsList.Where(x => CheckIDs.Contains(x.ID)).OrderBy(x => x.MenuTitle).ThenBy(x => x.Header).ToList();
            }
            else
            {
                ViewBag.FilterMode = 0;
                checklistsList = checklistsList.Where(x => CheckIDs.Contains(x.ID)).OrderBy(x => x.MenuTitle).ThenBy(x => x.Header).ToList();
            }

            string emplo = _employeeAppService.GetAll(true);
            List<SimpleEmployeeModel> emploModel = JsonConvert.DeserializeObject<List<SimpleEmployeeModel>>(emplo);

            foreach (var check in checklistsList)
            {
                check.Creator = emploModel.Where(x => x.EmployeeID == check.CreatorEmployeeID).FirstOrDefault();
            }

            ViewBag.Checklists = checklistsList;
            ViewBag.isAdmin = AccountIsAdmin;
            ViewBag.MyEmpolyeeID = AccountEmployeeID;

            return View();
        }

        [CustomAuthorize("checklistcatalog")]
        public ActionResult Catalog(string data = "", string location1 = "", string location2 = "", string location3 = "")
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string checklistLocation = "";
            if (data == "GetAll")
            {
            }
            else if (data == "Location")
            {
                long locat = 0;
                if (location3 != "" && location3 != "null")
                {
                    locat = Int64.Parse(location3);
                }
                else if (location2 != "" && location2 != "null")
                {
                    locat = Int64.Parse(location2);

                }
                else if (location1 != "" && location1 != "null")
                {
                    locat = Int64.Parse(location1);
                }

                if (locat > 0)
                {
                    checklistLocation = _checklistLocationAppService.FindBy("LocationID", locat, true);
                }
                else
                {
                    data = "GetAll";
                }
            }
            else
            {
                checklistLocation = _checklistLocationAppService.FindBy("LocationID", (long)Account.LocationID, true);
            }

            var CLmodel = checklistLocation.DeserializeToChecklistLocationList();

            List<long> CheckIDs = CLmodel.Select(x => x.ChecklistID).Distinct().ToList();

            string checklists = _checklistAppService.GetAll(true);
            List<ChecklistModel> checklistsList = checklists.DeserializeToChecklistList();

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            if (data == "GetAll")
            {
                ViewBag.FilterMode = 1;
            }
            else if (data == "Location")
            {
                ViewBag.FilterMode = 2;
                checklistsList = checklistsList.Where(x => CheckIDs.Contains(x.ID)).OrderBy(x => x.MenuTitle).ThenBy(x => x.Header).ToList();
            }
            else
            {
                ViewBag.FilterMode = 0;
                checklistsList = checklistsList.Where(x => CheckIDs.Contains(x.ID)).OrderBy(x => x.MenuTitle).ThenBy(x => x.Header).ToList();
            }

            string emplo = _employeeAppService.GetAll(true);
            List<SimpleEmployeeModel> emploModel = JsonConvert.DeserializeObject<List<SimpleEmployeeModel>>(emplo);

            foreach (var check in checklistsList)
            {
                check.Creator = emploModel.Where(x => x.EmployeeID == check.CreatorEmployeeID).FirstOrDefault();
            }

            ViewBag.Checklists = checklistsList;

            return View();
        }

        [CustomAuthorize("checklist")]
        public ActionResult Create()
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            ViewBag.LocationList = BindDropDownLocation();
            ViewBag.EmployeeList = BindDropDownEmployee();
            ViewBag.ReferenceList = BindDropDownReference();
            ViewBag.JobTitleList = BindDropDownJobTitle();

            string employee = _employeeAppService.GetBy("EmployeeID", AccountEmployeeID, true);
            EmployeeModel currentEmployee = employee.DeserializeToEmployee();

            Account.Employee = currentEmployee;

            ViewBag.Approver1 = Account.Employee.ReportToID1;
            ViewBag.Approver2 = Account.Employee.ReportToID2;


            /************************ JOB TITLE USER **********************/
            var dictionary = new Dictionary<string, List<SimpleEmployeeModel>>();


            string emplo = _employeeAppService.GetAll(true);
            List<SimpleEmployeeModel> emploModel = JsonConvert.DeserializeObject<List<SimpleEmployeeModel>>(emplo);
            //diri sendiri ndak boleh jadi approver
            //emploModel = emploModel.Where(x => x.EmployeeID != AccountEmployeeID).ToList();

            if (ViewBag.JobTitleList != null)
            {
                foreach (SelectListItem jobTitle in ViewBag.JobTitleList)
                {
                    var modelEmployee = new List<SimpleEmployeeModel>();

                    if (jobTitle.Value != "")
                    {
                        if (jobTitle.Value == "Default First Approver" || jobTitle.Value == "Default Second Approver")
                        {
                            SimpleEmployeeModel temp_employ = new SimpleEmployeeModel();
                            temp_employ.EmployeeID = "0";
                            temp_employ.FullName = "According to Approver from Peoplesoft";
                            modelEmployee.Add(temp_employ);
                        }
                        else
                        {
                            string user = _userAppService.FindBy("JobTitleID", jobTitle.Value, true);
                            if (!string.IsNullOrEmpty(user))
                            {
                                List<UserModel> userModel = user.DeserializeToUserList();
                                if (userModel != null)
                                {
                                    foreach (var usr in userModel)
                                    {
                                        SimpleEmployeeModel temp_employ = emploModel.Where(x => x.EmployeeID.Trim() == usr.EmployeeID.Trim()).FirstOrDefault();
                                        if (temp_employ != null)
                                        {
                                            modelEmployee.Add(temp_employ);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //beberapa user kadang memakai employeeID yg sama
                    modelEmployee = modelEmployee.Distinct().ToList();
                    dictionary.Add(jobTitle.Value, modelEmployee);
                }
            }

            ViewBag.JobTitleUsers = dictionary;

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            var allADGroup = new List<string>();

            allADGroup.Add("Default First Approver");
            allADGroup.Add("Default Second Approver");

            /*
            DirectoryEntry de = new DirectoryEntry("LDAP://DC=PMINTL,DC=NET");
            de.Username = "s-idfastdev";
            de.Password = "A7q$3#4$132g3%G";
            DirectorySearcher ds = new DirectorySearcher(de);

            try
            {
                // setup directory searcher
                ds.PageSize = 1000;
                ds.SearchScope = SearchScope.Subtree;
                ds.CacheResults = false;

                ds.Filter = "(&(objectCategory=group))";
                ds.PropertiesToLoad.Add("name");
                ds.PropertiesToLoad.Add("description");

                SearchResultCollection results = ds.FindAll();

                foreach (SearchResult res in results)
                {
                    String name = ((res.Properties["name"])[0]).ToString();
                    //string groupDescription = (res.Properties["description"])[0].ToString();

                    allADGroup.Add(name);
                }
            }
            catch (Exception ex)
            {
                
            }
            */

            ViewBag.AllADGroups = allADGroup;

            return View();
        }

        [HttpPost]
        public ActionResult GetUsersInGroup(string id)
        {
            List<string> _menuList = new List<string>();

            /*
            DirectoryEntry de = new DirectoryEntry("LDAP://DC=PMINTL,DC=NET");
            de.Username = "s-idfastdev";
            de.Password = "A7q$3#4$132g3%G";
            DirectorySearcher ds = new DirectorySearcher(de);

            try
            {
                string query = "(&(objectCategory=person)(objectClass=user)(memberOf=*))";
                ds.Filter = query;
                ds.PropertiesToLoad.Add("memberOf");
                ds.PropertiesToLoad.Add("name");

                System.DirectoryServices.SearchResultCollection mySearchResultColl = ds.FindAll();
                foreach (SearchResult result in mySearchResultColl)
                {
                    foreach (string prop in result.Properties["memberOf"])
                    {
                        if (prop.Contains(id))
                        {
                            _menuList.Add(result.Properties["name"][0].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            */

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [CustomAuthorize("checklist")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(ChecklistModel model, List<string> approver)
        {
            var lala = Request;
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Create");
                }

                Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                if (model.IconFile != null && model.IconFile.ContentLength > 0)
                {
                    var tempFile = model.IconFile;
                    model.Icon = unixTimestamp.ToString() + "_" + Path.GetFileName(tempFile.FileName);
                    try
                    {
                        tempFile.SaveAs(Server.MapPath("~/Uploads/checklist/icon/") + model.Icon);
                    }
                    catch
                    {
                        // do nothing
                    }

                    model.IconFile = null;
                }

                model.CreatorEmployeeID = AccountEmployeeID;
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;


                /*************** SIMPAN LOCATION ********************/

                List<LocationModel> pcList = new List<LocationModel>();
                List<LocationModel> departmentList = new List<LocationModel>();

                if (model.Location1 == null)
                {
                    string currentCountry = "ID";
                    string pcs = _locationAppService.FindBy("ParentCode", currentCountry, true);
                    pcList = pcs.DeserializeToLocationList();
                }
                else
                {
                    foreach (var locat in model.Location1)
                    {
                        string pcs = _locationAppService.GetById(locat);
                        pcList.Add(pcs.DeserializeToLocation());
                    }
                }

                if (pcList != null)
                {
                    foreach (var item in pcList)
                    {
                        var modelLocation = new ChecklistLocationModel();
                        modelLocation.LocationID = item.ID;
                        model.Locations.Add(modelLocation);

                        if (model.Location2 == null)
                        {
                            string departments = _locationAppService.FindBy("ParentID", item.ID, true);
                            var dl = departments.DeserializeToLocationList();

                            if (dl != null)
                            {
                                foreach (var locat in dl)
                                {
                                    departmentList.Add(locat);
                                }
                            }
                        }
                        else
                        {
                            foreach (var locat in model.Location2)
                            {
                                string pcs = _locationAppService.GetById(locat);
                                departmentList.Add(pcs.DeserializeToLocation());
                            }
                        }
                    }

                    if (departmentList != null)
                    {
                        foreach (var d in departmentList)
                        {
                            var modelLocation = new ChecklistLocationModel();
                            modelLocation.LocationID = d.ID;
                            model.Locations.Add(modelLocation);

                            string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                            List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                            if (model.Location3 == null)
                            {
                                if (subdepartmentList != null)
                                {
                                    foreach (var subdeb in subdepartmentList)
                                    {
                                        modelLocation = new ChecklistLocationModel();
                                        modelLocation.LocationID = subdeb.ID;
                                        model.Locations.Add(modelLocation);
                                    }
                                }
                            }
                            else
                            {
                                foreach (var locat in model.Location3)
                                {
                                    modelLocation = new ChecklistLocationModel();
                                    modelLocation.LocationID = locat;
                                    model.Locations.Add(modelLocation);
                                }
                            }
                        }
                    }
                }

                //************ Save images name in array ***************

                var images = new List<string>();

                if (Request.Files != null && Request.Files.Count > 0)
                {
                    foreach (string fileName in Request.Files)
                    {
                        var uploadedFile = Request.Files[fileName] as HttpPostedFileBase;
                        var image = "";
                        if (uploadedFile.ContentLength > 0)
                        {
                            unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                            image = unixTimestamp.ToString() + "_" + Path.GetFileName(uploadedFile.FileName);
                            uploadedFile.SaveAs(Server.MapPath("~/Uploads/checklist/component/") + image);
                        }
                        images.Add(image);
                    }
                }


                string data = JsonHelper<ChecklistModel>.Serialize(model);
                var ID = _checklistAppService.Add(data);
                var image_counter = 0;

                if (model.Components != null)
                {
                    for (int i = 0; i < model.Components.Count(); i++)
                    {
                        var component = new ChecklistComponentModel();

                        /*
                        if (i < (model.Components.Count() - 3) && model.Components[i] == "11001100" && model.Components[i + 1] == "" && model.Components[i + 2] == "input_text" && model.Components[i + 3] == "True")
                        {
                            i++; i++; i++;
                            continue;
                        }
                        */

                        component.ChecklistID = ID;
                        component.ModifiedBy = AccountName;
                        component.ModifiedDate = DateTime.Now;
                        component.Segment = "header";
                        component.OrderNum = i;
                        if (model.Components[i] == "11001100")
                        {
                            component.ComponentType = "label";
                            component.ComponentName = model.Components[++i];
                        }
                        else if (model.Components[i] == "input_option" || model.Components[i] == "input_reference" || model.Components[i] == "label")
                        {
                            component.ComponentType = model.Components[i].Trim();
                            component.ComponentName = model.Components[++i].Trim();
                            component.IsRequired = bool.Parse(model.Components[++i]);
                        }
                        else if (model.Components[i] == "image")
                        {
                            component.ComponentType = model.Components[i];
                            component.ComponentName = images[image_counter++];
                            component.IsRequired = bool.Parse(model.Components[++i]);
                        }
                        else
                        {
                            component.ComponentType = model.Components[i].Trim();
                            component.IsRequired = bool.Parse(model.Components[++i]);
                        }

                        data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                        var lalala = _checklistComponentAppService.Add(data);
                    }
                }

                int aprover_count = 0;
                for (var i = 0; i < approver.Count(); i++)
                {
                    if (approver[i] == "")
                    {
                        i++; i++; i++; i++;
                        continue;
                    }

                    List<string> tempEmployee = new List<string>();
                    var aprv = new ChecklistApproverModel();

                    aprv.ChecklistID = ID;
                    aprv.ADGroup = approver[i];

                    tempEmployee.Add(approver[++i]);
                    while (true)
                    {
                        if (IsNumber(approver[i + 1]))
                        {
                            tempEmployee.Add(approver[++i]);
                        }
                        else
                        {
                            break;
                        }
                    }

                    aprv.Approve = approver[++i];
                    aprv.Revise = approver[++i];
                    aprv.Edit = approver[++i];
                    aprv.Reject = approver[++i];
                    aprv.Tier = aprover_count++;
                    aprv.ModifiedBy = AccountName;
                    aprv.ModifiedDate = DateTime.Now;

                    foreach (var emploID in tempEmployee)
                    {
                        aprv.EmployeeID = emploID;

                        data = JsonHelper<ChecklistApproverModel>.Serialize(aprv);
                        var lalala = _checklistApproverAppService.Add(data);
                    }
                }

                return RedirectToAction("Content/" + ID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                Session["ResultLog"] = "error_Failed to create checklist";

                return RedirectToAction("Create");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult Content(long ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string checklist = _checklistAppService.GetById(ID);

            if (string.IsNullOrEmpty(checklist))
                return RedirectToAction("Create");

            ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
            ViewBag.Checklist = currentChecklist;

            string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
            var componentModel = component.DeserializeToChecklistComponentList();
            componentModel = componentModel.Where(x => x.Segment.Trim() == "content").ToList();

            if (componentModel.Count() > 0)
                return RedirectToAction("Edit/" + ID);

            ViewBag.ReferenceList = BindDropDownReference();

            return View();
        }

        [CustomAuthorize("checklist")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Content(long ChecklistID, int ColumnContent, List<string> headers, List<string> width, List<List<string>> model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    _logger.LogError("model tidak valid; " + JsonConvert.SerializeObject(headers) + JsonConvert.SerializeObject(width) + JsonConvert.SerializeObject(model), AccountID, AccountName);
                    Session["ResultLog"] = "error_Failed to build cheklist content, param mismatch";
                    return RedirectToAction("Content/" + ChecklistID);
                }

                //************ Save images name in array ***************

                var images = new List<string>();

                if (Request.Files != null && Request.Files.Count > 0)
                {
                    foreach (string fileName in Request.Files)
                    {
                        var uploadedFile = Request.Files[fileName] as HttpPostedFileBase;
                        var image = "";
                        if (uploadedFile.ContentLength > 0)
                        {
                            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                            image = unixTimestamp.ToString() + "_" + Path.GetFileName(uploadedFile.FileName);
                            uploadedFile.SaveAs(Server.MapPath("~/Uploads/checklist/component/") + image);
                        }
                        images.Add(image);
                    }
                }

                //************ UPDATE CHECKLIST ColumnContent ***************

                string checklist = _checklistAppService.GetById(ChecklistID, true);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();

                currentChecklist.ColumnContent = ColumnContent;

                checklist = JsonHelper<ChecklistModel>.Serialize(currentChecklist);
                _checklistAppService.Update(checklist);


                //************ SAVE HEADER ***************
                if (headers == null)
                {
                    _logger.LogError("model header tidak valid; " + JsonConvert.SerializeObject(headers), AccountID, AccountName);
                    Session["ResultLog"] = "error_Failed to build cheklist content, header param mismatch";
                    return RedirectToAction("Content/" + ChecklistID);
                }
                else
                {
                    for (int i = 0; i < headers.Count(); i++)
                    {
                        var component = new ChecklistComponentModel();

                        component.ChecklistID = ChecklistID;
                        component.ModifiedBy = AccountName;
                        component.ModifiedDate = DateTime.Now;
                        component.Segment = "content";
                        component.OrderNum = i;

                        component.ComponentType = "header";
                        component.ComponentName = headers[i].Trim();
                        component.AdditionalValue = width[i];

                        var data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                        _checklistComponentAppService.Add(data);
                    }
                }

                //************ SAVE CONTENT ***************

                if (model == null)
                {
                    _logger.LogError("model content tidak valid; " + JsonConvert.SerializeObject(model), AccountID, AccountName);
                    Session["ResultLog"] = "error_Failed to build cheklist content, content param mismatch";
                    return RedirectToAction("Content/" + ChecklistID);
                }
                else
                {
                    int image_counter = 0;
                    for (int i = 0; i < model.Count(); i++)
                    {
                        for (int j = 0; j < model[i].Count(); j++)
                        {
                            var component = new ChecklistComponentModel();

                            component.ComponentType = model[i][j].Trim();
                            component.ChecklistID = ChecklistID;
                            component.ModifiedBy = AccountName;
                            component.ModifiedDate = DateTime.Now;
                            component.Segment = "content";
                            component.OrderNum = j;
                            component.ColumnNum = i + 1;
                            component.AdditionalValue = "";

                            if (component.ComponentType == "input_option" || component.ComponentType == "input_radio" || component.ComponentType == "input_barcode" || component.ComponentType == "label" || component.ComponentType == "input_reference")
                            {
                                component.ComponentName = model[i][++j].Trim();
                            }
                            else if (component.ComponentType == "image")
                            {
                                if (images != null && images[image_counter] != null)
                                {
                                    component.ComponentName = images[image_counter];
                                    image_counter++;
                                }
                                else
                                {
                                    component.ComponentName = "";
                                    _logger.LogError("form tidak upload gambar padahal ada input file", AccountID, AccountName);
                                }
                            }

                            component.IsRequired = bool.Parse(model[i][++j]);

                            var data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                            var lala = _checklistComponentAppService.Add(data);
                        }
                    }
                }

                Session["ResultLog"] = "success_Checklist created";
                return RedirectToAction("Generated/" + ChecklistID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                Session["ResultLog"] = "error_Failed to build cheklist content";
                return RedirectToAction("Content/" + ChecklistID);
            }
        }


        [CustomAuthorize("checklist")]
        public ActionResult Edit(long ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            var model = new ChecklistEditModel();

            string checklist = _checklistAppService.GetById(ID);
            if (string.IsNullOrEmpty(checklist))
            {
                return RedirectToAction("Create");
            }
            else
            {
                model.Checklist = checklist.DeserializeToChecklist();

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                model.Components = component.DeserializeToChecklistComponentList().OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();

                var componentContent = model.Components.Where(x => x.Segment.Trim() == "content").ToList();

                if (componentContent.Count() < 1)
                    return RedirectToAction("Content/" + ID);

                ViewBag.ReferenceList = BindDropDownReference();

                string submit = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                ViewBag.hasSubmitted = 1;
                if (string.IsNullOrEmpty(submit))
                {
                    ViewBag.hasSubmitted = 0;
                }

                LocationTreeModel LocationTree = GetLocationTreeModel();
                ViewBag.LocationTree = LocationTree;

                var locations = _checklistLocationAppService.FindBy("ChecklistID", ID, true);
                ViewBag.ChecklistLocations = locations.DeserializeToChecklistLocationList().Select(x => x.LocationID).ToList();

                return View(model);
            }
        }

        [CustomAuthorize("checklist")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(ChecklistEditModel model, string removed_content, List<string> Headers = null, List<string> Content = null)
        {
            try
            {
                //if (!ModelState.IsValid)
                //{
                //    return RedirectToAction("Content");
                //}

                string checklist = _checklistAppService.GetById(model.Checklist.ID, true);
                if (string.IsNullOrEmpty(checklist))
                {

                }
                else
                {
                    ChecklistModel checklistModel = checklist.DeserializeToChecklist();

                    List<long> RemoveID = new List<long>();
                    if (removed_content != "")
                        RemoveID = (removed_content.Remove(removed_content.Length - 1)).Split(',').Select(long.Parse).ToList();

                    var images = new List<string>();
                    int image_counter = 1; // karena 0 adalah icon

                    if (Request.Files != null && Request.Files.Count > 0)
                    {
                        int counter = 0;
                        foreach (string fileName in Request.Files)
                        {
                            var uploadedFile = Request.Files[fileName] as HttpPostedFileBase;
                            var image = "";
                            if (uploadedFile.ContentLength > 0)
                            {
                                Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                                image = unixTimestamp.ToString() + "_" + Path.GetFileName(uploadedFile.FileName);
                                if (counter == 0)
                                    uploadedFile.SaveAs(Server.MapPath("~/Uploads/checklist/icon/") + image);
                                else
                                    uploadedFile.SaveAs(Server.MapPath("~/Uploads/checklist/component/") + image);
                            }
                            images.Add(image);
                            counter++;
                        }
                    }


                    checklistModel.MenuTitle = model.Checklist.MenuTitle;
                    checklistModel.Header = model.Checklist.Header;
                    checklistModel.FrequencyAmount = model.Checklist.FrequencyAmount;
                    checklistModel.FrequencyDivider = model.Checklist.FrequencyDivider;
                    checklistModel.FrequencyUnit = model.Checklist.FrequencyUnit;
                    if (images[0] != "")
                        checklistModel.Icon = images[0];

                    string checklistData = JsonHelper<ChecklistModel>.Serialize(checklistModel);
                    _checklistAppService.Update(checklistData);


                    /*************** HAPUS LOCATION SEBELUMNYA ********************/

                    var locations = _checklistLocationAppService.FindByNoTracking("ChecklistID", model.Checklist.ID.ToString(), true);
                    List<ChecklistLocationModel> LocationList = locations.DeserializeToChecklistLocationList();
                    foreach (var item in LocationList)
                    {
                        _checklistLocationAppService.Remove(item.ID);
                    }

                    /*************** SIMPAN LOCATION ********************/

                    List<LocationModel> pcList = new List<LocationModel>();
                    List<LocationModel> departmentList = new List<LocationModel>();

                    if (model.Checklist.Location1 == null)
                    {
                        string currentCountry = "ID";
                        string pcs = _locationAppService.FindBy("ParentCode", currentCountry, true);
                        pcList = pcs.DeserializeToLocationList();
                    }
                    else
                    {
                        foreach (var locat in model.Checklist.Location1)
                        {
                            string pcs = _locationAppService.GetById(locat);
                            pcList.Add(pcs.DeserializeToLocation());
                        }
                    }

                    if (pcList != null)
                    {
                        foreach (var item in pcList)
                        {
                            var modelLocation = new ChecklistLocationModel();
                            modelLocation.LocationID = item.ID;
                            model.Checklist.Locations.Add(modelLocation);

                            if (model.Checklist.Location2 == null)
                            {
                                string departments = _locationAppService.FindBy("ParentID", item.ID, true);
                                var dl = departments.DeserializeToLocationList();

                                if (dl != null)
                                {
                                    foreach (var locat in dl)
                                    {
                                        departmentList.Add(locat);
                                    }
                                }
                            }
                            else
                            {
                                foreach (var locat in model.Checklist.Location2)
                                {
                                    string pcs = _locationAppService.GetById(locat);
                                    departmentList.Add(pcs.DeserializeToLocation());
                                }
                            }
                        }

                        if (departmentList != null)
                        {
                            foreach (var d in departmentList)
                            {
                                var modelLocation = new ChecklistLocationModel();
                                modelLocation.LocationID = d.ID;
                                model.Checklist.Locations.Add(modelLocation);

                                string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                                List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                                if (model.Checklist.Location3 == null)
                                {
                                    if (subdepartmentList != null)
                                    {
                                        foreach (var subdeb in subdepartmentList)
                                        {
                                            modelLocation = new ChecklistLocationModel();
                                            modelLocation.LocationID = subdeb.ID;
                                            model.Checklist.Locations.Add(modelLocation);
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var locat in model.Checklist.Location3)
                                    {
                                        modelLocation = new ChecklistLocationModel();
                                        modelLocation.LocationID = locat;
                                        model.Checklist.Locations.Add(modelLocation);
                                    }
                                }
                            }
                        }
                    }

                    foreach (var locat in model.Checklist.Locations)
                    {
                        locat.ChecklistID = model.Checklist.ID;

                        var data = JsonHelper<ChecklistLocationModel>.Serialize(locat);
                        var aaa = _checklistLocationAppService.Add(data);
                    }


                    int column_counter = 0;
                    int order_counter = 0;
                    List<string> component_done = new List<string>();

                    for (int i = 0; i < model.Components.Count(); i++)
                    {
                        string compo = _checklistComponentAppService.GetById(model.Components[i].ID, true);
                        ChecklistComponentModel component = compo.DeserializeToChecklistComponent();

                        if (component.Segment == "header")
                        {
                            component.ComponentType = model.Components[i].ComponentType.Trim();
                            component.IsRequired = model.Components[i].IsRequired;
                            component.IsDeleted = model.Components[i].IsDeleted;

                            if (component.ComponentType == "label")
                            {
                                component.OrderNum = (model.Components[i].OrderNum - 1) * 2;
                                component.ComponentName = model.Components[i].ComponentName;
                            }
                            else if (component.ComponentType == "input_option" || component.ComponentType == "input_reference")
                            {
                                component.OrderNum = ((model.Components[i - 1].OrderNum - 1) * 2) + 1;
                                component.ComponentName = model.Components[i].ComponentName;
                            }
                            else if (component.ComponentType == "image")
                            {
                                component.OrderNum = ((model.Components[i - 1].OrderNum - 1) * 2) + 1;
                                if (images[image_counter] != "")
                                    component.ComponentName = images[image_counter];
                                image_counter++;
                            }
                            else
                            {
                                component.OrderNum = ((model.Components[i - 1].OrderNum - 1) * 2) + 1;
                            }
                        }
                        else
                        {
                            continue;
                        }

                        component.ModifiedBy = AccountName;
                        component.ModifiedDate = DateTime.Now;

                        compo = JsonHelper<ChecklistComponentModel>.Serialize(component);
                        _checklistComponentAppService.Update(compo);
                    }

                    if (Headers != null && Headers.Count() > 0)
                    {
                        for (int i = 0; i < Headers.Count(); i++)
                        {
                            var component = new ChecklistComponentModel();

                            component.ChecklistID = model.Checklist.ID;
                            component.ModifiedBy = AccountName;
                            component.ModifiedDate = DateTime.Now;
                            component.Segment = "header";

                            component.OrderNum = (Int32.Parse(Headers[i]) - 1) * 2;
                            component.ComponentType = "label";
                            component.ComponentName = Headers[++i];

                            var data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                            var labelID = _checklistComponentAppService.Add(data);

                            component.OrderNum = ((Int32.Parse(Headers[i - 1]) - 1) * 2) + 1;
                            component.ComponentType = Headers[++i].Trim();

                            if (component.ComponentType == "input_option" || component.ComponentType == "input_reference")
                            {
                                component.ComponentName = Headers[++i];
                            }
                            else if (component.ComponentType == "image")
                            {
                                if (images[image_counter] != "")
                                    component.ComponentName = images[image_counter];
                                image_counter++;
                            }
                            component.IsRequired = bool.Parse(Headers[++i]);

                            data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                            var componentID = _checklistComponentAppService.Add(data);

                            string compo = _checklistComponentAppService.GetById(labelID, true);
                            component = compo.DeserializeToChecklistComponent();
                            component.IsRequired = bool.Parse(Headers[i]);
                            compo = JsonHelper<ChecklistComponentModel>.Serialize(component);
                            _checklistComponentAppService.Update(compo);
                        }
                    }

                    for (int i = 0; i < model.Components.Count(); i++)
                    {
                        string compo = _checklistComponentAppService.GetById(model.Components[i].ID, true);
                        ChecklistComponentModel component = compo.DeserializeToChecklistComponent();

                        if (component.Segment == "header")
                        {
                            continue;
                        }
                        else
                        {
                            if (component.ComponentType == "header")
                            {
                                component.ComponentName = model.Components[i].ComponentName;
                                component.AdditionalValue = model.Components[i].AdditionalValue;
                            }
                            else
                            {
                                if (RemoveID.Contains(model.Components[i].ID))
                                {
                                    component.IsDeleted = true;
                                }

                                component.ComponentType = model.Components[i].ComponentType.Trim();
                                component.IsRequired = model.Components[i].IsRequired;

                                if (!component.IsDeleted && (column_counter == 0 || model.Components[i].ColumnNum != model.Components[i - 1].ColumnNum))
                                {
                                    column_counter++;
                                    if (model.Components[i].ColumnNum != model.Components[i - 1].ColumnNum)
                                    {
                                        order_counter = 0;
                                    }
                                    else
                                    {
                                        order_counter++;
                                    }
                                }
                                component.ColumnNum = column_counter;
                                component.OrderNum = order_counter;

                                if (component.ComponentType == "input_option" || component.ComponentType == "input_radio" || component.ComponentType == "input_barcode" || component.ComponentType == "label" || component.ComponentType == "input_reference")
                                {
                                    component.ComponentName = model.Components[i].ComponentName;
                                }
                                else if (component.ComponentType == "image")
                                {
                                    if (images[image_counter] != "")
                                        component.ComponentName = images[image_counter];
                                    image_counter++;
                                }
                            }
                        }

                        component.ModifiedBy = AccountName;
                        component.ModifiedDate = DateTime.Now;

                        compo = JsonHelper<ChecklistComponentModel>.Serialize(component);
                        _checklistComponentAppService.Update(compo);

                        /*
                        if (Content != null && Content.Count() > 0)
                        {
                            for (int j = 0; j < Content.Count(); j++)
                            {
                                if (Content[j] == model.Components[i].ColumnNum.ToString())
                                {
                                    component_done.Add(Content[j]);
                                    component = new ChecklistComponentModel();

                                    component.ChecklistID = model.Checklist.ID;
                                    component.ModifiedBy = AccountName;
                                    component.ModifiedDate = DateTime.Now;
                                    component.Segment = "content";

                                    component.ColumnNum = column_counter;
                                    component.OrderNum = ++order_counter;

                                    component.ComponentType = Content[++j];

                                    if (component.ComponentType == "input_option" || component.ComponentType == "input_radio" || component.ComponentType == "input_barcode" || component.ComponentType == "label" || component.ComponentType == "input_reference")
                                    {
                                        component.ComponentName = Content[++j];
                                    }
                                    else if (component.ComponentType == "image")
                                    {
                                        component.ComponentName = images[image_counter];
                                        image_counter++;
                                    }

                                    component.IsRequired = bool.Parse(Content[++j]);

                                    var data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                                    var componentID = _checklistComponentAppService.Add(data);
                                }
                            }
                        }
                        */
                    }

                    if (Content != null && Content.Count() > 0)
                    {
                        var temp_column = "";
                        for (int j = 0; j < Content.Count(); j++)
                        {
                            if (component_done.Contains(Content[j]) == true)
                            {
                                j++; j++;
                                continue;
                            }
                            else
                            {
                                if (j == 0)
                                {
                                    column_counter++;
                                    temp_column = Content[j];
                                    order_counter = 0;
                                }
                                else if (temp_column != Content[j])
                                {
                                    column_counter++;
                                    temp_column = Content[j];
                                    order_counter = 0;
                                }
                                else
                                {
                                    order_counter++;
                                }

                                var component = new ChecklistComponentModel();

                                component.ChecklistID = model.Checklist.ID;
                                component.ModifiedBy = AccountName;
                                component.ModifiedDate = DateTime.Now;
                                component.Segment = "content";

                                component.ColumnNum = column_counter;
                                component.OrderNum = order_counter;

                                component.ComponentType = Content[++j];

                                if (component.ComponentType == "input_option" || component.ComponentType == "input_radio" || component.ComponentType == "input_barcode" || component.ComponentType == "label" || component.ComponentType == "input_reference")
                                {
                                    component.ComponentName = Content[++j];
                                }
                                else if (component.ComponentType == "image")
                                {
                                    component.ComponentName = images[image_counter];
                                    image_counter++;
                                }

                                component.IsRequired = bool.Parse(Content[++j]);

                                var data = JsonHelper<ChecklistComponentModel>.Serialize(component);
                                var componentID = _checklistComponentAppService.Add(data);
                            }
                        }
                    }
                }
                Session["ResultLog"] = "success_Checklist edited";

                return RedirectToAction("Generated/" + model.Checklist.ID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                Session["ResultLog"] = "error_Failed to edit cheklist structure";

                return RedirectToAction("Edit/" + model.Checklist.ID);
            }
        }

        //[CustomAuthorize("checklist")]
        public ActionResult Generated(long ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string checklist = _checklistAppService.GetById(ID);
            if (string.IsNullOrEmpty(checklist))
            {
                return RedirectToAction("Create");
            }
            else
            {
                string employee = _employeeAppService.GetBy("EmployeeID", AccountEmployeeID, true);
                EmployeeModel currentEmployee = employee.DeserializeToEmployee();
                Account.Employee = currentEmployee;


                /******************* CEK APPROVER VALID/TIDAK ********************/
                string approver = _checklistApproverAppService.FindBy("ChecklistID", ID, true); ;
                List<ChecklistApproverModel> approverModel = approver.DeserializeToChecklistApproverList();

                var temp_error = "";
                var approverError = new List<string>();
                if (approverModel != null && approverModel.Count() > 0)
                {
                    approverModel = approverModel.OrderBy(x => x.Tier).ToList();

                    foreach (var apvr in approverModel)
                    {
                        // harusnya tidak mungkin ini ada
                        if (apvr.EmployeeID == null || apvr.EmployeeID.Trim() == "")
                        {
                            continue;
                        }

                        if (apvr.EmployeeID.Trim() == "0")
                        {
                            if (apvr.ADGroup == "Default First Approver")
                            {
                                if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                                {
                                    apvr.EmployeeID = Account.Employee.ReportToID2;
                                }
                                else
                                {
                                    apvr.EmployeeID = Account.Employee.ReportToID1;
                                }
                            }
                            else if (apvr.ADGroup == "Default Second Approver")
                            {
                                if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                                {
                                    var emploApprover1 = _employeeAppService.GetBy("EmployeeID", Account.Employee.ReportToID2, true).DeserializeToEmployee();
                                    if (string.IsNullOrWhiteSpace(emploApprover1.ReportToID1))
                                    {
                                        apvr.EmployeeID = emploApprover1.ReportToID2;
                                    }
                                    else
                                    {
                                        apvr.EmployeeID = emploApprover1.ReportToID1;
                                    }
                                }
                                else
                                {
                                    apvr.EmployeeID = Account.Employee.ReportToID2;
                                }
                            }
                        }

                        if (string.IsNullOrWhiteSpace(apvr.EmployeeID))
                        {
                            temp_error += "Your default approver is empty. ";
                        }
                        else
                        {
                            employee = _employeeAppService.GetBy("EmployeeID", apvr.EmployeeID, true);

                            if (string.IsNullOrWhiteSpace(employee))
                            {
                                if (!approverError.Contains(apvr.EmployeeID))
                                    approverError.Add(apvr.EmployeeID);

                            }
                        }
                    }
                }
                else
                {
                    temp_error += "This checklist has no approver at all. ";
                }

                if (approverError.Count() > 0)
                {
                    foreach (var appror in approverError)
                    {
                        temp_error += "Approver with EmployeeID = (" + appror + ") has no data in Employee table. ";
                    }
                }

                if (temp_error.Trim() != "")
                    ViewBag.ResultLog = "error_" + temp_error + "Submission can never be complete. Please contact support team.";

                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                ViewBag.Checklist = currentChecklist;

                string component = _checklistComponentAppService.FindBy("ChecklistID", currentChecklist.ID, true);
                List<ChecklistComponentModel> components = component.DeserializeToChecklistComponentList();
                //components = components.OrderBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();
                //components = components.Where(x => x.Segment.Trim() == "content").ToList();
                var componentContent = components.Where(x => x.Segment.Trim() == "content").ToList();

                if (componentContent.Count() < 1)
                    return RedirectToAction("Content/" + ID);

                employee = _employeeAppService.GetAll(true);
                List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
                employeeList = employeeList.OrderBy(x => x.FullName).ToList();
                ViewBag.Userlist = employeeList;

                string reference = _referenceDetailAppService.GetAll(true);
                List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
                ViewBag.Referencelist = referenceList;

                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                return View(components);
            }
        }

        //[CustomAuthorize("checklist")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Generated(long ChecklistID, List<ChecklistValueModel> results)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Content");
                }

                string employee = _employeeAppService.GetBy("EmployeeID", AccountEmployeeID, true);
                EmployeeModel currentEmployee = employee.DeserializeToEmployee();

                var submit = new ChecklistSubmitModel();

                submit.ChecklistID = ChecklistID;
                submit.UserID = AccountID;
                submit.CompleteDate = DateTime.Now;
                submit.date = DateTime.Now;
                submit.Shift = GetShift();
                submit.IsEdited = false;
                submit.IsComplete = false;
                submit.ModifiedBy = AccountName;
                submit.ModifiedDate = DateTime.Now;

                var submitData = JsonHelper<ChecklistSubmitModel>.Serialize(submit);
                var submitID = _checklistSubmitAppService.Add(submitData);

                string checklist = _checklistAppService.GetById(ChecklistID);
                ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();

                Account.Employee = currentEmployee;

                // Approval
                // Pertama simpan data submitter
                // Lalu Insert Approver pertama

                ChecklistApprovalModel approM = new ChecklistApprovalModel
                {
                    ChecklistSubmitID = submitID,
                    ChecklistApproverID = 0,
                    Role = "Requestor",
                    EmployeeID = AccountEmployeeID,
                    Status = "Submitted",
                    Comments = "",
                    ModifiedBy = AccountName,
                    ModifiedDate = DateTime.Now
                };

                string appro = JsonHelper<ChecklistApprovalModel>.Serialize(approM);
                _checklistApprovalAppService.Add(appro);

                string approver = _checklistApproverAppService.FindBy("ChecklistID", ChecklistID, true); ;
                List<ChecklistApproverModel> approverModel = approver.DeserializeToChecklistApproverList();
                if (approverModel != null && approverModel.Count() > 0)
                {
                    ChecklistApproverModel approvr = approverModel.OrderBy(x => x.Tier).FirstOrDefault();

                    approM = new ChecklistApprovalModel
                    {
                        ChecklistSubmitID = submitID,
                        ChecklistApproverID = approvr.ID,
                        Role = "Approver",
                        EmployeeID = approvr.EmployeeID,
                        Status = "Waiting for Approval",
                        Comments = "",
                        ModifiedBy = AccountName,
                        ModifiedDate = DateTime.Now
                    };

                    if (approM.EmployeeID.Trim() == "0")
                    {
                        if (approvr.ADGroup == "Default First Approver")
                        {
                            if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                            {
                                approM.EmployeeID = Account.Employee.ReportToID2;
                            }
                            else
                            {
                                approM.EmployeeID = Account.Employee.ReportToID1;
                            }
                        }
                        else if (approvr.ADGroup == "Default Second Approver")
                        {
                            if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                            {
                                var emploApprover1 = _employeeAppService.GetBy("EmployeeID", Account.Employee.ReportToID2, true).DeserializeToEmployee();
                                if (string.IsNullOrWhiteSpace(emploApprover1.ReportToID1))
                                {
                                    approM.EmployeeID = emploApprover1.ReportToID2;
                                }
                                else
                                {
                                    approM.EmployeeID = emploApprover1.ReportToID1;
                                }
                            }
                            else
                            {
                                approM.EmployeeID = Account.Employee.ReportToID2;
                            }
                        }
                    }


                    /* EMAIL Checklist; email approver pertama; currentEmployee sudah pasti yg submit */
                    string usr = _employeeAppService.GetBy("EmployeeID", approM.EmployeeID, true);
                    var emailReceiver = usr.DeserializeToEmployee();

                    if (!string.IsNullOrEmpty(usr) && !string.IsNullOrEmpty(emailReceiver.Email))
                    {
                        await EmailSender.SendEmailChecklist(emailReceiver.Email, "[" + emailReceiver.EmployeeID + "] " + emailReceiver.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + currentEmployee.EmployeeID + "] " + currentEmployee.FullName, submitID);
                    }
                    /* END EMAIL Checklist */

                    appro = JsonHelper<ChecklistApprovalModel>.Serialize(approM);
                    _checklistApprovalAppService.Add(appro);
                }


                // SIMPAN SUBMISSION
                if (results != null)
                {
                    foreach (var result in results)
                    {
                        Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

                        result.ChecklistSubmitID = submitID;
                        result.ModifiedBy = AccountName;
                        result.ModifiedDate = DateTime.Now;

                        if (result.ValueFile != null && result.ValueFile.ContentLength > 0)
                        {
                            var tempFile = result.ValueFile;
                            result.Value = unixTimestamp.ToString() + "_" + Path.GetFileName(tempFile.FileName);
                            tempFile.SaveAs(Server.MapPath("~/Uploads/checklist/value/") + result.Value);
                            result.ValueFile = null;
                        }

                        if (result.ValueType == "webcam")
                        {
                            if (result.Value == null || result.Value.Trim() == "")
                            {
                                result.Value = "_no_image.jpg";
                            }
                            else
                            {
                                Random r = new Random();
                                var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";

                                string filePath = Server.MapPath("~/Uploads/checklist/value/") + fileName;

                                var bytes = Convert.FromBase64String(result.Value.Replace("data:image/jpeg;base64,", ""));
                                using (var imageFile = new FileStream(filePath, FileMode.Create))
                                {
                                    imageFile.Write(bytes, 0, bytes.Length);
                                    imageFile.Flush();
                                }

                                result.Value = fileName;
                            }
                        }

                        string data = JsonHelper<ChecklistValueModel>.Serialize(result);
                        _checklistValueAppService.Add(data);
                    }
                }

                Session["ResultLog"] = "success_Checklist submitted";

                return RedirectToAction("History/" + ChecklistID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                Session["ResultLog"] = "error_Failed to Submit";
                return RedirectToAction("Generated/" + ChecklistID);
            }
        }

        //[CustomAuthorize("checklist")]
        public ActionResult EditValue(string ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            var IDS = ID.Split('_').ToList();

            string submit = _checklistSubmitAppService.GetById(Int64.Parse(IDS[0]));
            ChecklistSubmitModel currentSubmit = submit.DeserializeToChecklistSubmit();

            if (currentSubmit != null && currentSubmit.IsComplete == true)
            {
                Session["ResultLog"] = "error_Checklist status is complete, cannot be edited";
                return RedirectToAction("ViewOnly/" + Int64.Parse(IDS[0]));
            }

            string checklist = _checklistAppService.GetById(currentSubmit.ChecklistID);
            ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
            ViewBag.Checklist = currentChecklist;

            string component = _checklistComponentAppService.FindBy("ChecklistID", currentChecklist.ID, true);
            List<ChecklistComponentModel> components = component.DeserializeToChecklistComponentList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            employeeList = employeeList.OrderBy(x => x.FullName).ToList();
            ViewBag.Userlist = employeeList;

            string reference = _referenceDetailAppService.GetAll(true);
            List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
            ViewBag.Referencelist = referenceList;

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", currentSubmit.ID, true);
            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

            var dictionaryID = new Dictionary<string, long>();
            var dictionary = new Dictionary<string, string>();
            var dictionaryDate = new Dictionary<string, DateTime>();

            if (listValues != null)
            {
                foreach (var value in listValues)
                {
                    dictionaryID.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.ID);
                    dictionary.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.Value);
                    dictionaryDate.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, (value.ValueDate == null) ? DateTime.Now : (DateTime)value.ValueDate);
                }
            }

            ViewBag.ListID = dictionaryID;
            ViewBag.ListValues = dictionary;
            ViewBag.ListValueDates = dictionaryDate;

            string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", currentSubmit.ID, true);
            var ApprovalRekap = approval.DeserializeToChecklistApprovalList().OrderBy(x => x.ID).ToList();
            var temp_comment = "";
            foreach (var rekap in ApprovalRekap)
            {
                string usr = _employeeAppService.GetBy("EmployeeID", rekap.EmployeeID, true);
                rekap.User = usr.DeserializeToEmployee();

                if (temp_comment == rekap.Comments)
                    rekap.Comments = "";
                temp_comment = rekap.Comments;
            }
            ViewBag.ApprovalRekap = ApprovalRekap.OrderByDescending(x => x.ID).ToList();

            return View(components);
        }

        public ActionResult ViewOnly(string ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            var IDS = ID.Split('_').ToList();

            string submit = _checklistSubmitAppService.GetById(Int64.Parse(IDS[0]));
            ChecklistSubmitModel currentSubmit = submit.DeserializeToChecklistSubmit();

            string checklist = _checklistAppService.GetById(currentSubmit.ChecklistID);
            ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
            ViewBag.Checklist = currentChecklist;

            string component = _checklistComponentAppService.FindBy("ChecklistID", currentChecklist.ID, true);
            List<ChecklistComponentModel> components = component.DeserializeToChecklistComponentList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            employeeList = employeeList.OrderBy(x => x.FullName).ToList();
            ViewBag.Userlist = employeeList;

            string reference = _referenceDetailAppService.GetAll(true);
            List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
            ViewBag.Referencelist = referenceList;

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", currentSubmit.ID, true);
            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

            var dictionaryID = new Dictionary<string, long>();
            var dictionary = new Dictionary<string, string>();
            var dictionaryDate = new Dictionary<string, DateTime>();

            if (listValues != null)
            {
                foreach (var value in listValues)
                {
                    dictionaryID.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.ID);
                    dictionary.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.Value);
                    dictionaryDate.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, (value.ValueDate == null) ? DateTime.Now : (DateTime)value.ValueDate);
                }
            }


            ViewBag.ListID = dictionaryID;
            ViewBag.ListValues = dictionary;
            ViewBag.ListValueDates = dictionaryDate;

            string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", currentSubmit.ID, true);
            var ApprovalRekap = approval.DeserializeToChecklistApprovalList().OrderBy(x => x.ID).ToList();
            var temp_comment = "";
            foreach (var rekap in ApprovalRekap)
            {
                string usr = _employeeAppService.GetBy("EmployeeID", rekap.EmployeeID, true);
                rekap.User = usr.DeserializeToEmployee();

                if (temp_comment == rekap.Comments)
                    rekap.Comments = "";
                temp_comment = rekap.Comments;
            }
            ViewBag.ApprovalRekap = ApprovalRekap.OrderByDescending(x => x.ID).ToList();

            return View(components);
        }

        //[CustomAuthorize("checklist")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> EditValue(string SubmitID, List<ChecklistValueModel> results)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("History");
                }

                var IDS = SubmitID.Split('_').ToList();

                string submit = _checklistSubmitAppService.GetById(Int64.Parse(IDS[0]), true);
                ChecklistSubmitModel currentSubmit = submit.DeserializeToChecklistSubmit();

                if (currentSubmit != null && currentSubmit.IsComplete == true)
                {
                    Session["ResultLog"] = "error_Checklist status is complete, cannot be edited";
                    return RedirectToAction("ViewOnly/" + Int64.Parse(IDS[0]));
                }

                currentSubmit.IsEdited = true;

                submit = JsonHelper<ChecklistSubmitModel>.Serialize(currentSubmit);
                _checklistSubmitAppService.Update(submit);


                // anggaplah ulang dari awal; EditValue cuma bisa diakses ketika approver minta revisi; kalau revisi minta approval mulai dari awal
                string employee = _employeeAppService.GetBy("EmployeeID", AccountEmployeeID, true);
                EmployeeModel currentEmployee = employee.DeserializeToEmployee();

                string checklist = _checklistAppService.GetById(currentSubmit.ChecklistID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();

                Account.Employee = currentEmployee;

                // Approval
                // Pertama simpan data submitter
                // Lalu Insert Approver pertama

                ChecklistApprovalModel approM = new ChecklistApprovalModel
                {
                    ChecklistSubmitID = currentSubmit.ID,
                    ChecklistApproverID = 0,
                    Role = "Requestor",
                    EmployeeID = AccountEmployeeID,
                    Status = "Revised",
                    Comments = "",
                    ModifiedBy = AccountName,
                    ModifiedDate = DateTime.Now
                };

                string appro = JsonHelper<ChecklistApprovalModel>.Serialize(approM);
                _checklistApprovalAppService.Add(appro);

                string approver = _checklistApproverAppService.FindBy("ChecklistID", currentChecklist.ID, true); ;
                List<ChecklistApproverModel> approverModel = approver.DeserializeToChecklistApproverList();
                if (approverModel != null && approverModel.Count() > 0)
                {
                    ChecklistApproverModel approvr = approverModel.OrderBy(x => x.Tier).FirstOrDefault();

                    approM = new ChecklistApprovalModel
                    {
                        ChecklistSubmitID = currentSubmit.ID,
                        ChecklistApproverID = approvr.ID,
                        Role = "Approver",
                        EmployeeID = approvr.EmployeeID,
                        Status = "Waiting for Approval",
                        Comments = "",
                        ModifiedBy = AccountName,
                        ModifiedDate = DateTime.Now
                    };

                    if (approM.EmployeeID.Trim() == "0")
                    {
                        if (approvr.ADGroup == "Default First Approver")
                        {
                            if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                            {
                                approM.EmployeeID = Account.Employee.ReportToID2;
                            }
                            else
                            {
                                approM.EmployeeID = Account.Employee.ReportToID1;
                            }
                        }
                        else if (approvr.ADGroup == "Default Second Approver")
                        {
                            if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                            {
                                var emploApprover1 = _employeeAppService.GetBy("EmployeeID", Account.Employee.ReportToID2, true).DeserializeToEmployee();
                                if (string.IsNullOrWhiteSpace(emploApprover1.ReportToID1))
                                {
                                    approM.EmployeeID = emploApprover1.ReportToID2;
                                }
                                else
                                {
                                    approM.EmployeeID = emploApprover1.ReportToID1;
                                }
                            }
                            else
                            {
                                approM.EmployeeID = Account.Employee.ReportToID2;
                            }
                        }
                    }


                    /* EMAIL Checklist; email approver pertama; currentEmployee sudah pasti yg submit */
                    string usr = _employeeAppService.GetBy("EmployeeID", approM.EmployeeID, true);
                    var emailReceiver = usr.DeserializeToEmployee();

                    if (!string.IsNullOrEmpty(usr) && !string.IsNullOrEmpty(emailReceiver.Email))
                    {
                        await EmailSender.SendEmailChecklist(emailReceiver.Email, "[" + emailReceiver.EmployeeID + "] " + emailReceiver.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + currentEmployee.EmployeeID + "] " + currentEmployee.FullName, currentSubmit.ID);
                    }
                    /* END EMAIL Checklist */

                    appro = JsonHelper<ChecklistApprovalModel>.Serialize(approM);
                    _checklistApprovalAppService.Add(appro);
                }


                /*************************** KARENA SETELAH REVISE APPROVAL RESET, INI TIDAK BERLAKU LAGI

                string approval = _checklistApprovalAppService.FindByNoTracking("ChecklistSubmitID", currentSubmit.ID.ToString(), true);
				List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();

				if (approvalList != null)
				{
					approvalList = approvalList.OrderBy(x => x.ID).ToList();

					if (approvalList != null)
					{
						foreach (var appr in approvalList)
						{
							if (appr.Status.Trim() == "Waiting for Revised")
							{
								appr.Status = "Revised";
								appr.ModifiedBy = AccountName;
								appr.ModifiedDate = DateTime.Now;
								approval = JsonHelper<ChecklistApprovalModel>.Serialize(appr);
								_checklistApprovalAppService.Update(approval);
							}
							else if (appr.Status.Trim() == "Ask to be revised")
							{
								ChecklistApprovalModel apm = new ChecklistApprovalModel
								{
									ChecklistSubmitID = appr.ChecklistSubmitID,
									ChecklistApproverID = appr.ChecklistApproverID,
									Role = "Approver",
									EmployeeID = appr.EmployeeID,
									Status = "Waiting for Approval",
									Comments = "",
									ModifiedBy = AccountName,
									ModifiedDate = DateTime.Now
								};

                                // EMAIL Checklist; email ulang approver pertama setelah direvisi
                                string usr = _employeeAppService.GetBy("EmployeeID", approM.EmployeeID, true);
                                var emailReceiver = usr.DeserializeToEmployee();

                                if (!string.IsNullOrEmpty(usr) && !string.IsNullOrEmpty(emailReceiver.Email))
                                {
                                    await EmailSender.SendEmailChecklist(emailReceiver.Email, "[" + emailReceiver.EmployeeID + "] " + emailReceiver.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + currentEmployee.EmployeeID + "] " + currentEmployee.FullName, submitID);
                                }

                                string appro = JsonHelper<ChecklistApprovalModel>.Serialize(apm);
								_checklistApprovalAppService.Add(appro);
							}
						}
					}
				}
                */

                if (results != null)
                {
                    foreach (var result in results)
                    {
                        string value = _checklistValueAppService.GetById(result.ID, true);
                        if (!string.IsNullOrEmpty(value))
                        {
                            ChecklistValueModel valueModel = value.DeserializeToChecklistValue();

                            if (result.ValueFile != null && result.ValueFile.ContentLength > 0)
                            {
                                Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                                var tempFile = result.ValueFile;
                                result.Value = unixTimestamp.ToString() + "_" + Path.GetFileName(tempFile.FileName);
                                tempFile.SaveAs(Server.MapPath("~/Uploads/checklist/value/") + result.Value);
                                result.ValueFile = null;
                            }

                            if (valueModel.ValueType == "webcam")
                            {
                                if (result.Value == null || result.Value.Trim() == "")
                                {
                                    continue;
                                }
                                else
                                {
                                    Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                                    Random r = new Random();
                                    var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";

                                    string filePath = Server.MapPath("~/Uploads/checklist/value/") + fileName;

                                    var bytes = Convert.FromBase64String(result.Value.Replace("data:image/jpeg;base64,", ""));
                                    using (var imageFile = new FileStream(filePath, FileMode.Create))
                                    {
                                        imageFile.Write(bytes, 0, bytes.Length);
                                        imageFile.Flush();
                                    }

                                    result.Value = fileName;
                                }
                            }

                            if (result.Value != valueModel.Value)
                            {
                                ChecklistValueHistoryModel history = new ChecklistValueHistoryModel();

                                history.UserID = AccountID;
                                history.ChecklistSubmitID = currentSubmit.ID;
                                history.ChecklistValueID = result.ID;
                                history.OldValue = (valueModel.Value == null ? "" : valueModel.Value);
                                history.NewValue = (result.Value == null ? "" : result.Value);
                                history.ModifiedBy = AccountName;
                                history.ModifiedDate = DateTime.Now;

                                string historyjson = JsonHelper<ChecklistValueHistoryModel>.Serialize(history);
                                var baba = _checklistValueHistoryAppService.Add(historyjson);

                                valueModel.Value = result.Value;
                                if (result.ValueDate != null && result.ValueDate.ToString() != "")
                                    valueModel.ValueDate = result.ValueDate;

                                string valuejson = JsonHelper<ChecklistValueModel>.Serialize(valueModel);
                                _checklistValueAppService.Update(valuejson);
                            }
                        }
                    }
                }

                Session["ResultLog"] = "success_Submission edited";

                if (Int64.Parse(IDS[1]) > 0)
                    return RedirectToAction("History/" + Int64.Parse(IDS[1]));
                else
                    return RedirectToAction("History");

            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                Session["ResultLog"] = "error_Failed to edit submission";

                return RedirectToAction("EditValue/" + SubmitID);
            }
        }

        [CustomAuthorize("checklisthistory")]
        public ActionResult History(long ID = 0)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string history = "";

            if (ID == 0)
            {
                history = _checklistSubmitAppService.FindBy("UserID", AccountID, true);
            }
            else
            {
                history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
            }

            List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
            historyList = historyList.OrderBy(x => x.ID).ToList();

            if (historyList != null)
            {
                foreach (var his in historyList)
                {
                    if (his != null && his.UserID > 0)
                    {
                        try
                        {
                            string user = _userAppService.GetById(his.UserID);
                            string submitter = _employeeAppService.GetBy("EmployeeID", user.DeserializeToUser().EmployeeID, true);
                            EmployeeModel submitterModel = submitter.DeserializeToEmployee();

                            his.Submiter = submitterModel;

                            string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID);
                            List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();

                            /*
                            his.isApprover = false;
                            foreach (var approvl in approvalList)
                            {
                                if (approvl.EmployeeID.Trim() == AccountEmployeeID.Trim() && approvl.Role.Trim() == "Approver")
                                {
                                    his.isApprover = true;
                                }
                            }
                            */

                            ChecklistApprovalModel approvalM = approvalList.OrderBy(x => x.ID).LastOrDefault();

                            string approver = _employeeAppService.GetBy("EmployeeID", approvalM.EmployeeID, true);
                            EmployeeModel approverModel = approver.DeserializeToEmployee();

                            string conjuction = " from ";
                            if (approvalM.Status[approvalM.Status.Length - 1] == 'd')
                                conjuction = " by ";

                            his.Status = approvalM.Status + conjuction + "[" + approvalM.EmployeeID + "] " + approverModel.FullName;
                            his.Comment = approvalM.Comments;

                            if (approvalM.EmployeeID.Trim() == AccountEmployeeID.Trim())
                                his.IsEditable = true;

                            string checklist = _checklistAppService.GetById(his.ChecklistID, true);
                            ChecklistModel checklistModel = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();

                            his.Checklist = checklistModel;
                        }
                        catch(Exception e)
                        {

                        }
                        
                    }
                }
            }

            ViewBag.Histories = historyList;
            ViewBag.ReturnTo = ID;

            return View();
        }

        [CustomAuthorize("checklistapproval")]
        public ActionResult Approval()
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string history = _checklistSubmitAppService.GetAll(true);

            List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
            historyList = historyList.OrderBy(x => x.ID).ToList();

            if (historyList != null)
            {
                foreach (var his in historyList)
                {
                    if (his != null && his.UserID > 0)
                    {
                        string user = _userAppService.GetById(his.UserID);
                        string submitter = _employeeAppService.GetBy("EmployeeID", user.DeserializeToUser().EmployeeID, true);
                        EmployeeModel submitterModel = submitter.DeserializeToEmployee();

                        his.Submiter = submitterModel;

                        string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID);
                        List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();

                        his.isApprover = false;
                        foreach (var approvl in approvalList)
                        {
                            if (approvl.EmployeeID.Trim() == AccountEmployeeID.Trim() && approvl.Role.Trim() == "Approver")
                            {
                                his.isApprover = true;
                            }
                        }

                        ChecklistApprovalModel approvalM = approvalList.OrderBy(x => x.ID).LastOrDefault();

                        string approver = _employeeAppService.GetBy("EmployeeID", approvalM.EmployeeID, true);
                        EmployeeModel approverModel = approver.DeserializeToEmployee();

                        string conjuction = " from ";
                        if (approvalM.Status[approvalM.Status.Length - 1] == 'd')
                            conjuction = " by ";

                        his.Status = approvalM.Status + conjuction + "[" + approvalM.EmployeeID + "] " + approverModel.FullName;
                        his.Comment = approvalM.Comments;

                        if (approvalM.EmployeeID.Trim() == AccountEmployeeID.Trim())
                            his.IsEditable = true;

                        string checklist = _checklistAppService.GetById(his.ChecklistID, true);
                        ChecklistModel checklistModel = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();

                        his.Checklist = checklistModel;
                    }
                }
            }

            ViewBag.Histories = historyList.OrderByDescending(x => x.ModifiedDate).ToList();
            ViewBag.MyEmployeeID = AccountEmployeeID;

            return View();
        }

        [CustomAuthorize("checklistapproval")]
        public ActionResult Submitted(long ID)
        {
            if (Session["ResultLog"] != null)
            {
                ViewBag.ResultLog = Session["ResultLog"].ToString();
                Session["ResultLog"] = null;
            }

            string submit = _checklistSubmitAppService.GetById(ID);
            ChecklistSubmitModel currentSubmit = submit.DeserializeToChecklistSubmit();

            if (currentSubmit != null && currentSubmit.IsComplete == true)
            {
                Session["ResultLog"] = "error_Checklist status is complete, cannot be edited";
                return RedirectToAction("ViewOnly/" + ID);
            }

            string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", ID, true);
            List<ChecklistApprovalModel> approvalModel = approval.DeserializeToChecklistApprovalList();
            approvalModel = approvalModel.Where(x => x.ChecklistSubmitID == ID && x.EmployeeID.Trim() == AccountEmployeeID.Trim()).ToList();

            string approver = _checklistApproverAppService.GetById(approvalModel[0].ChecklistApproverID);
            ChecklistApproverModel approverModel = approver.DeserializeToChecklistApprover();
            ViewBag.Approver = approverModel;

            string checklist = _checklistAppService.GetById(currentSubmit.ChecklistID);
            ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
            ViewBag.Checklist = currentChecklist;

            string component = _checklistComponentAppService.FindBy("ChecklistID", currentChecklist.ID, true);
            List<ChecklistComponentModel> components = component.DeserializeToChecklistComponentList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            employeeList = employeeList.OrderBy(x => x.FullName).ToList();
            ViewBag.Userlist = employeeList;

            string reference = _referenceDetailAppService.GetAll(true);
            List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
            ViewBag.Referencelist = referenceList;

            LocationTreeModel model = GetLocationTreeModel();
            ViewBag.LocationTree = model;

            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", ID, true);
            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

            var dictionaryID = new Dictionary<string, long>();
            var dictionary = new Dictionary<string, string>();
            var dictionaryDate = new Dictionary<string, DateTime>();

            if (listValues != null)
            {
                foreach (var value in listValues)
                {
                    dictionaryID.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.ID);
                    dictionary.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, value.Value);
                    dictionaryDate.Add(value.ChecklistComponentID + "_" + value.OrderNum + "_" + value.ColumnNum, (value.ValueDate == null) ? DateTime.Now : (DateTime)value.ValueDate);
                }
            }

            ViewBag.ListID = dictionaryID;
            ViewBag.ListValues = dictionary;
            ViewBag.ListValueDates = dictionaryDate;

            approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", currentSubmit.ID, true);
            var ApprovalRekap = approval.DeserializeToChecklistApprovalList().OrderBy(x => x.ID).ToList();
            var temp_comment = "";
            foreach (var rekap in ApprovalRekap)
            {
                string usr = _employeeAppService.GetBy("EmployeeID", rekap.EmployeeID, true);
                rekap.User = usr.DeserializeToEmployee();

                if (temp_comment == rekap.Comments)
                    rekap.Comments = "";
                temp_comment = rekap.Comments;
            }
            ViewBag.ApprovalRekap = ApprovalRekap.OrderByDescending(x => x.ID).ToList();

            return View(components);
        }

        [CustomAuthorize("checklistapproval")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Submitted(long SubmitID, List<ChecklistValueModel> results, string status, string comments)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    Session["ResultLog"] = "error_Data sent is not valid";
                    return RedirectToAction("History");
                }

                string submit = _checklistSubmitAppService.GetById(SubmitID);
                ChecklistSubmitModel currentSubmit = submit.DeserializeToChecklistSubmit();

                if (currentSubmit != null && currentSubmit.IsComplete == true)
                {
                    Session["ResultLog"] = "error_Checklist status is complete, cannot be edited";
                    return RedirectToAction("ViewOnly/" + SubmitID);
                }

                string approval = _checklistApprovalAppService.FindByNoTracking("ChecklistSubmitID", SubmitID.ToString(), true);
                List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();

                var approM = approvalList.OrderBy(x => x.ID).LastOrDefault();
                //hanya untuk validasi approver
                if (approM.EmployeeID == AccountEmployeeID)
                {
                    approM.Status = status;
                    approM.Comments = comments;
                    approM.ModifiedBy = AccountName;
                    approM.ModifiedDate = DateTime.Now;

                    if (status == "Approve")
                        approM.Status = "Approved";
                    else if (status == "Edit")
                        approM.Status = "Edited & Approved";
                    else if (status == "Revise")
                        approM.Status = "Ask to be revised";
                    else if (status == "Reject")
                        approM.Status = "Rejected";

                    string appro = JsonHelper<ChecklistApprovalModel>.Serialize(approM);
                    _checklistApprovalAppService.Update(appro);

                    if (status == "Approve" || status == "Edit")
                    {
                        submit = _checklistSubmitAppService.GetById(approM.ChecklistSubmitID);
                        ChecklistSubmitModel submitM = submit.DeserializeToChecklistSubmit();

                        string approver = _checklistApproverAppService.GetById(approM.ChecklistApproverID);
                        ChecklistApproverModel approverModel = approver.DeserializeToChecklistApprover();

                        string approverL = _checklistApproverAppService.FindBy("ChecklistID", submitM.ChecklistID, true);
                        List<ChecklistApproverModel> approverList = approverL.DeserializeToChecklistApproverList();

                        approverList = approverList.Where(x => x.Tier > approverModel.Tier && x.EmployeeID != AccountEmployeeID).ToList();

                        if (approverList.Count() > 0)
                        {
                            var approverM = approverList.OrderBy(x => x.Tier).FirstOrDefault();

                            string user = _userAppService.GetById(submitM.UserID);
                            var modelUser = user.DeserializeToUser();

                            string employee = _employeeAppService.GetBy("EmployeeID", modelUser.EmployeeID, true);
                            EmployeeModel submitter = employee.DeserializeToEmployee();

                            if (approverM.EmployeeID.Trim() == "0")
                            {
                                if (approverM.ADGroup == "Default First Approver")
                                {
                                    if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                                    {
                                        approverM.EmployeeID = Account.Employee.ReportToID2;
                                    }
                                    else
                                    {
                                        approverM.EmployeeID = Account.Employee.ReportToID1;
                                    }
                                }
                                else if (approverM.ADGroup == "Default Second Approver")
                                {
                                    if (string.IsNullOrWhiteSpace(Account.Employee.ReportToID1))
                                    {
                                        var emploApprover1 = _employeeAppService.GetBy("EmployeeID", Account.Employee.ReportToID2, true).DeserializeToEmployee();
                                        if (string.IsNullOrWhiteSpace(emploApprover1.ReportToID1))
                                        {
                                            approverM.EmployeeID = emploApprover1.ReportToID2;
                                        }
                                        else
                                        {
                                            approverM.EmployeeID = emploApprover1.ReportToID1;
                                        }
                                    }
                                    else
                                    {
                                        approverM.EmployeeID = Account.Employee.ReportToID2;
                                    }
                                }
                            }

                            if (approverM.EmployeeID != AccountEmployeeID)
                            {
                                ChecklistApprovalModel apm = new ChecklistApprovalModel
                                {
                                    ChecklistSubmitID = SubmitID,
                                    ChecklistApproverID = approverM.ID,
                                    Role = "Approver",
                                    EmployeeID = approverM.EmployeeID,
                                    Status = "Waiting for Approval",
                                    Comments = comments, //biar muncul di history
                                    ModifiedBy = AccountName,
                                    ModifiedDate = DateTime.Now
                                };

                                /* EMAIL Checklist; untuk approver kedua dst */
                                string checklist = _checklistAppService.GetById(approverM.ChecklistID);
                                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();

                                string usr = _employeeAppService.GetBy("EmployeeID", apm.EmployeeID, true);
                                var emailReceiver = usr.DeserializeToEmployee();

                                if (!string.IsNullOrEmpty(usr) && !string.IsNullOrEmpty(emailReceiver.Email))
                                {
                                    await EmailSender.SendEmailChecklist(emailReceiver.Email, "[" + emailReceiver.EmployeeID + "] " + emailReceiver.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + submitter.EmployeeID + "] " + submitter.FullName, SubmitID);
                                }
                                /* END EMAIL Checklist */

                                appro = JsonHelper<ChecklistApprovalModel>.Serialize(apm);
                                _checklistApprovalAppService.Add(appro);
                            }
                            else
                            {
                                string checklist = _checklistSubmitAppService.GetById(approM.ChecklistSubmitID, true);
                                ChecklistSubmitModel model = checklist.DeserializeToChecklistSubmit();

                                model.IsComplete = true; //biar nanti bisa dihitung di report

                                string data = JsonHelper<ChecklistSubmitModel>.Serialize(model);
                                _checklistSubmitAppService.Update(data);

                                /* EMAIL Checklist; untuk ngasih tau submitter kalau sudah approved */
                                string checklst = _checklistAppService.GetById(submitM.ChecklistID);
                                ChecklistModel currentChecklist = checklst.DeserializeToChecklist();

                                string usr = _employeeAppService.GetBy("EmployeeID", modelUser.EmployeeID, true);
                                var emailReceiver = usr.DeserializeToEmployee();

                                if (!string.IsNullOrEmpty(employee) && !string.IsNullOrEmpty(submitter.Email))
                                {
                                    await EmailSender.SendEmailChecklistSubmitter(submitter.Email, "[" + submitter.EmployeeID + "] " + submitter.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + AccountEmployeeID + "] " + AccountName, SubmitID, approM.Status);
                                }
                                /* END EMAIL Checklist */
                            }
                        }
                        else
                        {
                            string checklist = _checklistSubmitAppService.GetById(approM.ChecklistSubmitID, true);
                            ChecklistSubmitModel model = checklist.DeserializeToChecklistSubmit();

                            model.IsComplete = true; //biar nanti bisa dihitung di report

                            string data = JsonHelper<ChecklistSubmitModel>.Serialize(model);
                            _checklistSubmitAppService.Update(data);

                            /* EMAIL Checklist; untuk ngasih tau submitter kalau status checklistnya complete */
                            string checklst = _checklistAppService.GetById(submitM.ChecklistID);
                            ChecklistModel currentChecklist = checklst.DeserializeToChecklist();

                            string user = _userAppService.GetById(submitM.UserID);
                            var modelUser = user.DeserializeToUser();

                            string employee = _employeeAppService.GetBy("EmployeeID", modelUser.EmployeeID, true);
                            EmployeeModel submitter = employee.DeserializeToEmployee();

                            if (!string.IsNullOrEmpty(employee) && !string.IsNullOrEmpty(submitter.Email))
                            {
                                await EmailSender.SendEmailChecklistSubmitter(submitter.Email, "[" + submitter.EmployeeID + "] " + submitter.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + AccountEmployeeID + "] " + AccountName, SubmitID, approM.Status + " (Completed)");
                            }
                            /* END EMAIL Checklist */
                        }
                    }
                    else if (status == "Revise")
                    {
                        submit = _checklistSubmitAppService.GetById(approM.ChecklistSubmitID);
                        ChecklistSubmitModel submitM = submit.DeserializeToChecklistSubmit();

                        string user = _userAppService.GetById(submitM.UserID);
                        var modelUser = user.DeserializeToUser();

                        ChecklistApprovalModel apm = new ChecklistApprovalModel
                        {
                            ChecklistSubmitID = SubmitID,
                            ChecklistApproverID = 0,
                            Role = "Requestor",
                            EmployeeID = modelUser.EmployeeID,
                            Status = "Waiting for Revised",
                            Comments = comments, //biar muncul di history
                            ModifiedBy = AccountName,
                            ModifiedDate = DateTime.Now
                        };

                        appro = JsonHelper<ChecklistApprovalModel>.Serialize(apm);
                        _checklistApprovalAppService.Add(appro);

                        /* EMAIL Checklist; untuk ngasih tau submitter kalau checklistnya perlu direvisi */
                        string checklst = _checklistAppService.GetById(submitM.ChecklistID);
                        ChecklistModel currentChecklist = checklst.DeserializeToChecklist();

                        string employee = _employeeAppService.GetBy("EmployeeID", modelUser.EmployeeID, true);
                        EmployeeModel submitter = employee.DeserializeToEmployee();

                        if (!string.IsNullOrEmpty(employee) && !string.IsNullOrEmpty(submitter.Email))
                        {
                            await EmailSender.SendEmailChecklistSubmitter(submitter.Email, "[" + submitter.EmployeeID + "] " + submitter.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + AccountEmployeeID + "] " + AccountName, SubmitID, "Need to be revised");
                        }
                        /* END EMAIL Checklist */
                    }
                    else if (status == "Reject")
                    {
                        //selesai, gak bisa diapa2in lagi

                        /* EMAIL Checklist; untuk ngasih tau submitter kalau checklistnya di-reject */
                        submit = _checklistSubmitAppService.GetById(approM.ChecklistSubmitID);
                        ChecklistSubmitModel submitM = submit.DeserializeToChecklistSubmit();

                        string checklst = _checklistAppService.GetById(submitM.ChecklistID);
                        ChecklistModel currentChecklist = checklst.DeserializeToChecklist();

                        string user = _userAppService.GetById(submitM.UserID);
                        var modelUser = user.DeserializeToUser();

                        string employee = _employeeAppService.GetBy("EmployeeID", modelUser.EmployeeID, true);
                        EmployeeModel submitter = employee.DeserializeToEmployee();

                        if (!string.IsNullOrEmpty(employee) && !string.IsNullOrEmpty(submitter.Email))
                        {
                            await EmailSender.SendEmailChecklistSubmitter(submitter.Email, "[" + submitter.EmployeeID + "] " + submitter.FullName, currentChecklist.MenuTitle, currentChecklist.Header, "[" + AccountEmployeeID + "] " + AccountName, SubmitID, approM.Status + " (Completed)");
                        }
                        /* END EMAIL Checklist */
                    }
                }
                else
                {
                    Session["ResultLog"] = "error_Incorrect Approver. Who are you?";
                    return RedirectToAction("History");
                }

                if (results != null && status == "Edit")
                {
                    foreach (var result in results)
                    {
                        string value = _checklistValueAppService.GetById(result.ID, true);
                        if (!string.IsNullOrEmpty(value))
                        {
                            ChecklistValueModel valueModel = value.DeserializeToChecklistValue();

                            if (result.ValueFile != null && result.ValueFile.ContentLength > 0)
                            {
                                Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                                var tempFile = result.ValueFile;
                                result.Value = unixTimestamp.ToString() + "_" + Path.GetFileName(tempFile.FileName);
                                tempFile.SaveAs(Server.MapPath("~/Uploads/checklist/value/") + result.Value);
                                result.ValueFile = null;
                            }

                            if (result.Value != valueModel.Value)
                            {
                                ChecklistValueHistoryModel history = new ChecklistValueHistoryModel();

                                history.UserID = AccountID;
                                history.ChecklistSubmitID = SubmitID;
                                history.ChecklistValueID = result.ID;
                                history.OldValue = (valueModel.Value == null ? "" : valueModel.Value);
                                history.NewValue = (result.Value == null ? "" : result.Value);

                                string historyjson = JsonHelper<ChecklistValueHistoryModel>.Serialize(history);
                                var baba = _checklistValueHistoryAppService.Add(historyjson);

                                valueModel.Value = result.Value;
                                string valuejson = JsonHelper<ChecklistValueModel>.Serialize(valueModel);
                                _checklistValueAppService.Update(valuejson);
                            }
                        }
                    }
                }

                Session["ResultLog"] = "success_Approval status changed";

                return RedirectToAction("Approval");
            }
            catch (Exception ex)
            {
                Session["ResultLog"] = "error_Failed to change approval status";

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Submitted/" + SubmitID);
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult Delete(long id)
        {
            try
            {
                string checklist = _checklistAppService.GetById(id, true);
                ChecklistModel model = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();

                model.IsDeleted = true;

                string data = JsonHelper<ChecklistModel>.Serialize(model);
                _checklistAppService.Update(data);

                Session["ResultLog"] = "success_Checklist deleted";
            }
            catch (Exception ex)
            {
                Session["ResultLog"] = "error_Failed to delete checklist";

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }


            return RedirectToAction("Index");
        }

        [CustomAuthorize("checklist")]
        public ActionResult Report(long ID)
        {
            try
            {
                var model = new ChecklistReportSummaryModel();

                if (Session["ResultLog"] != null)
                {
                    ViewBag.ResultLog = Session["ResultLog"].ToString();
                    Session["ResultLog"] = null;
                }

                string checklist = _checklistAppService.GetById(ID, true);
                model.Checklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();

                string emplo = _employeeAppService.GetBy("EmployeeID", model.Checklist.CreatorEmployeeID, true);
                model.Checklist.Creator = emplo.DeserializeToSimpleEmployee();

                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();

                var dictionary = new Dictionary<string, int>();

                foreach (var his in historyList)
                {
                    if (his != null && his.UserID > 0)
                    {
                        string user = _userAppService.GetById(his.UserID);
                        string submitter = _employeeAppService.GetBy("EmployeeID", user.DeserializeToUser().EmployeeID, true);
                        EmployeeModel submitterModel = submitter.DeserializeToEmployee();

                        his.Submiter = submitterModel;

                        string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID);
                        List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();

                        ChecklistApprovalModel approvalM = approvalList.OrderBy(x => x.ID).LastOrDefault();

                        string approver = _employeeAppService.GetBy("EmployeeID", approvalM.EmployeeID, true);
                        EmployeeModel approverModel = approver.DeserializeToEmployee();

                        string conjuction = " from ";
                        if (approvalM.Status[approvalM.Status.Length - 1] == 'd')
                            conjuction = " by ";

                        his.Status = approvalM.Status + conjuction + "[" + approvalM.EmployeeID + "] " + approverModel.FullName;

                        if (dictionary.ContainsKey(approvalM.Status))
                        {
                            dictionary[approvalM.Status] = dictionary[approvalM.Status] + 1;
                        }
                        else
                        {
                            dictionary.Add(approvalM.Status, 1);
                        }
                    }
                }

                ViewBag.Summary = dictionary;

                return View(model);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportAdherence(long ID, string location1 = "", string location2 = "", string location3 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.location3 = location3;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location3 != "")
                {
                    arrayLocat.Add(Int64.Parse(location3));
                }
                else if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();

                ChecklistReportAdherenceModel report = new ChecklistReportAdherenceModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();

                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> userList = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }


                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        his.ContentOk = 0;
                        his.ContentAll = 0;

                        string user = _userAppService.GetById(his.UserID);
                        his.User = user.DeserializeToUser();

                        string submitter = _employeeAppService.GetBy("EmployeeID", his.User.EmployeeID, true);
                        his.Submiter = submitter.DeserializeToEmployee();

                        var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();
                        string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                        List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                        foreach (var compo in CompoContent)
                        {
                            if (compo.IsRequired == true)
                            {
                                var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else if (compo.ComponentType == "input_barcode")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value == compo.ComponentName)
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else if (compo.ComponentType == "input_webcam")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && val.Value != "_no_image.jpg")
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && val.Value.Length > 0)
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                            }
                        }
                    }
                }

                report.ReportItems = historyList;
                report.TotalAdherence = (float)(historyList.Sum(x => x.ContentOk) * 100) / historyList.Sum(x => x.ContentAll);

                return View(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportAdherenceXls(long ID, string location1 = "", string location2 = "", string location3 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.location3 = location3;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location3 != "")
                {
                    arrayLocat.Add(Int64.Parse(location3));
                }
                else if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();

                ChecklistReportAdherenceModel report = new ChecklistReportAdherenceModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                string emplo = _employeeAppService.GetBy("EmployeeID", currentChecklist.CreatorEmployeeID, true);
                currentChecklist.Creator = emplo.DeserializeToSimpleEmployee();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();

                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> userList = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }


                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        his.ContentOk = 0;
                        his.ContentAll = 0;

                        string user = _userAppService.GetById(his.UserID);
                        his.User = user.DeserializeToUser();

                        string submitter = _employeeAppService.GetBy("EmployeeID", his.User.EmployeeID, true);
                        his.Submiter = submitter.DeserializeToEmployee();

                        var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();
                        string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                        List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                        foreach (var compo in CompoContent)
                        {
                            if (compo.IsRequired == true)
                            {
                                var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else if (compo.ComponentType == "input_barcode")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value == compo.ComponentName)
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else if (compo.ComponentType == "input_webcam")
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && val.Value != "_no_image.jpg")
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                                else
                                {
                                    foreach (var val in Values)
                                    {
                                        if (val.Value != null && val.Value.Length > 0)
                                        {
                                            his.ContentOk++;
                                        }

                                        his.ContentAll++;
                                    }
                                }
                            }
                        }
                    }
                }

                report.ReportItems = historyList;
                report.TotalAdherence = (float)(historyList.Sum(x => x.ContentOk) * 100) / historyList.Sum(x => x.ContentAll);

                ExcelPackage Ep = new ExcelPackage();
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Checklist-Report-Adherence");

                using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                {
                    var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                    excelImage.SetPosition(0, 0, 0, 0);
                }

                Sheet.Cells["A3"].Value = UIResources.Title;
                Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                Sheet.Cells["B3"].Value = "Checklist Adherence Report for :" + currentChecklist.MenuTitle;
                Sheet.Cells["B4"].Value = AccountName;
                Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");


                Sheet.Cells[8, 1].Value = "Menu Title";
                Sheet.Cells[9, 1].Value = "Header";
                Sheet.Cells[10, 1].Value = "Creator";
                Sheet.Cells[11, 1].Value = "This Checklist should be submitted " + currentChecklist.FrequencyAmount + " times every " + currentChecklist.FrequencyDivider + " " + currentChecklist.FrequencyUnit;

                float perHari = currentChecklist.FrequencyAmount;
                if (currentChecklist.FrequencyUnit == "Day")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider);
                else if (currentChecklist.FrequencyUnit == "Minute")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider) * 1440;
                else if (currentChecklist.FrequencyUnit == "Hour")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider) * 24;
                else if (currentChecklist.FrequencyUnit == "Shift")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider) * 3;
                else if (currentChecklist.FrequencyUnit == "Week")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider) / 7;
                else if (currentChecklist.FrequencyUnit == "Month")
                    perHari = ((float)currentChecklist.FrequencyAmount / currentChecklist.FrequencyDivider) / 30;

                if (currentChecklist.FrequencyUnit != "Day")
                {
                    Sheet.Cells[11, 4].Value = "(equal to " + perHari + " times every 1 Day)";
                }

                Sheet.Cells[8, 2].Value = currentChecklist.MenuTitle;
                Sheet.Cells[9, 2].Value = currentChecklist.Header;
                Sheet.Cells[10, 2].Value = (currentChecklist.Creator == null ? "" : currentChecklist.Creator.FullName);

                Sheet.Cells[13, 1].Value = "Period";
                Sheet.Cells[14, 1].Value = "Total Submissions";
                Sheet.Cells[15, 1].Value = "Content Adherence";
                Sheet.Cells[16, 1].Value = "Completion Adherence";


                Sheet.Cells[13, 2].Value = ((DateTime)ViewBag.DateFrom).ToString("dd-MMM-yy") + "  -  " + ((DateTime)ViewBag.DateTo).ToString("dd-MMM-yy");
                Sheet.Cells[13, 4].Value = "(" + Math.Abs(((DateTime)ViewBag.DateTo - (DateTime)ViewBag.DateFrom).TotalDays) + " days difference)";
                Sheet.Cells[14, 2].Value = historyList.Count();
                Sheet.Cells[15, 2].Value = System.Math.Round(report.TotalAdherence, 2) + "%";
                Sheet.Cells[16, 2].Value = historyList.Count() + " * 100 / (" + perHari + " x " + Math.Abs((((DateTime)ViewBag.DateTo).Date - ((DateTime)ViewBag.DateFrom).Date).TotalDays) + ") = " + System.Math.Round(((float)historyList.Count() * 100 / (perHari * Math.Abs(((DateTime)ViewBag.DateTo - (DateTime)ViewBag.DateFrom).TotalDays))), 2) + "%";


                Sheet.Cells[18, 1].Value = "Date/Time";
                Sheet.Cells[18, 2].Value = "Shift/Group";
                Sheet.Cells[18, 3].Value = "Name";
                Sheet.Cells[18, 4].Value = "Location";
                Sheet.Cells[18, 5].Value = "Adherence";

                var counter = 19;
                foreach (var his in report.ReportItems)
                {
                    Sheet.Cells[counter, 1].Value = ((DateTime)his.CompleteDate).ToString("dd-MMM-yy HH:mm");
                    Sheet.Cells[counter, 2].Value = his.Shift + "/" + his.User.GroupName;
                    Sheet.Cells[counter, 3].Value = his.Submiter.FullName;
                    Sheet.Cells[counter, 4].Value = his.User.Location;
                    Sheet.Cells[counter, 5].Value = System.Math.Round(((float)(his.ContentOk * 100) / his.ContentAll), 2) + "%";
                    counter++;
                }

                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(HEADER_COLOR);

                using (var range = Sheet.Cells[18, 1, 18, 5])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                }

                Sheet.Column(1).Width = 22;
                Sheet.Column(2).Width = 15;
                Sheet.Column(3).Width = 30;
                Sheet.Column(4).Width = 15;
                Sheet.Column(5).Width = 15;

                Sheet.Cells[18, 1, --counter, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //Sheet.Cells["A:D"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=checklist_adherence_" + ID + ".xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();

                return View(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportByHeader(long ID, string header_id = "", string location1 = "", string location2 = "", string location3 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.location3 = location3;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location3 != "")
                {
                    arrayLocat.Add(Int64.Parse(location3));
                }
                else if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();
                var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").OrderBy(x => x.ID).ToList();

                var headerList = new Dictionary<long, string>();

                for (var i = 1; i < CompoHeader.Count(); i++)
                {
                    if ((CompoHeader[i].ComponentType == "input_text" || CompoHeader[i].ComponentType == "input_number" || CompoHeader[i].ComponentType == "input_checkbox" || CompoHeader[i].ComponentType == "input_option") && CompoHeader[i].IsRequired == true && CompoHeader[i - 1].ComponentType == "label")
                    {
                        headerList.Add(CompoHeader[i].ID, CompoHeader[i - 1].ComponentName);
                    }
                }

                if (headerList.Count() == 0)
                {
                    //terpaksa hitung juga yang ndak mandatory; daripada ndak punya acuan
                    for (var i = 1; i < CompoHeader.Count(); i++)
                    {
                        if ((CompoHeader[i].ComponentType == "input_text" || CompoHeader[i].ComponentType == "input_number" || CompoHeader[i].ComponentType == "input_checkbox" || CompoHeader[i].ComponentType == "input_option") && CompoHeader[i].IsRequired == false && CompoHeader[i - 1].ComponentType == "label")
                        {
                            headerList.Add(CompoHeader[i].ID, CompoHeader[i - 1].ComponentName);
                        }
                    }
                }

                if (headerList.Count() == 0)
                {
                    // gak punya header; ndak bisa pakai menu ini
                    Session["ResultLog"] = "error_Checklist header not found";
                    return RedirectToAction("Report/" + ID);
                }

                ViewBag.Headers = headerList;
                ViewBag.HeaderSelected = (header_id == "") ? headerList.Keys.First() : Int64.Parse(header_id);

                ChecklistReportModel report = new ChecklistReportModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> userList = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }

                List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var material = new ChecklistReportSubmitModel();

                            material.CheklistSubmitID = his.ID;
                            material.User = his.ModifiedBy;
                            material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
                            material.Datetime = (DateTime)his.ModifiedDate;

                            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                            //var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").ToList();
                            var Value = listValues.Where(x => x.ChecklistComponentID == ViewBag.HeaderSelected && x.ChecklistSubmitID == his.ID).FirstOrDefault();

                            material.Header = Value.Value;
                            material.YesOn = 0;
                            material.Counter = 0;

                            var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

                            foreach (var compo in CompoContent)
                            {
                                if (compo.IsRequired == true)
                                {
                                    var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                    if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_barcode")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value == compo.ComponentName)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_webcam")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value != "_no_image.jpg")
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value.Length > 0)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                }
                            }

                            RawMaterials.Add(material);
                        }
                    }
                }

                RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

                var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
                var MaterialGroup = new ChecklistReportSubmitGroupModel();

                var compo_temp = "";
                foreach (var raw in RawMaterials)
                {
                    if (raw.Date + "_" + raw.Header != compo_temp)
                    {
                        if (compo_temp != "")
                            MaterialGroups.Add(MaterialGroup);

                        MaterialGroup = new ChecklistReportSubmitGroupModel();

                        MaterialGroup.Date = raw.Date;
                        MaterialGroup.User = raw.User;
                        MaterialGroup.Header = raw.Header;
                        MaterialGroup.SubmitCount = 1;
                        MaterialGroup.Counter = raw.Counter;
                        MaterialGroup.YesOn = raw.YesOn;

                        compo_temp = raw.Date+"_"+raw.Header;
                    }
                    else
                    {
                        MaterialGroup.SubmitCount++;
                        MaterialGroup.Counter += raw.Counter;
                        MaterialGroup.YesOn += raw.YesOn;
                    }
                }
                MaterialGroups.Add(MaterialGroup);

                report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
                report.ReportItems = MaterialGroups;

                return View(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportByHeaderXls(long ID, string header_id = "", string location1 = "", string location2 = "", string location3 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.location3 = location3;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location3 != "")
                {
                    arrayLocat.Add(Int64.Parse(location3));
                }
                else if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();
                var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").OrderBy(x => x.ID).ToList();

                var headerList = new Dictionary<long, string>();

                for (var i = 1; i < CompoHeader.Count(); i++)
                {
                    if (CompoHeader[i].ComponentType == "input_text" && CompoHeader[i].IsRequired == true && CompoHeader[i - 1].ComponentType == "label")
                    {
                        headerList.Add(CompoHeader[i].ID, CompoHeader[i - 1].ComponentName);
                    }
                }

                ViewBag.Headers = headerList;
                ViewBag.HeaderSelected = (header_id == "") ? headerList.Keys.First() : Int64.Parse(header_id);

                ChecklistReportModel report = new ChecklistReportModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> users = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = users.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }

                List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var material = new ChecklistReportSubmitModel();

                            material.CheklistSubmitID = his.ID;
                            material.User = his.ModifiedBy;
                            material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
                            material.Datetime = (DateTime)his.ModifiedDate;

                            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                            //var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").ToList();
                            var Value = listValues.Where(x => x.ChecklistComponentID == ViewBag.HeaderSelected && x.ChecklistSubmitID == his.ID).FirstOrDefault();

                            material.Header = Value.Value;
                            material.YesOn = 0;
                            material.Counter = 0;

                            var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

                            foreach (var compo in CompoContent)
                            {
                                if (compo.IsRequired == true)
                                {
                                    var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                    if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_barcode")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value == compo.ComponentName)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_webcam")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value != "_no_image.jpg")
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value.Length > 0)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                }
                            }

                            RawMaterials.Add(material);
                        }
                    }
                }

                RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

                var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
                var MaterialGroup = new ChecklistReportSubmitGroupModel();

                var compo_temp = "";
                foreach (var raw in RawMaterials)
                {
                    if (raw.Date + "_" + raw.Header != compo_temp)
                    {
                        if (compo_temp != "")
                            MaterialGroups.Add(MaterialGroup);

                        MaterialGroup = new ChecklistReportSubmitGroupModel();

                        MaterialGroup.Date = raw.Date;
                        MaterialGroup.User = raw.User;
                        MaterialGroup.Header = raw.Header;
                        MaterialGroup.SubmitCount = 1;
                        MaterialGroup.Counter = raw.Counter;
                        MaterialGroup.YesOn = raw.YesOn;

                        compo_temp = raw.Date + "_" + raw.Header;
                    }
                    else
                    {
                        MaterialGroup.SubmitCount++;
                        MaterialGroup.Counter += raw.Counter;
                        MaterialGroup.YesOn += raw.YesOn;
                    }
                }
                MaterialGroups.Add(MaterialGroup);

                report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
                report.ReportItems = MaterialGroups;

                string[] alphabet = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

                ExcelPackage Ep = new ExcelPackage();
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("UserGroupType");
                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#808080");
                //Sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //Sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(colFromHex);


                var counter = 0; var temp = ""; List<string> date = new List<string>();
                foreach (var item in report.ReportItems)
                {
                    if (temp != item.Date + item.User)
                    {
                        temp = item.Date + item.User;
                        date.Add(item.Date);

                        Sheet.Cells[string.Format("{0}1", alphabet[++counter])].Value = item.Date;
                    }
                }

                counter = 0; temp = ""; List<string> userList = new List<string>();
                foreach (var item in report.ReportItems)
                {
                    if (temp != item.Date + item.User)
                    {
                        temp = item.Date + item.User;
                        userList.Add(item.User);

                        Sheet.Cells[string.Format("{0}2", alphabet[++counter])].Value = item.User;
                    }
                }

                var zz = 0;
                for (zz = 0; zz < report.Header.Count(); zz++)
                {
                    counter = 0;
                    Sheet.Cells[string.Format("{0}" + (zz + 3), alphabet[counter++])].Value = report.Header[zz];

                    for (var j = 0; j < date.Count(); j++)
                    {
                        string lala = "";
                        foreach (var item in report.ReportItems)
                        {
                            if (report.Header[zz] == item.Header && date[j] == item.Date && userList[j] == item.User)
                            {
                                float value = item.YesOn * 10000 / item.Counter;
                                value = value / 100;
                                item.Percentage = value;

                                lala = value.ToString();
                            }
                        }

                        Sheet.Cells[string.Format("{0}" + (zz + 3), alphabet[counter++])].Value = lala;
                    }
                }

                var startRow = zz + 5;
                Sheet.Cells[string.Format("A{0}", startRow)].Value = "AVERAGE";

                for (zz = 0; zz < report.Header.Count(); zz++)
                {
                    Sheet.Cells[string.Format("{0}" + (startRow), alphabet[zz + 1])].Value = report.Header[zz];
                }


                var dates = report.ReportItems.Select(x => x.Date).Distinct().ToList();
                for (zz = 0; zz < dates.Count(); zz++)
                {
                    Sheet.Cells[string.Format("{0}" + (++startRow), alphabet[0])].Value = @dates[zz];

                    for (var j = 0; j < report.Header.Count(); j++)
                    {
                        var average = report.ReportItems.Where(x => x.Header == report.Header[j] && x.Date == dates[zz]).ToList();
                        double percentage = 0;
                        if (average.Count() > 0)
                        {
                            percentage = Math.Round((Double)average.Average(x => x.Percentage), 2);
                        }

                        Sheet.Cells[string.Format("{0}" + (startRow), alphabet[j + 1])].Value = percentage;
                    }
                }

                startRow = startRow + 3;
                Sheet.Cells[string.Format("A{0}", startRow)].Value = "Frekwensi submit seharusnya: " + report.Checklist.FrequencyAmount + " times every " + report.Checklist.FrequencyDivider + " " + report.Checklist.FrequencyUnit;

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Actual: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount);

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Seharusnya: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.Checklist.FrequencyAmount + " x " + report.ReportItems.Count() + " = " + report.Checklist.FrequencyAmount * report.ReportItems.Count();

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Target Audience Adherence: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount) * 100 / (report.Checklist.FrequencyAmount * report.ReportItems.Count()) + " %";


                //Sheet.Cells["D1"].Value = "Group Name";

                int row = 2;
                foreach (var item in report.ReportItems)
                {
                    //Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                    row++;
                }

                Sheet.Cells["A:D"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=checklist_report_" + ID + ".xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();

                ViewBag.Result = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportByLocation(long ID, string header_id = "", string location1 = "", string location2 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();
                var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").OrderBy(x => x.ID).ToList();

                var headerList = new Dictionary<long, string>();

                for (var i = 1; i < CompoHeader.Count(); i++)
                {
                    if (CompoHeader[i].ComponentType == "input_text" && CompoHeader[i].IsRequired == true && CompoHeader[i - 1].ComponentType == "label")
                    {
                        headerList.Add(CompoHeader[i].ID, CompoHeader[i - 1].ComponentName);
                    }
                }

                ChecklistReportModel report = new ChecklistReportModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> userList = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }

                List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var material = new ChecklistReportSubmitModel();

                            material.CheklistSubmitID = his.ID;
                            material.User = his.ModifiedBy;
                            material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
                            material.Datetime = (DateTime)his.ModifiedDate;

                            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                            string user = _userAppService.GetById(his.UserID);
                            var modelUser = user.DeserializeToUser();

                            var reference = _referenceDetailAppService.GetAll(true);
                            List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
                            //referenceList = referenceList.Where(x => x.ReferenceID == 5 || x.ReferenceID == 6).ToList();
                            referenceList = referenceList.Where(x => x.ReferenceID == 5).ToList();

                            LocationModel pcList = new LocationModel();
                            string pcs = _locationAppService.GetById((long)modelUser.LocationID);
                            pcList = pcs.DeserializeToLocation();

                            if (pcList.ParentID != 1)
                            {
                                var parent = pcList.ParentID;

                                pcs = _locationAppService.GetById(parent);
                                pcList = pcs.DeserializeToLocation();
                            }
                            if (pcList.ParentID != 1)
                            {
                                var parent = pcList.ParentID;

                                pcs = _locationAppService.GetById(parent);
                                pcList = pcs.DeserializeToLocation();
                            }

                            string header = "";
                            foreach (var rf in referenceList)
                            {
                                if (rf.Code == pcList.Code)
                                {
                                    header = rf.Description;
                                }
                            }

                            material.Header = header;
                            material.YesOn = 0;
                            material.Counter = 0;

                            var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

                            foreach (var compo in CompoContent)
                            {
                                if (compo.IsRequired == true)
                                {
                                    var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                    if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_barcode")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value == compo.ComponentName)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_webcam")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value != "_no_image.jpg")
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value.Length > 0)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                }
                            }

                            RawMaterials.Add(material);
                        }
                    }
                }

                RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

                var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
                var MaterialGroup = new ChecklistReportSubmitGroupModel();

                var compo_temp = "";
                foreach (var raw in RawMaterials)
                {
                    if (raw.Header != compo_temp)
                    {
                        if (compo_temp != "")
                            MaterialGroups.Add(MaterialGroup);

                        MaterialGroup = new ChecklistReportSubmitGroupModel();

                        MaterialGroup.Date = raw.Date;
                        MaterialGroup.User = raw.User;
                        MaterialGroup.Header = raw.Header;
                        MaterialGroup.SubmitCount = 1;
                        MaterialGroup.Counter = raw.Counter;
                        MaterialGroup.YesOn = raw.YesOn;

                        compo_temp = raw.Header;
                    }
                    else
                    {
                        MaterialGroup.SubmitCount++;
                        MaterialGroup.Counter += raw.Counter;
                        MaterialGroup.YesOn += raw.YesOn;
                    }
                }
                MaterialGroups.Add(MaterialGroup);

                report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
                report.ReportItems = MaterialGroups;

                return View(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportByLocationXls(long ID, string header_id = "", string location1 = "", string location2 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();
                var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").OrderBy(x => x.ID).ToList();

                var headerList = new Dictionary<long, string>();

                for (var i = 1; i < CompoHeader.Count(); i++)
                {
                    if (CompoHeader[i].ComponentType == "input_text" && CompoHeader[i].IsRequired == true && CompoHeader[i - 1].ComponentType == "label")
                    {
                        headerList.Add(CompoHeader[i].ID, CompoHeader[i - 1].ComponentName);
                    }
                }

                ViewBag.Headers = headerList;
                ViewBag.HeaderSelected = (header_id == "") ? headerList.Keys.First() : Int64.Parse(header_id);

                ChecklistReportModel report = new ChecklistReportModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> users = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = users.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }

                List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

                if (historyList.Count() > 0)
                {
                    historyList = historyList.OrderByDescending(x => x.ID).ToList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var material = new ChecklistReportSubmitModel();

                            material.CheklistSubmitID = his.ID;
                            material.User = his.ModifiedBy;
                            material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
                            material.Datetime = (DateTime)his.ModifiedDate;

                            string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
                            List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

                            string user = _userAppService.GetById(his.UserID);
                            var modelUser = user.DeserializeToUser();

                            var reference = _referenceDetailAppService.GetAll(true);
                            List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
                            //referenceList = referenceList.Where(x => x.ReferenceID == 5 || x.ReferenceID == 6).ToList();
                            referenceList = referenceList.Where(x => x.ReferenceID == 5).ToList();

                            LocationModel pcList = new LocationModel();
                            string pcs = _locationAppService.GetById((long)modelUser.LocationID);
                            pcList = pcs.DeserializeToLocation();

                            if (pcList.ParentID != 1)
                            {
                                var parent = pcList.ParentID;

                                pcs = _locationAppService.GetById(parent);
                                pcList = pcs.DeserializeToLocation();
                            }
                            if (pcList.ParentID != 1)
                            {
                                var parent = pcList.ParentID;

                                pcs = _locationAppService.GetById(parent);
                                pcList = pcs.DeserializeToLocation();
                            }

                            string header = "";
                            foreach (var rf in referenceList)
                            {
                                if (rf.Code == pcList.Code)
                                {
                                    header = rf.Description;
                                }
                            }

                            material.Header = header;
                            material.YesOn = 0;
                            material.Counter = 0;

                            var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

                            foreach (var compo in CompoContent)
                            {
                                if (compo.IsRequired == true)
                                {
                                    var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

                                    if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on" || val.Value.Trim().ToLower() == "ya" || val.Value.Trim().ToLower() == "sip" || val.Value.Trim().ToLower() == "joss" || val.Value.Trim().ToLower() == "bagus" || val.Value.Trim().ToLower() == "benar"))
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_barcode")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value == compo.ComponentName)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else if (compo.ComponentType == "input_webcam")
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value != "_no_image.jpg")
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                    else
                                    {
                                        foreach (var val in Values)
                                        {
                                            if (val.Value != null && val.Value.Length > 0)
                                            {
                                                material.YesOn++;
                                            }

                                            material.Counter++;
                                        }
                                    }
                                }
                            }

                            RawMaterials.Add(material);
                        }
                    }
                }

                RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

                var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
                var MaterialGroup = new ChecklistReportSubmitGroupModel();

                var compo_temp = "";
                foreach (var raw in RawMaterials)
                {
                    if (raw.Header != compo_temp)
                    {
                        if (compo_temp != "")
                            MaterialGroups.Add(MaterialGroup);

                        MaterialGroup = new ChecklistReportSubmitGroupModel();

                        MaterialGroup.Date = raw.Date;
                        MaterialGroup.User = raw.User;
                        MaterialGroup.Header = raw.Header;
                        MaterialGroup.SubmitCount = 1;
                        MaterialGroup.Counter = raw.Counter;
                        MaterialGroup.YesOn = raw.YesOn;

                        compo_temp = raw.Header;
                    }
                    else
                    {
                        MaterialGroup.SubmitCount++;
                        MaterialGroup.Counter += raw.Counter;
                        MaterialGroup.YesOn += raw.YesOn;
                    }
                }
                MaterialGroups.Add(MaterialGroup);

                report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
                report.ReportItems = MaterialGroups;

                string[] alphabet = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

                ExcelPackage Ep = new ExcelPackage();
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("UserGroupType");
                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#808080");
                //Sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //Sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(colFromHex);


                var counter = 0; var temp = ""; List<string> date = new List<string>();
                foreach (var item in report.ReportItems)
                {
                    if (temp != item.Date + item.User)
                    {
                        temp = item.Date + item.User;
                        date.Add(item.Date);

                        Sheet.Cells[string.Format("{0}1", alphabet[++counter])].Value = item.Date;
                    }
                }

                counter = 0; temp = ""; List<string> userList = new List<string>();
                foreach (var item in report.ReportItems)
                {
                    if (temp != item.Date + item.User)
                    {
                        temp = item.Date + item.User;
                        userList.Add(item.User);

                        Sheet.Cells[string.Format("{0}2", alphabet[++counter])].Value = item.SubmitCount + " Submissions";
                    }
                }

                var zz = 0;
                for (zz = 0; zz < report.Header.Count(); zz++)
                {
                    counter = 0;
                    Sheet.Cells[string.Format("{0}" + (zz + 3), alphabet[counter++])].Value = report.Header[zz];

                    for (var j = 0; j < date.Count(); j++)
                    {
                        string lala = "";
                        foreach (var item in report.ReportItems)
                        {
                            if (report.Header[zz] == item.Header && date[j] == item.Date && userList[j] == item.User)
                            {
                                float value = item.YesOn * 10000 / item.Counter;
                                value = value / 100;
                                item.Percentage = value;

                                lala = value.ToString();
                            }
                        }

                        Sheet.Cells[string.Format("{0}" + (zz + 3), alphabet[counter++])].Value = lala+"%";
                    }
                }

                var startRow = zz + 5;
                Sheet.Cells[string.Format("A{0}", startRow)].Value = "AVERAGE";

                for (zz = 0; zz < report.Header.Count(); zz++)
                {
                    Sheet.Cells[string.Format("{0}" + (startRow), alphabet[zz + 1])].Value = report.Header[zz];
                }


                var dates = report.ReportItems.Select(x => x.Date).Distinct().ToList();
                for (zz = 0; zz < dates.Count(); zz++)
                {
                    Sheet.Cells[string.Format("{0}" + (++startRow), alphabet[0])].Value = @dates[zz];

                    for (var j = 0; j < report.Header.Count(); j++)
                    {
                        var average = report.ReportItems.Where(x => x.Header == report.Header[j] && x.Date == dates[zz]).ToList();
                        double percentage = 0;
                        if (average.Count() > 0)
                        {
                            percentage = Math.Round((Double)average.Average(x => x.Percentage), 2);
                        }

                        Sheet.Cells[string.Format("{0}" + (startRow), alphabet[j + 1])].Value = percentage+"%";
                    }
                }

                startRow = startRow + 3;
                Sheet.Cells[string.Format("A{0}", startRow)].Value = "Frekwensi submit seharusnya: " + report.Checklist.FrequencyAmount + " times every " + report.Checklist.FrequencyDivider + " " + report.Checklist.FrequencyUnit;

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Actual: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount);

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Seharusnya: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.Checklist.FrequencyAmount + " x " + report.ReportItems.Count() + " = " + report.Checklist.FrequencyAmount * report.ReportItems.Count();

                Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Target Audience Adherence: ";
                Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount) * 100 / (report.Checklist.FrequencyAmount * report.ReportItems.Count()) + " %";


                //Sheet.Cells["D1"].Value = "Group Name";

                int row = 2;
                foreach (var item in report.ReportItems)
                {
                    //Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                    row++;
                }

                Sheet.Cells["A:D"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=checklist_report_" + ID + ".xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();

                ViewBag.Result = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [CustomAuthorize("checklist")]
        public ActionResult ReportRawData(long ID, int download = 0, string location1 = "", string location2 = "", string location3 = "", string dateFrom = "", string dateTo = "", string shift = "")
        {
            try
            {
                LocationTreeModel model = GetLocationTreeModel();
                ViewBag.LocationTree = model;

                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

                ViewBag.location1 = location1;
                ViewBag.location2 = location2;
                ViewBag.location3 = location3;
                ViewBag.DateFrom = (dateFrom == "") ? monday : Convert.ToDateTime(dateFrom);
                ViewBag.DateTo = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);
                ViewBag.shift = shift;

                List<ParentChilds> Locations = new List<ParentChilds>();
                var arrayLocat = new List<long>();

                if (location3 != "")
                {
                    arrayLocat.Add(Int64.Parse(location3));
                }
                else if (location2 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            foreach (var depart in location.Departments)
                            {
                                if (depart.LocationID.ToString() == location2)
                                {
                                    arrayLocat.Add(depart.LocationID);
                                    foreach (var subd in depart.SubDepartments)
                                    {
                                        arrayLocat.Add(subd.LocationID);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (location1 != "")
                {
                    foreach (var location in model.ProductionCenters)
                    {
                        if (location.LocationID.ToString() == location1)
                        {
                            arrayLocat.Add(location.LocationID);
                            foreach (var depart in location.Departments)
                            {
                                arrayLocat.Add(depart.LocationID);
                                foreach (var subd in depart.SubDepartments)
                                {
                                    arrayLocat.Add(subd.LocationID);
                                }
                            }
                        }
                    }
                }

                string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                var Components = component.DeserializeToChecklistComponentList();

                ChecklistReportModel report = new ChecklistReportModel();

                string checklist = _checklistAppService.GetById(ID);
                ChecklistModel currentChecklist = checklist.DeserializeToChecklist();
                string emplo = _employeeAppService.GetBy("EmployeeID", currentChecklist.CreatorEmployeeID, true);
                currentChecklist.Creator = emplo.DeserializeToSimpleEmployee();

                report.Checklist = currentChecklist;

                //string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
                Components = Components.OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();

                string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
                List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();

                // hanya ambil data yang sesuai lokasi
                if (location1 != "")
                {
                    List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

                    string user = _userAppService.GetAll(true);
                    List<UserModel> userList = user.DeserializeToUserList();

                    foreach (var his in historyList)
                    {
                        if (his != null && his.UserID > 0)
                        {
                            var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
                            if (arrayLocat.Contains((long)usr.LocationID))
                            {
                                historyListTemp.Add(his);
                            }
                        }
                    }

                    historyList = historyListTemp;
                }

                ViewBag.SubmitCount = -1;
                if (download == 1)
                {
                    ViewBag.SubmitCount = historyList.Count();
                    if (historyList.Count() > 0)
                    {
                        ExcelPackage Ep = new ExcelPackage();
                        ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Raw Data Checklist");
                        System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#808080");
                        //Sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //Sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(colFromHex);

                        using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                        {
                            var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                            excelImage.SetPosition(0, 0, 0, 0);
                        }

                        Sheet.Cells["A3"].Value = UIResources.Title;
                        Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                        Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                        Sheet.Cells["B3"].Value = "Checklist Adherence Report for :" + currentChecklist.MenuTitle;
                        Sheet.Cells["B4"].Value = AccountName;
                        Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                        Sheet.Cells[8, 1].Value = "Menu Title";
                        Sheet.Cells[9, 1].Value = "Header";
                        Sheet.Cells[10, 1].Value = "Creator";
                        Sheet.Cells[11, 1].Value = "Total Submissions";

                        Sheet.Cells[8, 2].Value = currentChecklist.MenuTitle;
                        Sheet.Cells[9, 2].Value = currentChecklist.Header;
                        Sheet.Cells[10, 2].Value = (currentChecklist.Creator == null ? "" : currentChecklist.Creator.FullName);
                        Sheet.Cells[11, 2].Value = historyList.Count();

                        Sheet.Cells[14, 4].Value = "Tanggal :";
                        Sheet.Cells[15, 4].Value = "PIC :";

                        Sheet.Cells[15, 1].Value = "Type";
                        Sheet.Cells[15, 2].Value = "Content";
                        Sheet.Cells[15, 3].Value = "Additional Value";
                        Sheet.Cells[16, 1].Value = "Header Section";
                        Sheet.Cells["A14:D15"].Style.Font.Bold = true;
                        Sheet.Cells["A15:C15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells["D14:D15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells["A15:C15"].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                        Sheet.Cells["D14:D15"].Style.Fill.BackgroundColor.SetColor(Color.Aqua);

                        var counterTurun = 17; var counterKanan = 5;

                        foreach (var compo in Components)
                        {
                            if (compo.Segment == "header")
                            {
                                Sheet.Cells[counterTurun, 1].Value = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(compo.ComponentType.Replace("_", " "));
                                Sheet.Cells[counterTurun, 2].Value = compo.ComponentName;
                                Sheet.Cells[counterTurun, 3].Value = compo.AdditionalValue;

                                counterKanan = 5;
                                foreach (var histo in historyList)
                                {
                                    Sheet.Cells[14, counterKanan].Value = histo.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm:ss");

                                    string user = _userAppService.GetById(histo.UserID);
                                    string submitter = _employeeAppService.GetBy("EmployeeID", user.DeserializeToUser().EmployeeID, true);
                                    EmployeeModel submitterModel = submitter.DeserializeToEmployee();

                                    Sheet.Cells[15, counterKanan].Value = submitterModel.FullName;

                                    string values = _checklistValueAppService.FindBy("ChecklistSubmitID", histo.ID, true);
                                    List<ChecklistValueModel> listValue = values.DeserializeToChecklistValueList();

                                    listValue = listValue.Where(x => x.ChecklistComponentID == compo.ID).ToList();

                                    if (listValue.Count() > 0)
                                    {
                                        var modelValue = listValue.FirstOrDefault();
                                        if (modelValue != null && modelValue.Value != null)
                                            Sheet.Cells[counterTurun, counterKanan].Value = modelValue.Value;
                                    }

                                    counterKanan++;
                                }

                                counterTurun++;
                            }
                        }
                        using (var range = Sheet.Cells[16, 1, 16, counterKanan - 1])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }

                        var startContent = counterTurun;
                        Sheet.Cells[counterTurun++, 1].Value = "Content Section";

                        foreach (var compo in Components)
                        {
                            if (compo.Segment == "content")
                            {
                                Sheet.Cells[counterTurun, 1].Value = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(compo.ComponentType.Replace("_", " "));
                                Sheet.Cells[counterTurun, 2].Value = compo.ComponentName;
                                Sheet.Cells[counterTurun, 3].Value = compo.AdditionalValue;

                                counterKanan = 5;
                                foreach (var histo in historyList)
                                {
                                    Sheet.Cells[14, counterKanan].Value = histo.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm");

                                    string user = _userAppService.GetById(histo.UserID);
                                    string submitter = _employeeAppService.GetBy("EmployeeID", user.DeserializeToUser().EmployeeID, true);
                                    EmployeeModel submitterModel = submitter.DeserializeToEmployee();

                                    Sheet.Cells[15, counterKanan].Value = submitterModel.FullName;

                                    string values = _checklistValueAppService.FindBy("ChecklistSubmitID", histo.ID, true);
                                    List<ChecklistValueModel> listValue = values.DeserializeToChecklistValueList().Where(x => x.ChecklistComponentID == compo.ID).ToList();

                                    if (listValue.Count() > 0)
                                    {
                                        var modelValue = listValue.FirstOrDefault();
                                        if (modelValue != null && modelValue.Value != null)
                                        {
                                            if (compo.ComponentType == "input_barcode")
                                            {
                                                Sheet.Cells[counterTurun, counterKanan].Value = modelValue.Value + " " + (modelValue.Value == compo.ComponentName ? "(match at " + modelValue.ValueDate + ")" : "(not match)");
                                            }
                                            else
                                            {
                                                Sheet.Cells[counterTurun, counterKanan].Value = modelValue.Value;
                                            }
                                        }

                                    }

                                    counterKanan++;
                                }

                                counterTurun++;
                            }
                        }

                        using (var range = Sheet.Cells[startContent, 1, startContent, counterKanan - 1])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }

                        Sheet.Column(1).Width = 25;
                        Sheet.Column(2).Width = 30;
                        Sheet.Column(3).Width = 17;
                        Sheet.Cells[14, 5, 14, counterKanan - 1].AutoFitColumns();

                        //Sheet.Cells["A:D"].AutoFitColumns();
                        Response.Clear();
                        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        Response.AddHeader("content-disposition", "attachment;filename=checklist_raw_data_" + ID + ".xlsx");
                        Response.BinaryWrite(Ep.GetAsByteArray());
                        Response.End();
                    }
                }
                return View(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return RedirectToAction("Index");
            }
        }

        /*
        public ActionResult GenerateExcel(string ID)
		{
			try
			{
				var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
				var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

				ViewBag.DateFrom = monday;
				ViewBag.DateTo = DateTime.Now;

				ChecklistReportModel report = new ChecklistReportModel();

				string checklist = _checklistAppService.GetById(Int64.Parse(ID));
				ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
				report.Checklist = currentChecklist;

				string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
				var Components = component.DeserializeToChecklistComponentList().OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


				string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
				List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
				historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo).ToList();
				historyList = historyList.OrderByDescending(x => x.ID).ToList();

				List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

				if (historyList != null)
				{
					foreach (var his in historyList)
					{
						if (his != null && his.UserID > 0)
						{
							//yang dihitung hanya yg sudah approved
							string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();
							approvalList = approvalList.Where(x => x.Status != "Approve").ToList();

							if (approvalList.Count() > 0)
							{
								continue;
							}

							var material = new ChecklistReportSubmitModel();

							material.CheklistSubmitID = his.ID;
							material.User = his.ModifiedBy;
							material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
							material.Datetime = (DateTime)his.ModifiedDate;

							string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

							var CompoHeader = Components.Where(x => x.Segment.Trim() == "header").ToList();
							var header = "";
							foreach (var head in CompoHeader)
							{
								if (head.ComponentType == "input_text")
								{
									var Values = listValues.Where(x => x.ChecklistComponentID == head.ID && x.ChecklistSubmitID == his.ID).ToList();

									foreach (var Value in Values)
									{
										header = Value.Value;
										break;
									}
									break;
								}
							}

							material.Header = header;
							material.YesOn = 0;
							material.Counter = 0;

							var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

							foreach (var compo in CompoContent)
							{
								if (compo.IsRequired == true)
								{
									var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

									if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
									{
										foreach (var val in Values)
										{
											if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on"))
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else if (compo.ComponentType == "input_barcode")
									{
										foreach (var val in Values)
										{
											if (val.Value == compo.ComponentName)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else
									{
										foreach (var val in Values)
										{
											if (val.Value.Length > 0)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
								}
							}
							RawMaterials.Add(material);
						}
					}
				}

				RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

				var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
				var MaterialGroup = new ChecklistReportSubmitGroupModel();

				var compo_temp = "";
				foreach (var raw in RawMaterials)
				{
					if (raw.Header != compo_temp)
					{
						if (compo_temp != "")
							MaterialGroups.Add(MaterialGroup);

						MaterialGroup = new ChecklistReportSubmitGroupModel();

						MaterialGroup.Date = raw.Date;
						MaterialGroup.User = raw.User;
						MaterialGroup.Header = raw.Header;
						MaterialGroup.SubmitCount = 1;
						MaterialGroup.Counter = raw.Counter;
						MaterialGroup.YesOn = raw.YesOn;

						compo_temp = raw.Header;
					}
					else
					{
						MaterialGroup.SubmitCount++;
						MaterialGroup.Counter += raw.Counter;
						MaterialGroup.YesOn += raw.YesOn;
					}
				}
				MaterialGroups.Add(MaterialGroup);

				report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
				report.ReportItems = MaterialGroups;

				string[] alphabet = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("UserGroupType");
				System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#808080");
				//Sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
				//Sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(colFromHex);


				var counter = 0; var temp = ""; List<string> date = new List<string>();
				foreach (var item in report.ReportItems)
				{
					if (temp != item.Date + item.User)
					{
						temp = item.Date + item.User;
						date.Add(item.Date);

						Sheet.Cells[string.Format("{0}1", alphabet[++counter])].Value = item.Date;
					}
				}

				counter = 0; temp = ""; List<string> user = new List<string>();
				foreach (var item in report.ReportItems)
				{
					if (temp != item.Date + item.User)
					{
						temp = item.Date + item.User;
						user.Add(item.User);

						Sheet.Cells[string.Format("{0}2", alphabet[++counter])].Value = item.User;
					}
				}

				var i = 0;
				for (i = 0; i < report.Header.Count(); i++)
				{
					counter = 0;
					Sheet.Cells[string.Format("{0}" + (i + 3), alphabet[counter++])].Value = report.Header[i];

					for (var j = 0; j < date.Count(); j++)
					{
						string lala = "";
						foreach (var item in report.ReportItems)
						{
							if (report.Header[i] == item.Header && date[j] == item.Date && user[j] == item.User)
							{
								float value = item.YesOn * 10000 / item.Counter;
								value = value / 100;
								item.Percentage = value;

								lala = value.ToString();
							}
						}

						Sheet.Cells[string.Format("{0}" + (i + 3), alphabet[counter++])].Value = lala;
					}
				}

				var startRow = i + 5;
				Sheet.Cells[string.Format("A{0}", startRow)].Value = "AVERAGE";

				for (i = 0; i < report.Header.Count(); i++)
				{
					Sheet.Cells[string.Format("{0}" + (startRow), alphabet[i + 1])].Value = report.Header[i];
				}


				var dates = report.ReportItems.Select(x => x.Date).Distinct().ToList();
				for (i = 0; i < dates.Count(); i++)
				{
					Sheet.Cells[string.Format("{0}" + (++startRow), alphabet[0])].Value = @dates[i];

					for (var j = 0; j < report.Header.Count(); j++)
					{
						var average = report.ReportItems.Where(x => x.Header == report.Header[j] && x.Date == dates[i]).ToList();
						double percentage = 0;
						if (average.Count() > 0)
						{
							percentage = Math.Round((Double)average.Average(x => x.Percentage), 2);
						}

						Sheet.Cells[string.Format("{0}" + (startRow), alphabet[j + 1])].Value = percentage;
					}
				}

				startRow = startRow + 3;
				Sheet.Cells[string.Format("A{0}", startRow)].Value = "Frekwensi submit seharusnya: " + report.Checklist.FrequencyAmount + " times every " + report.Checklist.FrequencyDivider + " " + report.Checklist.FrequencyUnit;

				Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Actual: ";
				Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount);

				Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Jumlah Submit Seharusnya: ";
				Sheet.Cells[string.Format("B{0}", startRow)].Value = report.Checklist.FrequencyAmount + " x " + report.ReportItems.Count() + " = " + report.Checklist.FrequencyAmount * report.ReportItems.Count();

				Sheet.Cells[string.Format("A{0}", ++startRow)].Value = "Target Audience Adherence: ";
				Sheet.Cells[string.Format("B{0}", startRow)].Value = report.ReportItems.Sum(x => x.SubmitCount) * 100 / (report.Checklist.FrequencyAmount * report.ReportItems.Count()) + " %";


				//Sheet.Cells["D1"].Value = "Group Name";

				int row = 2;
				foreach (var item in report.ReportItems)
				{
					//Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
					row++;
				}

				Sheet.Cells["A:D"].AutoFitColumns();
				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=checklist_report_" + ID + ".xlsx");
				Response.BinaryWrite(Ep.GetAsByteArray());
				Response.End();

				ViewBag.Result = true;
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		[CustomAuthorize("checklist")]
		public ActionResult ReportByLocation(long ID)
		{
			try
			{
				LocationTreeModel model = GetLocationTreeModel();
				ViewBag.LocationTree = model;

				var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
				var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

				ViewBag.location1 = "";
				ViewBag.DateFrom = monday;
				ViewBag.DateTo = DateTime.Now;

				ChecklistReportModel report = new ChecklistReportModel();

				string checklist = _checklistAppService.GetById(ID);
				ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
				report.Checklist = currentChecklist;

				string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
				var Components = component.DeserializeToChecklistComponentList().OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


				string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
				List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();
                historyList = historyList.Where(x => x.date >= ViewBag.DateFrom && x.date <= ViewBag.DateTo && x.IsComplete == true).ToList();
                historyList = historyList.OrderByDescending(x => x.ID).ToList();

				List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

				if (historyList != null)
				{
					foreach (var his in historyList)
					{
						if (his != null && his.UserID > 0)
						{
							//yang dihitung hanya yg sudah approved
							string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();
							approvalList = approvalList.Where(x => x.Status != "Approve").ToList();

							if (approvalList.Count() > 0)
							{
								continue;
							}

							var material = new ChecklistReportSubmitModel();

							material.CheklistSubmitID = his.ID;
							material.User = his.ModifiedBy;
							material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
							material.Datetime = (DateTime)his.ModifiedDate;

							string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

							string user = _userAppService.GetById(his.UserID);
							var modelUser = user.DeserializeToUser();

							var reference = _referenceDetailAppService.GetAll(true);
							List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
							//referenceList = referenceList.Where(x => x.ReferenceID == 5 || x.ReferenceID == 6).ToList();
							referenceList = referenceList.Where(x => x.ReferenceID == 5).ToList();

							LocationModel pcList = new LocationModel();
							string pcs = _locationAppService.GetById((long)modelUser.LocationID);
							pcList = pcs.DeserializeToLocation();

							if (pcList.ParentID != 0)
							{
								var parent = pcList.ParentID;

								pcs = _locationAppService.GetById(parent);
								pcList = pcs.DeserializeToLocation();
							}
							if (pcList.ParentID != 0)
							{
								var parent = pcList.ParentID;

								pcs = _locationAppService.GetById(parent);
								pcList = pcs.DeserializeToLocation();
							}

							string header = "";
							foreach (var rf in referenceList)
							{
								if (rf.Code == pcList.Code)
								{
									header = rf.Description;
								}
							}

							material.Header = header;
							material.YesOn = 0;
							material.Counter = 0;

							var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

							foreach (var compo in CompoContent)
							{
								if (compo.IsRequired == true)
								{
									var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

									if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
									{
										foreach (var val in Values)
										{
											if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on"))
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else if (compo.ComponentType == "input_barcode")
									{
										foreach (var val in Values)
										{
											if (val.Value == compo.ComponentName)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else
									{
										foreach (var val in Values)
										{
											if (val.Value.Length > 0)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
								}
							}
							RawMaterials.Add(material);
						}
					}
				}

				RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

				var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
				var MaterialGroup = new ChecklistReportSubmitGroupModel();

				var compo_temp = "";
				foreach (var raw in RawMaterials)
				{
					if (raw.Header != compo_temp)
					{
						if (compo_temp != "")
							MaterialGroups.Add(MaterialGroup);

						MaterialGroup = new ChecklistReportSubmitGroupModel();

						MaterialGroup.Date = raw.Date;
						MaterialGroup.User = raw.User;
						MaterialGroup.Header = raw.Header;
						MaterialGroup.SubmitCount = 1;
						MaterialGroup.Counter = raw.Counter;
						MaterialGroup.YesOn = raw.YesOn;

						compo_temp = raw.Header;
					}
					else
					{
						MaterialGroup.SubmitCount++;
						MaterialGroup.Counter += raw.Counter;
						MaterialGroup.YesOn += raw.YesOn;
					}
				}
				MaterialGroups.Add(MaterialGroup);

				report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
				report.ReportItems = MaterialGroups;

				return View(report);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return RedirectToAction("Index");
			}
		}

		[CustomAuthorize("checklist")]
		[HttpPost]
		public ActionResult ReportByLocation(long ID, string location1, string location2, string dateFrom, string dateTo, string shift)
		{
			try
			{
				LocationTreeModel model = GetLocationTreeModel();
				ViewBag.LocationTree = model;

				ViewBag.location1 = location1;
				ViewBag.DateFrom = DateTime.Parse(dateFrom);
				ViewBag.DateTo = DateTime.Parse(dateTo);
				ViewBag.shift = shift;

				List<ParentChilds> Locations = new List<ParentChilds>();
				var arrayLocat = new List<long>();

				if (location2 != "")
				{
					foreach (var location in model.ProductionCenters)
					{
						if (location.ID.ToString() == location1)
						{
							foreach (var depart in location.Departments)
							{
								if (depart.ID.ToString() == location2)
								{
									arrayLocat.Add(depart.ID);
									foreach (var subd in depart.SubDepartments)
									{
										arrayLocat.Add(subd.ID);
									}
								}
							}
						}
					}
				}
				else if (location1 != "")
				{
					foreach (var location in model.ProductionCenters)
					{
						if (location.ID.ToString() == location1)
						{
							arrayLocat.Add(location.ID);
							foreach (var depart in location.Departments)
							{
								arrayLocat.Add(depart.ID);
								foreach (var subd in depart.SubDepartments)
								{
									arrayLocat.Add(subd.ID);
								}
							}
						}
					}
				}



				ChecklistReportModel report = new ChecklistReportModel();

				string checklist = _checklistAppService.GetById(ID);
				ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : checklist.DeserializeToChecklist();
				report.Checklist = currentChecklist;

				string component = _checklistComponentAppService.FindBy("ChecklistID", ID, true);
				var Components = component.DeserializeToChecklistComponentList().OrderByDescending(x => x.Segment).ThenBy(x => x.ColumnNum).ThenBy(x => x.OrderNum).ToList();


				string history = _checklistSubmitAppService.FindBy("ChecklistID", ID, true);
				List<ChecklistSubmitModel> historyList = history.DeserializeToChecklistSubmitList();

				DateTime DateFrom = Convert.ToDateTime(dateFrom);
				DateTime DateTo = Convert.ToDateTime(dateTo);
				historyList = historyList.Where(x => x.date >= DateFrom && x.date <= DateTo).ToList();
				if (location1 != "")
				{
					List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();

					string user = _userAppService.GetAll(true);
					List<UserModel> userList = user.DeserializeToUserList();

					foreach (var his in historyList)
					{
						if (his != null && his.UserID > 0)
						{
							var usr = userList.Where(x => x.ID == his.UserID).FirstOrDefault();
							if (arrayLocat.Contains((long)usr.LocationID))
							{
								historyListTemp.Add(his);
							}
						}
					}

					historyList = historyListTemp;
				}

				if (shift != "")
				{
					List<ChecklistSubmitModel> historyListTemp = new List<ChecklistSubmitModel>();
					var startTime = 0;
					var endTime = 0;

					if (shift == "s1")
					{
						startTime = 6;
						endTime = 14;
					}
					else if (shift == "s2")
					{
						startTime = 14;
						endTime = 22;
					}
					else if (shift == "s3")
					{
						startTime = 22;
						endTime = 6;
					}
					else if (shift == "ls1")
					{
						startTime = 6;
						endTime = 18;
					}
					else if (shift == "ls2")
					{
						startTime = 18;
						endTime = 06;
					}


					foreach (var his in historyList)
					{
						if (his != null && his.UserID > 0)
						{
							var thisTime = ((DateTime)his.ModifiedDate).ToString("HH");
							if (Int32.Parse(thisTime) >= startTime && Int32.Parse(thisTime) < endTime)
							{
								historyListTemp.Add(his);
							}
						}
					}

					historyList = historyListTemp;
				}

				historyList = historyList.OrderByDescending(x => x.ID).ToList();

				List<ChecklistReportSubmitModel> RawMaterials = new List<ChecklistReportSubmitModel>();

				if (historyList != null)
				{
					foreach (var his in historyList)
					{
						if (his != null && his.UserID > 0)
						{
							//yang dihitung hanya yg sudah approved
							string approval = _checklistApprovalAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistApprovalModel> approvalList = approval.DeserializeToChecklistApprovalList();
							approvalList = approvalList.Where(x => x.Status != "Approve").ToList();

							if (approvalList.Count() > 0)
							{
								continue;
							}

							var material = new ChecklistReportSubmitModel();

							material.CheklistSubmitID = his.ID;
							material.User = his.ModifiedBy;
							material.Date = DateTime.Parse(his.ModifiedDate.ToString()).ToString("dd-MMM-yy");
							material.Datetime = (DateTime)his.ModifiedDate;

							string values = _checklistValueAppService.FindBy("ChecklistSubmitID", his.ID, true);
							List<ChecklistValueModel> listValues = values.DeserializeToChecklistValueList();

							string user = _userAppService.GetById(his.UserID);
							var modelUser = user.DeserializeToUser();

							var reference = _referenceDetailAppService.GetAll(true);
							List<ReferenceDetailModel> referenceList = reference.DeserializeToRefDetailList();
							//referenceList = referenceList.Where(x => x.ReferenceID == 5 || x.ReferenceID == 6).ToList();
							referenceList = referenceList.Where(x => x.ReferenceID == 5).ToList();

							LocationModel pcList = new LocationModel();
							string pcs = _locationAppService.GetById((long)modelUser.LocationID);
							pcList = pcs.DeserializeToLocation();

							string header = "";
							foreach (var rf in referenceList)
							{
								if (rf.Code == pcList.Code)
								{
									header = rf.Description;
								}
							}

							material.Header = header;

							material.YesOn = 0;
							material.Counter = 0;

							var CompoContent = Components.Where(x => x.Segment.Trim() == "content").ToList();

							foreach (var compo in CompoContent)
							{
								if (compo.IsRequired == true)
								{
									var Values = listValues.Where(x => x.ChecklistComponentID == compo.ID && x.ChecklistSubmitID == his.ID).ToList();

									if (compo.ComponentType == "input_option" || compo.ComponentType == "input_radio" || compo.ComponentType == "input_checkbox")
									{
										foreach (var val in Values)
										{
											if (val.Value != null && (val.Value.Trim().ToLower() == "ok" || val.Value.Trim().ToLower() == "yes" || val.Value.Trim().ToLower() == "on"))
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else if (compo.ComponentType == "input_barcode")
									{
										foreach (var val in Values)
										{
											if (val.Value == compo.ComponentName)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
									else
									{
										foreach (var val in Values)
										{
											if (val.Value.Length > 0)
											{
												material.YesOn++;
											}

											material.Counter++;
										}
									}
								}
							}
							RawMaterials.Add(material);
						}
					}
				}

				RawMaterials = RawMaterials.OrderBy(x => x.Datetime).ThenBy(x => x.User).ThenBy(x => x.Header).ToList();

				var MaterialGroups = new List<ChecklistReportSubmitGroupModel>();
				var MaterialGroup = new ChecklistReportSubmitGroupModel();

				var compo_temp = "";
				if (RawMaterials.Count() > 0)
				{
					foreach (var raw in RawMaterials)
					{
						if (raw.Header != compo_temp)
						{
							if (compo_temp != "")
								MaterialGroups.Add(MaterialGroup);

							MaterialGroup = new ChecklistReportSubmitGroupModel();

							MaterialGroup.Date = raw.Date;
							MaterialGroup.User = raw.User;
							MaterialGroup.Header = raw.Header;
							MaterialGroup.SubmitCount = 1;
							MaterialGroup.Counter = raw.Counter;
							MaterialGroup.YesOn = raw.YesOn;

							compo_temp = raw.Header;
						}
						else
						{
							MaterialGroup.SubmitCount++;
							MaterialGroup.Counter += raw.Counter;
							MaterialGroup.YesOn += raw.YesOn;
						}
					}
					MaterialGroups.Add(MaterialGroup);
				}

				report.Header = MaterialGroups.Select(x => x.Header).Distinct().ToList();
				report.ReportItems = MaterialGroups;

				return View(report);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return RedirectToAction("ReportByLocation/" + ID);
			}
		}
        */

        [CustomAuthorize("checklist")]
        public ActionResult DeleteSubmission(string id)
        {
            try
            {
                var IDS = id.Split('_').ToList();

                string checklist = _checklistSubmitAppService.GetById(Int64.Parse(IDS[0]), true);
                ChecklistSubmitModel model = checklist.DeserializeToChecklistSubmit();

                model.IsDeleted = true;

                string data = JsonHelper<ChecklistSubmitModel>.Serialize(model);
                _checklistSubmitAppService.Update(data);

                Session["ResultLog"] = "success_Submission deleted";

                if (Int64.Parse(IDS[1]) > 0)
                    return RedirectToAction("History/" + Int64.Parse(IDS[1]));
                else
                    return RedirectToAction("History");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                Session["ResultLog"] = "error_Failed to delete submission";
                return RedirectToAction("Index");
            }
        }

        public ActionResult DashboardWeekly()
        {
            string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
            int[] Submitted = { 0, 0, 0, 0, 0 };
            int[] Approved = { 0, 0, 0, 0, 0 };
            var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

            try
            {
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM ChecklistSubmits WHERE (convert(date,[CompleteDate]) >= '" + monday.Date.ToShortDateString() + "' AND IsDeleted = 0)";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<ChecklistSubmitModel> submissions = jsondata.DeserializeToChecklistSubmitList();

                if (submissions != null && submissions.Count() > 0)
                {
                    var userIDs = submissions.Select(x => x.UserID).Distinct().ToList();

                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM Users WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", userIDs) + ")";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }

                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    List<UserModel> users = jsondata.DeserializeToUserList();

                    foreach (var user in users)
                    {
                        var user_subs = submissions.Where(x => x.UserID == user.ID).ToList();
                        var submitted = user_subs.Count();
                        var approved = user_subs.Where(x => x.IsComplete == true).Count();

                        Submitted[0] = Submitted[0] + submitted;
                        Approved[0] = Approved[0] + approved;

                        if (user.Location.Contains("ID-PJ"))
                        {
                            Submitted[1] = Submitted[1] + submitted;
                            Approved[1] = Approved[1] + approved;
                        }
                        else if (user.Location.Contains("ID-PK"))
                        {
                            Submitted[2] = Submitted[2] + submitted;
                            Approved[2] = Approved[2] + approved;
                        }
                        else if (user.Location.Contains("ID-PI"))
                        {
                            Submitted[3] = Submitted[3] + submitted;
                            Approved[3] = Approved[3] + approved;
                        }
                        else if (user.Location.Contains("ID-PB"))
                        {
                            Submitted[4] = Submitted[4] + submitted;
                            Approved[4] = Approved[4] + approved;
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }
            return Json(new { Header = Header, Submitted = Submitted, Approved = Approved }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DashboardDaily()
        {
            string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
            int[] Submitted = { 0, 0, 0, 0, 0 };
            int[] Approved  = { 0, 0, 0, 0, 0 };

            try
            {

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM ChecklistSubmits WHERE (convert(date,[CompleteDate]) >= '" + DateTime.Now.Date.ToShortDateString() + "' AND IsDeleted = 0)";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<ChecklistSubmitModel> submissions = jsondata.DeserializeToChecklistSubmitList();

                if (submissions != null && submissions.Count() > 0)
                {
                    var userIDs = submissions.Select(x => x.UserID).Distinct().ToList();

                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM Users WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", userIDs) + ")";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }

                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    List<UserModel> users = jsondata.DeserializeToUserList();

                    foreach (var user in users)
                    {
                        var user_subs = submissions.Where(x => x.UserID == user.ID).ToList();
                        var submitted = user_subs.Count();
                        var approved = user_subs.Where(x => x.IsComplete == true).Count();

                        Submitted[0] = Submitted[0] + submitted;
                        Approved[0] = Approved[0] + approved;

                        if (user.Location.Contains("ID-PJ"))
                        {
                            Submitted[1] = Submitted[1] + submitted;
                            Approved[1] = Approved[1] + approved;
                        }
                        else if (user.Location.Contains("ID-PK"))
                        {
                            Submitted[2] = Submitted[2] + submitted;
                            Approved[2] = Approved[2] + approved;
                        }
                        else if (user.Location.Contains("ID-PI"))
                        {
                            Submitted[3] = Submitted[3] + submitted;
                            Approved[3] = Approved[3] + approved;
                        }
                        else if (user.Location.Contains("ID-PB"))
                        {
                            Submitted[4] = Submitted[4] + submitted;
                            Approved[4] = Approved[4] + approved;
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }

            return Json(new { Header = Header, Submitted = Submitted, Approved = Approved }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Error(int ID = 1)
        {
            ViewBag.Code = ID;
            return View();
        }

        private LocationTreeModel GetLocationTreeModel()
        {
            LocationTreeModel model = new LocationTreeModel();

            int index = 1;

            // get production center list
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var item in pcList)
            {
                string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
                ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
                model.ProductionCenters.Add(pcModel);
            }

            // get department list
            foreach (var pc in model.ProductionCenters)
            {
                LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
                string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
                List<LocationModel> departmentList = departments.DeserializeToLocationList();

                foreach (var d in departmentList)
                {
                    string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
                    DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

                    string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                    List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                    foreach (var subdeb in subdepartmentList)
                    {
                        string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
                        deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
                    }

                    pc.Departments.Add(deptModel);
                }
            }

            return model;
        }

        private Boolean IsNumber(String value)
        {
            value = value.Trim();
            return value.All(Char.IsDigit);
        }

        [HttpPost]
        public string GetDateTime()
        {
            return DateTime.Now.ToString();
        }
    }
}
