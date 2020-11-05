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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Mvc;

namespace Fast.Web.Controllers
{
    [CustomAuthorize("mpp")]
    public class MPPController : BaseController<MppModel>
    {
        #region ::Variables::
        private readonly IMppAppService _mppAppService;
        private readonly IWppStpAppService _wppStpAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IReferenceDetailAppService _refDetailAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly ILoggerAppService _logger;
        private readonly IUserAppService _userAppService;
        private readonly IEmployeeAppService _empAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IJobTitleAppService _jobTitleAppService;
        private readonly IMenuAppService _menuService;
        private readonly IMppChangesAppService _mppChangesAppService;
        private readonly ICalendarHolidayAppService _calendarHolidayAppService;
        private readonly ICalendarAppService _calendarAppService;
        private readonly IUserMachineAppService _userMachineAppService;
        private readonly IEmployeeLeaveAppService _employeeLeaveAppService;
        private readonly IEmployeeOvertimeAppService _employeeOvertimeAppService;
        private readonly IMachineAllocationAppService _machineAllocationAppService;
        private readonly IUserMachineTypeAppService _employeeSkillAppService;
        private readonly IUserRoleAppService _userRoleAppService;
        private readonly ITrainingAppService _trainingAppService;
        private readonly IManPowerAppService _manPowerAppService;
        #endregion

        #region ::Constructor::
        public MPPController(
            IManPowerAppService manPowerAppService,
            ITrainingAppService trainingAppService,
            IMppAppService mppAppService,
            IWppStpAppService wppStpAppService,
            IReferenceAppService referenceAppService,
            ILocationAppService locationAppService,
            IMachineAppService machineAppService,
            IUserAppService userAppService,
            IEmployeeAppService empAppService,
            IJobTitleAppService jobTitleService,
            IMenuAppService menuService,
            IReferenceDetailAppService refDetailAppService,
            ILoggerAppService logger,
            IMppChangesAppService mppChangesAppService,
            ICalendarHolidayAppService calendarHolidayAppService,
            ICalendarAppService calendarAppService,
            IUserMachineAppService userMachineAppService,
            IEmployeeLeaveAppService employeeLeaveAppService,
            IMachineAllocationAppService machineAllocationAppService,
            IUserMachineTypeAppService employeeSkillAppService,
            IEmployeeOvertimeAppService employeeOvertimeAppService,
            IUserRoleAppService userRoleAppService)
        {
            _manPowerAppService = manPowerAppService;
            _trainingAppService = trainingAppService;
            _wppStpAppService = wppStpAppService;
            _refDetailAppService = refDetailAppService;
            _referenceAppService = referenceAppService;
            _mppAppService = mppAppService;
            _logger = logger;
            _locationAppService = locationAppService;
            _machineAppService = machineAppService;
            _userAppService = userAppService;
            _empAppService = empAppService;
            _jobTitleAppService = jobTitleService;
            _menuService = menuService;
            _mppChangesAppService = mppChangesAppService;
            _calendarHolidayAppService = calendarHolidayAppService;
            _userMachineAppService = userMachineAppService;
            _employeeLeaveAppService = employeeLeaveAppService;
            _calendarAppService = calendarAppService;
            _employeeOvertimeAppService = employeeOvertimeAppService;
            _machineAllocationAppService = machineAllocationAppService;
            _employeeSkillAppService = employeeSkillAppService;
            _userRoleAppService = userRoleAppService;
        }
        #endregion

        #region ::Ajax call::
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
        #endregion

        #region ::DropDown::
        private List<SelectListItem> BindDropDownMachine()
        {
            string dataList = _machineAppService.GetAll(true);
            List<MachineModel> dataModelList = dataList.DeserializeToMachineList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Code,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownJobTitle()
        {
            string dataList = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> dataModelList = dataList.DeserializeToJobTitleList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Title,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownUser()
        {
            string dataList = _userAppService.GetAll(true);
            List<UserModel> dataModelList = dataList.DeserializeToUserList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                string emp = _empAppService.GetBy("EmployeeID", data.EmployeeID, true);
                EmployeeModel empModel = emp.DeserializeToEmployee();

                _menuList.Add(new SelectListItem
                {
                    Text = empModel.FullName,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        private List<SelectListItem> BindDropDownUserWithFilter(int mppID, string location, string empMachine, string stateMach)
        {
            MppModel model = GetMpp(mppID);

            string pcLocation = location.Substring(0, 5);
            long pcID = _locationAppService.GetLocationID(pcLocation);
            List<long> locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");

            string outUser = string.Empty;
            List<SelectListItem> _menuList = new List<SelectListItem>();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("EmployeeMachine", empMachine));
            filters.Add(new QueryFilter("GroupName", stateMach));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string mppList = _mppAppService.Find(filters);
            List<MppModel> mpp = mppList.DeserializeToMppList();
            mpp = mpp.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

            string machineList = _machineAppService.GetAll(true);
            List<MachineModel> machineData = machineList.DeserializeToMachineList();
            MachineModel dataMachine = machineData.Where(x => locIDList.Any(y => y == x.LocationID) && x.Code == empMachine).FirstOrDefault();

            if (dataMachine != null)
            {
                string userMachineList = _userMachineAppService.FindBy("MachineID", dataMachine.ID, true);
                List<UserMachineModel> dataUserMachine = userMachineList.DeserializeToUserMachineList().Distinct().ToList();

                string emLeaveList = _employeeLeaveAppService.GetAll();
                List<EmployeeLeaveModel> emLeaveModelList = emLeaveList.DeserializeToEmployeeLeaveList();

                if (dataUserMachine.Count > 0)
                {
                    foreach (var item in dataUserMachine)
                    {
                        string empID = GetEmployeeID(item.UserID);
                        foreach (var itemMpp in mpp)
                        {
                            if (itemMpp.EmployeeID.Trim() != empID)
                            {
                                bool isEmLeaveA = IsEmployeeOnLeave(itemMpp.EmployeeID, model.Date, emLeaveModelList);
                                if (isEmLeaveA == false && !_menuList.Any(x => x.Value == itemMpp.EmployeeID))
                                {
                                    _menuList.Add(new SelectListItem
                                    {
                                        Text = itemMpp.EmployeeName,
                                        Value = itemMpp.EmployeeID
                                    });
                                }
                            }
                        }
                    }
                }
                else
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = model.EmployeeName,
                        Value = model.EmployeeID
                    });
                }
            }
            else
            {
                _menuList.Add(new SelectListItem
                {
                    Text = model.EmployeeName,
                    Value = model.EmployeeID
                });
            }

            return _menuList;
        }

        private List<SelectListItem> GetUserListByFilterAndShift(int mppID, long locID, string empMachine, string shift)
        {
            MppModel model = GetMpp(mppID);

            string outUser = string.Empty;
            List<SelectListItem> _menuList = new List<SelectListItem>();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", model.Date.ToString()));
            filters.Add(new QueryFilter("LocationID", locID.ToString()));
            filters.Add(new QueryFilter("EmployeeMachine", empMachine, Operator.Contains));
            filters.Add(new QueryFilter("StatusMPP", "Normal"));
            filters.Add(new QueryFilter("Shift", shift));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string mppList = _mppAppService.Find(filters);
            List<MppModel> mpp = mppList.DeserializeToMppList().OrderBy(x => x.StartDate).Distinct().ToList();

            string emLeaveList = _employeeLeaveAppService.GetAll();
            List<EmployeeLeaveModel> emLeaveModelList = emLeaveList.DeserializeToEmployeeLeaveList();

            foreach (var itemMpp in mpp)
            {
                bool isEmLeaveA = IsEmployeeOnLeave(itemMpp.EmployeeID, model.Date, emLeaveModelList);
                if (isEmLeaveA == false && !_menuList.Any(x => x.Value == itemMpp.EmployeeID))
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = itemMpp.EmployeeName,
                        Value = itemMpp.ID.ToString()
                    });
                }
            }

            return _menuList;
        }
        #endregion

        #region ::Public Methods::
        public ActionResult Index()
        {
            GetTempData();

            MppModel model = GetIndexModel();

            return View(model);
        }

        public ActionResult List()
        {
            MppModel model = GetIndexModel();

            return View(model);
        }

        public ActionResult Status()
        {
            MppModel model = GetIndexModel();
            ViewBag.ShiftList = BindDropDownShift();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupType(_refDetailAppService);

            return View(model);
        }

        public static List<SelectListItem> BindDropDownShift()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "- Select -",
                Value = "0"
            });

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

            _menuList.Add(new SelectListItem
            {
                Text = "NS",
                Value = "NS"
            });

            return _menuList;
        }

        public ActionResult Detail(string machineCode, string date, long locID, string locType, bool isNextDay = false, string shift = "1")
        {
            GetTempData();

            DateTime dateFL = DateTime.Parse(date);
            if (isNextDay)
                dateFL = dateFL.AddDays(1);

            string groupname = GetGroupName(locID, shift, dateFL);

            List<QueryFilter> filter = new List<QueryFilter>();
            filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd")));
            filter.Add(new QueryFilter("Shift", shift.ToString()));
            filter.Add(new QueryFilter("GroupName", groupname));

            string mpps = _mppAppService.Find(filter);
            List<MppModel> mppList = mpps.DeserializeToMppList();

            List<long> locIDList = _locationAppService.GetLocIDListByLocType(locID, locType);

            mppList = mppList.Where(x => x.EmployeeMachine.Contains(machineCode) && locIDList.Any(y => y == x.LocationID)).OrderBy(x => x.JobTitle).ToList();

            string machineAllocations = _machineAllocationAppService.FindBy("MachineCode", machineCode);
            List<MachineAllocationModel> machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();
            machineAllocationList = machineAllocationList.OrderBy(x => x.MachineCategory).ToList();

            MppAllocationModel model = new MppAllocationModel();
            model.AllocationList = new List<MachineAllocationModel>();
            model.MppList = new List<MppModel>();

            if (mppList.Count > machineAllocationList.Count)
            {
                int x = 0;
                foreach (var item in mppList)
                {
                    model.MppList.Add(item);
                    if (x < machineAllocationList.Count)
                    {
                        model.AllocationList.Add(machineAllocationList[x++]);
                    }
                    else
                    {
                        model.AllocationList.Add(new MachineAllocationModel());
                    }
                }
            }
            else
            {
                int x = 0;
                foreach (var item in machineAllocationList)
                {
                    model.AllocationList.Add(item);
                    if (x < mppList.Count)
                    {
                        model.MppList.Add(mppList[x++]);
                    }
                    else
                    {
                        model.MppList.Add(new MppModel());
                    }
                }
            }

            var allocationList = model.AllocationList.Where(x => x.MachineCategory != null).OrderBy(x => x.MachineCategory).ToList();
            int allocationListNumber = allocationList.Count;
            var temp = model.AllocationList.Where(x => x.MachineCategory == null).ToList();
            if (temp.Count > 0)
                allocationList.AddRange(temp);

            var mppModelList = model.MppList.OrderBy(x => x.JobTitle).ToList();
            var tempMpp = model.MppList.Where(x => x.JobTitle == null).ToList();
            if (tempMpp.Count > 0)
                mppModelList.AddRange(tempMpp);

            List<MppModel> mpList = new List<MppModel>();
            List<MachineAllocationModel> maList = new List<MachineAllocationModel>();

            if (allocationListNumber > 0)
            {
                foreach (var item in allocationList)
                {
                    if (item.MachineCategory != null)
                    {
                        var mpTempList = mppModelList.Where(x => x.JobTitle != null && x.JobTitle.Contains(item.MachineCategory)).ToList();
                        if (mpTempList.Count > 0)
                        {
                            int index = 0;
                            foreach (var mp in mpTempList)
                            {
                                if (index == 0)
                                {
                                    maList.Add(item);
                                    mpList.Add(mp);
                                    index++;
                                }
                                else
                                {
                                    mpList.Add(mp);
                                    maList.Add(new MachineAllocationModel());
                                }
                            }
                        }
                        else
                        {
                            maList.Add(item);
                            mpList.Add(new MppModel());
                        }
                    }
                    else
                    {
                        mpList.Add(new MppModel());
                    }
                }
            }
            else
            {
                mpList = mppModelList;
                maList = allocationList;
            }

            model.AllocationList = maList;
            model.MppList = mpList;

            ViewBag.MachineCode = machineCode;
            ViewBag.LocationID = locID;
            ViewBag.LocationType = locType;
            ViewBag.MppDate = date;
            ViewBag.IsNextDay = isNextDay;
            ViewBag.Shift = shift;

            return View(model);
        }

        public List<List<SelectListItem>> GetUserWithCondition(string machineCode, long locID, string locType, string date, string shift, string JobTitle, string StatusMpp, string empID = "")
        {
            var today = DateTime.Parse(date);
            var yesterday = today.AddDays(-1);
            var tomorrow = today.AddDays(1);

            // list of list untuk akomodasi return 2 list (LS1 & LS2)
            var model = new List<List<SelectListItem>>();

            //hanya ambil yang job title nya cocok
            string dataList = _jobTitleAppService.FindBy("RoleName", JobTitle, true);
            var JobTitleIDs = dataList.DeserializeToJobTitleList().Select(x => x.ID).ToList();

            string dataList2 = _userRoleAppService.FindBy("RoleName", JobTitle, true);
            var UserIDs = dataList2.DeserializeToUserRoleList().Select(x => x.UserID).ToList();

            string user = _userAppService.GetAll(true);
            var userModelList = user.DeserializeToUserList();
            List<UserModel> UserList = new List<UserModel>();

            string empSkills = _employeeSkillAppService.GetAll(true);
            List<UserMachineTypeModel> empSkillList = empSkills.DeserializeToUserMachineTypeList();

            string userMachines = _userMachineAppService.GetAll(true);
            List<UserMachineModel> userMachineList = userMachines.DeserializeToUserMachineList();

            MachineModel machineModel = _machineAppService.GetBy("Code", machineCode).DeserializeToMachine();

            string location = _locationAppService.GetLocationFullCode(locID);

            List<long> locIDList = new List<long>();
            long pcID = locID;

            if (locType == "department")
            {
                // get the production center
                location = location.Substring(0, 5);
                pcID = _locationAppService.GetLocationID(location);

                locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
            }
            else
            {
                locIDList = _locationAppService.GetLocIDListByLocType(locID, locType);
                UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
            }

            if (UserList.Count == 0)
            {
                if (location == "ID-PB")
                {
                    pcID = _locationAppService.GetLocationID("ID-PJ");
                    locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                    UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
                }
                else if (location == "ID-PJ")
                {
                    pcID = _locationAppService.GetLocationID("ID-PB");
                    locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                    UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
                }
                else if (location == "ID-PI")
                {
                    pcID = _locationAppService.GetLocationID("ID-PK");
                    locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                    UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
                }
                else if (location == "ID-PK")
                {
                    pcID = _locationAppService.GetLocationID("ID-PI");
                    locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
                    UserList = userModelList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                    UserList = UserList.Where(x => JobTitleIDs.Contains(x.JobTitleID) || UserIDs.Contains(x.ID)).ToList();
                }
            }

            List<EmployeeModel> empList = new List<EmployeeModel>();
            //get nama lengkap
            foreach (var usr in UserList)
            {
                if (empList.Count == 0)
                {
                    empList = _empAppService.GetAll(true).DeserializeToEmployeeList();
                }

                EmployeeModel empModel = empList.Where(x => x.EmployeeID.Trim() == usr.EmployeeID.Trim()).FirstOrDefault();

                usr.FullName = empModel == null ? string.Empty : empModel.FullName;
            }

            string groupname = GetGroupName(locID, shift, today);

            if (StatusMpp == "Normal")
            {
                // exclude user yg masuk hari ini		
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", today.ToString()));
                filters.Add(new QueryFilter("JobTitle", JobTitle));
                filters.Add(new QueryFilter("GroupName", groupname));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string mppList = _mppAppService.Find(filters);
                var employeeIDs = mppList.DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID)).Select(x => x.EmployeeID).ToList();
                if (!string.IsNullOrEmpty(empID))
                    employeeIDs.Add(empID);

                var excludeUser = UserList.Where(x => employeeIDs.Contains(x.EmployeeID)).ToList();
                UserList = UserList.Except(excludeUser).ToList();

                // exclude user yg cuti hari ini
                string emLeaveList = _employeeLeaveAppService.GetAll();
                List<EmployeeLeaveModel> emLeaveModelList = emLeaveList.DeserializeToEmployeeLeaveList();
                emLeaveModelList = emLeaveModelList.Where(x => x.StartDate.Value.Date >= today.Date && x.EndDate.Value.Date <= today.Date).ToList();

                employeeIDs = emLeaveModelList.Select(x => x.EmployeeID).Distinct().ToList();

                excludeUser = UserList.Where(x => employeeIDs.Contains(x.EmployeeID)).ToList();
                UserList = UserList.Except(excludeUser).ToList();

                //jika shift 1, exclude yg kemarin kerja di shift 3 (kan gak mungkin kerja nerus)
                //jika shift 3, exclude yg besok kerja di shift 1
                if (shift.Trim() == "1" || shift.Trim() == "3")
                {
                    filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    if (shift.Trim() == "1")
                    {
                        filters.Add(new QueryFilter("Shift", "3"));
                        filters.Add(new QueryFilter("Date", yesterday.ToString()));
                    }
                    else if (shift.Trim() == "3")
                    {
                        filters.Add(new QueryFilter("Shift", "1"));
                        filters.Add(new QueryFilter("Date", tomorrow.ToString()));
                    }

                    mppList = _mppAppService.Find(filters);
                    employeeIDs = mppList.DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID)).Select(x => x.EmployeeID).ToList();

                    excludeUser = UserList.Where(x => employeeIDs.Contains(x.EmployeeID)).ToList();
                    UserList = UserList.Except(excludeUser).ToList();
                }

                UserList = ValidateUserByMachineCode(UserList, empSkillList, userMachineList, machineModel);
                UserList = UserList.OrderBy(x => x.FullName).ToList();

                model.Add(UserList.Select(x => new SelectListItem { Text = x.FullName, Value = x.EmployeeID }).ToList());
            }
            else
            {
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("JobTitle", JobTitle));
                filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                filters.Add(new QueryFilter("StatusMPP", "Normal"));
                filters.Add(new QueryFilter("GroupName", groupname));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<MppModel> mppLS1List = _mppAppService.Find(filters).DeserializeToMppList();
                mppLS1List = mppLS1List.GroupBy(p => p.EmployeeID).Select(g => g.First()).ToList();

                mppLS1List = mppLS1List.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

                // exclude user working today similar shift
                var tempList = mppLS1List.Where(x => x.Date.Date == today.Date && x.Shift.Trim() == shift.Trim()).ToList();
                mppLS1List = mppLS1List.Except(tempList).ToList();

                // exclude user yg cuti hari ini
                string emLeaveList = _employeeLeaveAppService.GetAll();
                List<EmployeeLeaveModel> emLeaveModelList = emLeaveList.DeserializeToEmployeeLeaveList();
                emLeaveModelList = emLeaveModelList.Where(x => x.StartDate.Value.Date >= today.Date && x.EndDate.Value.Date <= today.Date).ToList();

                var employeeIDs = emLeaveModelList.Select(x => x.EmployeeID).Distinct().ToList();

                var excludeUser = mppLS1List.Where(x => employeeIDs.Contains(x.EmployeeID)).ToList();
                mppLS1List = mppLS1List.Except(excludeUser).ToList();

                // exclude current user
                if (!string.IsNullOrEmpty(empID))
                    mppLS1List = mppLS1List.Where(x => x.EmployeeID != empID).ToList();

                mppLS1List = mppLS1List.Where(x => x.Shift != null && x.Shift.Trim() != shift.Trim()).OrderBy(x => x.EmployeeName).ToList();

                model.Add(mppLS1List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());

                model.Add(mppLS1List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());

                #region ::Helmy::
                //if (shift.Trim() == "1")
                //{
                //	// Ambil LS1 dari Shift 3 kemaren status normal
                //	ICollection<QueryFilter> filters = new List<QueryFilter>();
                //	filters.Add(new QueryFilter("Date", yesterday.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("Shift", "3"));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS1List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID) && x.EmployeeID != empID).ToList();
                //	mppLS1List = mppLS1List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS1List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());

                //	// Ambil LS2 dari shift 2 today status normal
                //	filters.Add(new QueryFilter("Date", today.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("Shift", "2"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS2List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                //	mppLS2List = mppLS2List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS2List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());
                //}
                //else if (shift.Trim() == "2")
                //{
                //	// Ambil LS1 dari Shift 1 today status normal
                //	ICollection<QueryFilter> filters = new List<QueryFilter>();
                //	filters.Add(new QueryFilter("Date", today.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("Shift", "1"));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS1List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID) && x.EmployeeID != empID).ToList();
                //	mppLS1List = mppLS1List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS1List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());

                //	// Ambil LS2 dari shift 3 today status normal
                //	filters.Add(new QueryFilter("Date", today.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("Shift", "3"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS2List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                //	mppLS2List = mppLS2List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS2List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());
                //}
                //else
                //{
                //	// Ambil LS1 dari Shift 2 today status normal
                //	ICollection<QueryFilter> filters = new List<QueryFilter>();
                //	filters.Add(new QueryFilter("Date", today.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("Shift", "2"));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS1List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID) && x.EmployeeID != empID).ToList();
                //	mppLS1List = mppLS1List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS1List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());

                //	// Ambil LS2 dari shift 1 besok status normal
                //	filters.Add(new QueryFilter("Date", tomorrow.ToString()));
                //	filters.Add(new QueryFilter("JobTitle", JobTitle));
                //	filters.Add(new QueryFilter("EmployeeMachine", machineCode));
                //	filters.Add(new QueryFilter("StatusMPP", "Normal"));
                //	filters.Add(new QueryFilter("Shift", "1"));
                //	filters.Add(new QueryFilter("IsDeleted", "0"));
                //	List<MppModel> mppLS2List = _mppAppService.Find(filters).DeserializeToMppList().Where(x => locIDList.Any(y => y == x.LocationID)).ToList();
                //	mppLS2List = mppLS2List.OrderBy(x => x.EmployeeName).ToList();

                //	model.Add(mppLS2List.Select(x => new SelectListItem { Text = x.EmployeeName, Value = x.ID.ToString() }).ToList());
                //}
                #endregion
            }

            return model;
        }

        [HttpPost]
        public ActionResult AddUserFromStatus(MppModel model)
        {
            try
            {
                model.StartDate = model.Date;
                model.EndDate = model.Date;
                model.Year = model.Date.Year;
                model.Week = GetCurrentWeekNumber(model.Date);
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                model.Remark = "Added from MPP Status";

                _mppAppService.Add(JsonHelper<MppModel>.Serialize(model));

                return Json(new { Status = "Success", Text = "Add Success", Type = "success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "Failed", Text = "Add Failed", Type = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult AddUserFromStatus(string empID, string date, long locID, string shift)
        {
            UserModel user = _userAppService.GetBy("EmployeeID", empID).DeserializeToUser();
            EmployeeModel empModel = _empAppService.GetBy("EmployeeID", empID, true).DeserializeToEmployee();

            var machineList = GetMachineList(user);
            ViewBag.MachineList = machineList;
            ViewBag.JobTitles = GetRoleList(user);
            ViewBag.StatusMppList = DropDownHelper.BindDropDownSimpleStatusMpp();
            ViewBag.Shift = DropDownHelper.BindDropDownShift();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
            ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);

            DateTime mppDate = DateTime.Parse(date);
            var model = new MppModel();
            model.EmployeeID = empID;
            model.EmployeeName = empModel.FullName;
            model.EmployeeMachine = machineList.Count > 0 ? machineList[0].Value : null;
            model.Date = mppDate;
            model.Location = _locationAppService.GetLocationFullCode(locID);
            model.LocationID = locID;
            model.Shift = shift == "-" ? "1" : shift;
            model.GroupType = empModel.GroupType == null ? "4G" : empModel.GroupType;
            model.GroupName = GetGroupName(locID, model.Shift, mppDate);

            return PartialView(model);
        }

        private List<SelectListItem> GetRoleList(UserModel user)
        {
            string userRole = _jobTitleAppService.GetRoleNameByJobTitleId(user.JobTitleID);
            List<UserRoleModel> userRoles = _userRoleAppService.FindBy("UserID", user.ID).DeserializeToUserRoleList();
            List<string> roleList = userRoles.OrderBy(x => x.RoleName).Select(x => x.RoleName).ToList();
            roleList.Add(userRole);

            List<SelectListItem> _menuList = new List<SelectListItem>();
            foreach (var item in roleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item,
                    Value = item
                });
            }

            return _menuList;
        }

        private List<SelectListItem> GetMachineList(UserModel user)
        {
            List<UserMachineTypeModel> empSkillList = _employeeSkillAppService.GetAll(true).DeserializeToUserMachineTypeList();
            List<UserMachineModel> userMachineList = _userMachineAppService.GetAll(true).DeserializeToUserMachineList();
            List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList();

            string location = _locationAppService.GetLocationFullCode(user.LocationID.Value);
            long pcID = _locationAppService.GetLocationID(location.Substring(0, 5));
            List<long> locIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");

            machineList = machineList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

            var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
            var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();

            var machineCodeList1 = machineList.Where(x => uskills.Any(y => y.MachineTypeID == x.MachineTypeID)).ToList();
            var machineCodeList2 = machineList.Where(x => umachines.Any(y => y.MachineID == x.ID)).ToList();

            List<MachineModel> machineUserList = new List<MachineModel>();

            if (uskills.Count == 0 && umachines.Count == 0)
            {
                machineUserList = machineList;
            }
            else
            {
                machineUserList.AddRange(machineCodeList1);
                machineUserList.AddRange(machineCodeList2);
            }

            List<string> machineCodeList = machineUserList.OrderBy(x => x.Code).Select(x => x.Code).Distinct().ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();
            foreach (var item in machineCodeList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item,
                    Value = item
                });
            }

            return _menuList;
        }

        public ActionResult AddUser(string machineCode, string date, long locID, string locType, bool isNextDay = false, string shift = "1")
        {
            var model = new MppModel();

            string machineAllocations;
            List<MachineAllocationModel> machineAllocationList;
            if (machineCode != "")
            {
                machineAllocations = _machineAllocationAppService.FindBy("MachineCode", machineCode, true);
                machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();
            }
            else
            {
                machineAllocations = _machineAllocationAppService.GetAll();
                machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();
            }
            var JobTitles = machineAllocationList.Select(x => x.MachineCategory).Distinct().ToList();

            ViewBag.JobTitles = JobTitles.Select(x => new SelectListItem { Text = x, Value = x }).ToList();
            ViewBag.StatusMppList = DropDownHelper.BindDropDownSimpleStatusMpp();
            ViewBag.Shift = DropDownHelper.BindDropDownShift();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
            ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);

            var dateFL = DateTime.Parse(date);
            if (isNextDay)
                dateFL = dateFL.AddDays(1);
            model.EmployeeMachine = machineCode;
            model.Date = dateFL;
            model.LocID = locID;
            model.LocType = locType;
            model.LocationID = locID;
            model.GroupName = GetGroupName(locID, shift, dateFL);

            ViewBag.UserList = new List<SelectListItem>();
            if (JobTitles.Count() > 0)
                ViewBag.UserList = GetUserWithCondition(machineCode, locID, locType, dateFL.ToString(), shift, JobTitles[0], ViewBag.StatusMppList[0].Value)[0];

            return PartialView(model);
        }

        [HttpPost]
        public ActionResult AddUser(MppModel model)
        {
            try
            {
                model.StartDate = model.Date;
                model.EndDate = model.Date;
                model.Year = model.Date.Year;
                model.Week = GetCurrentWeekNumber(model.Date);
                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;
                model.Location = _locationAppService.GetLocationFullCode(model.LocationID);
                model.GroupName = GetGroupName(model.LocationID, model.Shift, model.Date);

                if (model.StatusMPP == "Normal")
                {
                    string emp = _empAppService.GetBy("EmployeeID", model.EmployeeID, true);
                    EmployeeModel empModel = emp.DeserializeToEmployee();

                    model.EmployeeName = empModel == null ? string.Empty : empModel.FullName;

                    _mppAppService.Add(JsonHelper<MppModel>.Serialize(model));
                }
                else
                {
                    if (model.EmployeeIDLS1 != "0")
                    {
                        model.StatusMPP = "LongShift1";
                        model.EmployeeID = model.EmployeeIDLS1;

                        string emp = _empAppService.GetBy("EmployeeID", model.EmployeeIDLS1, true);
                        EmployeeModel empModel = emp.DeserializeToEmployee();

                        model.EmployeeName = empModel == null ? string.Empty : empModel.FullName;

                        var lala = _mppAppService.Add(JsonHelper<MppModel>.Serialize(model));
                    }

                    if (model.EmployeeIDLS2 != "0")
                    {
                        model.StatusMPP = "LongShift2";
                        model.EmployeeID = model.EmployeeIDLS2;

                        var emp = _empAppService.GetBy("EmployeeID", model.EmployeeIDLS2, true);
                        var empModel = emp.DeserializeToEmployee();

                        model.EmployeeName = empModel == null ? string.Empty : empModel.FullName;

                        var lala = _mppAppService.Add(JsonHelper<MppModel>.Serialize(model));
                    }
                }

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.CreateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Detail", "MPP", new
            {
                machineCode = model.EmployeeMachine,
                date = model.Date.ToString("dd-MMM-yyyy"),
                locID = model.LocID,
                locType = model.LocType,
                shift = model.Shift.Trim()
            });
        }

        [HttpPost]
        public object GetUserList(string machineCode, long locID, string locType, string date, string shift, string JobTitle, string StatusMpp)
        {
            var SelectUsers = GetUserWithCondition(machineCode, locID, locType, date, shift, JobTitle, StatusMpp);
            return Json(SelectUsers, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public object GetUserListOnEdit(long mppID, string status)
        {
            MppModel mpp = _mppAppService.GetById(mppID).DeserializeToMpp();
            string pcLocation = mpp.Location.Substring(0, 5);
            long pcID = _locationAppService.GetLocationID(pcLocation);

            var SelectUsers = GetUserWithCondition(mpp.EmployeeMachine, pcID, "productioncenter", mpp.Date.ToString(), mpp.Shift, mpp.JobTitle, status, mpp.EmployeeID);

            return Json(SelectUsers, JsonRequestBehavior.AllowGet);
        }

        public ActionResult RemoveUser(string machineCode, string date, long locID, bool isNextDay = false)
        {
            DateTime dateFL = DateTime.Parse(date);
            if (isNextDay)
                dateFL = dateFL.AddDays(1);

            List<QueryFilter> filter = new List<QueryFilter>();
            filter.Add(new QueryFilter("LocationID", locID.ToString()));
            filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd")));

            string mpps = _mppAppService.Find(filter);
            List<MppModel> mppList = mpps.DeserializeToMppList();
            mppList = mppList.Where(x => x.EmployeeMachine.Contains(machineCode)).OrderBy(x => x.JobTitle).ToList();

            string machineAllocations = _machineAllocationAppService.FindBy("MachineCode", machineCode);
            List<MachineAllocationModel> machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();

            MppAllocationModel model = new MppAllocationModel();
            model.AllocationList = new List<MachineAllocationModel>();
            model.MppList = new List<MppModel>();

            if (mppList.Count > machineAllocationList.Count)
            {
                int x = 0;
                foreach (var item in mppList)
                {
                    model.MppList.Add(item);
                    if (x < machineAllocationList.Count)
                    {
                        model.AllocationList.Add(machineAllocationList[x++]);
                    }
                    else
                    {
                        model.AllocationList.Add(new MachineAllocationModel());
                    }
                }
            }
            else
            {
                int x = 0;
                foreach (var item in machineAllocationList)
                {
                    model.AllocationList.Add(item);
                    if (x < mppList.Count)
                    {
                        model.MppList.Add(mppList[x++]);
                    }
                    else
                    {
                        model.MppList.Add(new MppModel());
                    }
                }
            }

            ViewBag.MachineCode = machineCode;
            var temp = model.AllocationList.Where(x => x.MachineCategory == null).ToList();
            model.AllocationList = model.AllocationList.Where(x => x.MachineCategory != null).OrderBy(x => x.MachineCategory).ToList();
            if (temp.Count > 0)
                model.AllocationList.AddRange(temp);
            model.MppList = model.MppList.OrderBy(x => x.JobTitle).ToList();

            return PartialView(model);
        }

        public ActionResult Dashboard()
        {
            return View();
        }

        public ActionResult Edit(int id, long locID, string locType)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BuildEmptyList();
                ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
                ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
                ViewBag.StatusMppList = DropDownHelper.BindDropDownStatusMpp();
                ViewBag.UserLS1List = DropDownHelper.BuildEmptyList();
                ViewBag.UserLS2List = DropDownHelper.BuildEmptyList();

                MppModel model = GetMpp(id);
                model.Access = GetAccess(WebConstants.MenuSlug.MPP, _menuService);
                model.LocID = locID;
                model.LocType = locType;

                if (!string.IsNullOrEmpty(model.EmployeeMachine))
                {
                    string pcLocation = model.Location.Substring(0, 5);
                    long pcID = _locationAppService.GetLocationID(pcLocation);

                    var SelectUsers = GetUserWithCondition(model.EmployeeMachine, pcID, "productioncenter", model.Date.ToString(), model.Shift, model.JobTitle, "Normal", model.EmployeeID);

                    ViewBag.UserList = SelectUsers[0];
                }
                else
                {
                    ViewBag.UserList = DropDownHelper.BuildEmptyList();
                }

                return PartialView(model);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult Edit(MppModel model)
        {
            try
            {
                MppModel modelOld = GetMpp(Convert.ToInt32(model.ID));

                if (model.StatusMPP == "")
                    model.StatusMPP = "Normal";

                if (model.StatusMPP == "Normal")
                {
                    MppChangesModel changesModel = new MppChangesModel();
                    changesModel.MPPID = model.ID;
                    changesModel.FieldName = "EmployeeID";
                    changesModel.OldValue = modelOld.EmployeeID;
                    changesModel.NewValue = model.IDNormal;
                    changesModel.DataType = "NonNumeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;

                    _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(changesModel));

                    // get the new emp name 
                    string emp = _empAppService.GetBy("EmployeeID", model.IDNormal, true);
                    EmployeeModel empModel = emp.DeserializeToEmployee();

                    modelOld.ModifiedBy = AccountName;
                    modelOld.ModifiedDate = DateTime.Now;
                    modelOld.EmployeeID = model.IDNormal;
                    modelOld.EmployeeName = empModel.FullName;
                    modelOld.StatusMPP = model.StatusMPP;

                    _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelOld));
                }
                else
                {
                    if (model.IDLS1 > 0)
                    {
                        MppModel LS1 = _mppAppService.GetById(model.IDLS1).DeserializeToMpp();

                        MppChangesModel changesModel = new MppChangesModel();
                        changesModel.MPPID = model.ID;
                        changesModel.FieldName = "EmployeeID";
                        changesModel.OldValue = modelOld.EmployeeID;
                        changesModel.NewValue = LS1.EmployeeID;
                        changesModel.DataType = "NonNumeric";
                        changesModel.ModifiedBy = AccountName;
                        changesModel.ModifiedDate = DateTime.Now;

                        _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(changesModel));

                        // get the new emp and status 
                        modelOld.ModifiedBy = AccountName;
                        modelOld.ModifiedDate = DateTime.Now;
                        modelOld.EmployeeID = LS1.EmployeeID;
                        modelOld.EmployeeName = LS1.EmployeeName;
                        modelOld.StatusMPP = "LongShift1";

                        _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelOld));
                    }

                    if (model.IDLS2 > 0)
                    {
                        MppModel LS2 = _mppAppService.GetById(model.IDLS2).DeserializeToMpp();

                        MppModel newMpp = new MppModel();
                        newMpp.GroupName = model.GroupName;
                        newMpp.GroupType = model.GroupType;
                        newMpp.Year = model.Date.Year;
                        newMpp.Week = GetCurrentWeekNumber(model.Date);
                        newMpp.StartDate = model.Date;
                        newMpp.EndDate = model.Date;
                        newMpp.Date = model.Date;
                        newMpp.LocationID = model.LocationID;
                        newMpp.Location = _locationAppService.GetLocationFullCode(model.LocationID);
                        newMpp.ModifiedDate = DateTime.Now;
                        newMpp.ModifiedBy = AccountName;
                        newMpp.JobTitle = model.JobTitle;
                        newMpp.EmployeeID = LS2.EmployeeID;
                        newMpp.EmployeeName = LS2.EmployeeName;
                        newMpp.EmployeeMachine = LS2.EmployeeMachine;
                        newMpp.StatusMPP = "LongShift2";
                        newMpp.Shift = model.Shift;

                        _mppAppService.Add(JsonHelper<MppModel>.Serialize(newMpp));
                    }
                }

                SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Detail", "MPP", new
            {
                machineCode = model.EmployeeMachine,
                date = model.Date.ToString("dd-MMM-yyyy"),
                locID = model.LocID,
                locType = model.LocType,
                shift = model.Shift.Trim()
            });
        }

        public ActionResult EditFromStatus(int id)
        {
            try
            {
                ViewBag.MachineList = DropDownHelper.BuildEmptyList();
                ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
                ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
                ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
                ViewBag.StatusMppList = DropDownHelper.BindDropDownStatusMpp();
                ViewBag.UserLS1List = DropDownHelper.BuildEmptyList();
                ViewBag.UserLS2List = DropDownHelper.BuildEmptyList();

                MppModel model = GetMpp(id);
                model.Access = GetAccess(WebConstants.MenuSlug.MPP, _menuService);

                if (!string.IsNullOrEmpty(model.EmployeeMachine))
                {
                    string pcLocation = model.Location.Substring(0, 5);
                    long pcID = _locationAppService.GetLocationID(pcLocation);

                    var SelectUsers = GetUserWithCondition(model.EmployeeMachine, pcID, "productioncenter", model.Date.ToString(), model.Shift, model.JobTitle, "Normal", model.EmployeeID);

                    ViewBag.UserList = SelectUsers[0];
                }
                else
                {
                    ViewBag.UserList = DropDownHelper.BuildEmptyList();
                }

                return PartialView(model);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult EditFromStatus(MppModel model)
        {
            try
            {
                MppModel modelOld = GetMpp(Convert.ToInt32(model.ID));

                if (model.StatusMPP == "")
                    model.StatusMPP = "Normal";

                if (model.StatusMPP == "Normal")
                {
                    MppChangesModel changesModel = new MppChangesModel();
                    changesModel.MPPID = model.ID;
                    changesModel.FieldName = "EmployeeID";
                    changesModel.OldValue = modelOld.EmployeeID;
                    changesModel.NewValue = model.IDNormal;
                    changesModel.DataType = "NonNumeric";
                    changesModel.ModifiedBy = AccountName;
                    changesModel.ModifiedDate = DateTime.Now;

                    _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(changesModel));

                    // get the new emp name 
                    string emp = _empAppService.GetBy("EmployeeID", model.IDNormal, true);
                    EmployeeModel empModel = emp.DeserializeToEmployee();

                    modelOld.ModifiedBy = AccountName;
                    modelOld.ModifiedDate = DateTime.Now;
                    modelOld.EmployeeID = model.IDNormal;
                    modelOld.EmployeeName = empModel.FullName;
                    modelOld.StatusMPP = model.StatusMPP;

                    _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelOld));
                }
                else
                {
                    if (model.IDLS1 > 0)
                    {
                        MppModel LS1 = _mppAppService.GetById(model.IDLS1).DeserializeToMpp();

                        MppChangesModel changesModel = new MppChangesModel();
                        changesModel.MPPID = model.ID;
                        changesModel.FieldName = "EmployeeID";
                        changesModel.OldValue = modelOld.EmployeeID;
                        changesModel.NewValue = LS1.EmployeeID;
                        changesModel.DataType = "NonNumeric";
                        changesModel.ModifiedBy = AccountName;
                        changesModel.ModifiedDate = DateTime.Now;

                        _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(changesModel));

                        // get the new emp and status 
                        modelOld.ModifiedBy = AccountName;
                        modelOld.ModifiedDate = DateTime.Now;
                        modelOld.EmployeeID = LS1.EmployeeID;
                        modelOld.EmployeeName = LS1.EmployeeName;
                        modelOld.StatusMPP = "LongShift1";

                        _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelOld));
                    }

                    if (model.IDLS2 > 0)
                    {
                        MppModel LS2 = _mppAppService.GetById(model.IDLS2).DeserializeToMpp();

                        MppModel newMpp = new MppModel();
                        newMpp.GroupName = model.GroupName;
                        newMpp.GroupType = model.GroupType;
                        newMpp.Year = model.Date.Year;
                        newMpp.Week = GetCurrentWeekNumber(model.Date);
                        newMpp.StartDate = model.Date;
                        newMpp.EndDate = model.Date;
                        newMpp.Date = model.Date;
                        newMpp.LocationID = model.LocationID;
                        newMpp.Location = _locationAppService.GetLocationFullCode(model.LocationID);
                        newMpp.ModifiedDate = DateTime.Now;
                        newMpp.ModifiedBy = AccountName;
                        newMpp.JobTitle = model.JobTitle;
                        newMpp.EmployeeID = LS2.EmployeeID;
                        newMpp.EmployeeName = LS2.EmployeeName;
                        newMpp.EmployeeMachine = LS2.EmployeeMachine;
                        newMpp.StatusMPP = "LongShift2";
                        newMpp.Shift = model.Shift;

                        _mppAppService.Add(JsonHelper<MppModel>.Serialize(newMpp));
                    }
                }

                return Json(new { Status = "Success", Text = "Update Success", Type = "success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "Failed", Text = "Update Failed", Type = "error" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult EditLeave(int id)
        {
            ViewBag.MachineList = DropDownHelper.BuildEmptyList();
            ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.GroupTypeList = DropDownHelper.BuildEmptyList();
            ViewBag.StatusMppList = DropDownHelper.BindDropDownStatusMpp();
            ViewBag.UserList = DropDownHelper.BuildEmptyList(); ;

            MppModel model = GetMpp(id);
            model.Access = GetAccess(WebConstants.MenuSlug.MPP, _menuService);
            model.ProdCenterID = AccountProdCenterID;
            model.DepartmentID = AccountDepartmentID;
            model.SubDepartmentID = AccountLocationID;

            if (!string.IsNullOrEmpty(model.EmployeeMachine) && model.LocationID != 0)
            {
                if (model.Shift.Trim() == "1")
                {
                    ViewBag.UserLS1List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "2");
                    ViewBag.UserLS2List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "3");
                }
                else if (model.Shift.Trim() == "2")
                {
                    ViewBag.UserLS1List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "1");
                    ViewBag.UserLS2List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "3");
                }
                else if (model.Shift.Trim() == "3")
                {
                    ViewBag.UserLS1List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "1");
                    ViewBag.UserLS2List = GetUserListByFilterAndShift(id, model.LocationID, model.EmployeeMachine, "2");
                }
                else
                {
                    ViewBag.UserLS1List = DropDownHelper.BuildEmptyList();
                    ViewBag.UserLS2List = DropDownHelper.BuildEmptyList();
                }
            }
            else
            {
                ViewBag.UserLS1List = DropDownHelper.BuildEmptyList();
                ViewBag.UserLS2List = DropDownHelper.BuildEmptyList();
            }

            return PartialView(model);
        }

        [HttpPost]
        public ActionResult EditLeave(MppModel model)
        {
            try
            {
                if (model.IDLS1 == 0 || model.IDLS2 == 0)
                {
                    SetFalseTempData("The replacement for LongShift are missing");
                    return RedirectToAction("Index");
                }

                MppModel modelOld = GetMpp(Convert.ToInt32(model.ID));
                modelOld.IsDeleted = true;
                modelOld.ModifiedBy = AccountName;
                modelOld.ModifiedDate = DateTime.Now;

                _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelOld));

                MppModel modelLS1 = GetMpp(Convert.ToInt32(model.IDLS1));
                modelLS1.StatusMPP = "LongShift1";
                modelLS1.ModifiedBy = AccountName;
                modelLS1.ModifiedDate = DateTime.Now;

                _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelLS1));

                MppModel modelLS2 = GetMpp(Convert.ToInt32(model.IDLS2));
                modelLS2.StatusMPP = "LongShift2";
                modelLS2.ModifiedBy = AccountName;
                modelLS2.ModifiedDate = DateTime.Now;

                _mppAppService.Update(JsonHelper<MppModel>.Serialize(modelLS2));

                MppChangesModel changesModel = new MppChangesModel();
                changesModel.MPPID = model.ID;
                changesModel.FieldName = "IsDeleted";
                changesModel.OldValue = modelOld.IsDeleted.ToString();
                changesModel.NewValue = "true";
                changesModel.DataType = "NonNumeric";
                changesModel.ModifiedBy = AccountName;
                changesModel.ModifiedDate = DateTime.Now;

                _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(changesModel));

                MppChangesModel modelChangesLS1 = new MppChangesModel();
                modelChangesLS1.MPPID = model.IDLS1;
                modelChangesLS1.FieldName = "StatusMPP";
                modelChangesLS1.OldValue = "Normal";
                modelChangesLS1.NewValue = "LongShift1";
                modelChangesLS1.DataType = "NonNumeric";
                modelChangesLS1.ModifiedBy = AccountName;
                modelChangesLS1.ModifiedDate = DateTime.Now;

                _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(modelChangesLS1));

                MppChangesModel modelChangesLS2 = new MppChangesModel();
                modelChangesLS2.MPPID = model.IDLS2;
                modelChangesLS2.FieldName = "StatusMPP";
                modelChangesLS2.OldValue = "Normal";
                modelChangesLS2.NewValue = "LongShift2";
                modelChangesLS2.DataType = "NonNumeric";
                modelChangesLS2.ModifiedBy = AccountName;
                modelChangesLS2.ModifiedDate = DateTime.Now;

                _mppChangesAppService.Add(JsonHelper<MppChangesModel>.Serialize(modelChangesLS2));

                SetTrueTempData(UIResources.UpdateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult Delete(int id)
        {
            try
            {
                _mppAppService.Remove(id);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllMachineList(string dateFilter, long locID, string locType, long pcID)
        {
            DateTime dateFL = DateTime.Parse(dateFilter);
            DateTime nextDateFL = dateFL.AddDays(1);

            List<long> locationIDList = _locationAppService.GetLocIDListByLocType(locID, locType);

            List<QueryFilter> filter = new List<QueryFilter>();
            filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
            filter.Add(new QueryFilter("Date", nextDateFL.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));

            string mppList = _mppAppService.Find(filter);
            List<MppModel> mppModelList = mppList.DeserializeToMppList();
            mppModelList = mppModelList.Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();

            List<MppModel> mppModelList1 = mppModelList.Where(x => x.Date.Date == dateFL.Date).ToList();
            List<MppModel> mppModelList2 = mppModelList.Where(x => x.Date.Date == dateFL.AddDays(1).Date).ToList();

            filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
            filter.Add(new QueryFilter("Date", nextDateFL.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));

            string wpps = _wppStpAppService.Find(filter);
            List<WppStpModel> wppList1 = wpps.DeserializeToWppStpList().Where(x => x.Date.Date == dateFL.Date && x.LocationID == pcID).ToList();
            List<WppStpModel> wppList2 = wpps.DeserializeToWppStpList().Where(x => x.Date.Date == dateFL.AddDays(1).Date && x.LocationID == pcID).ToList();

            List<MachineModel> machineList1 = new List<MachineModel>();
            List<MachineModel> machineList2 = new List<MachineModel>();

            string machines = _machineAppService.GetAll(true);
            List<MachineModel> machineList = machines.DeserializeToMachineList();

            //Machine code tidak ada di WPP today, isi semua mesin tapi warna grey
            if (wppList1.Count == 0)
            {
                machineList1 = machines.DeserializeToMachineList().Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();
                machineList1.ForEach(x => x.IsSelectedDate = true);
            }
            else
            {
                List<string> wppMachineCodeList = wppList1.Where(x => !string.IsNullOrEmpty(x.Packer)).Select(x => x.Packer).Distinct().ToList();
                var temp = wppList1.Where(x => !string.IsNullOrEmpty(x.Maker)).Select(x => x.Maker).Distinct().ToList();
                if (temp.Count > 0)
                    wppMachineCodeList.AddRange(temp);

                foreach (var item in wppMachineCodeList)
                {
                    var wpp = wppList1.Where(x => x.Packer.Trim() == item.Trim() || x.Maker.Trim() == item.Trim()).First();

                    var machine = machineList.Where(x => x.Code.ToLower() == item.ToLower()).FirstOrDefault();
                    if (machine != null)
                    {
                        var tempMachine = machineList1.Where(x => x.Code == machine.Code).FirstOrDefault();
                        if (tempMachine != null)
                        {
                            tempMachine.IsExistInWpp = true;
                            tempMachine.IsShift1Zero = wpp.Shift1 == 0;
                            tempMachine.IsShift2Zero = wpp.Shift2 == 0;
                            tempMachine.IsShift3Zero = wpp.Shift3 == 0;
                        }
                        else
                        {
                            //machine red
                            machine.IsExistInWpp = true;
                            machine.IsShift1Zero = wpp.Shift1 == 0;
                            machine.IsShift2Zero = wpp.Shift2 == 0;
                            machine.IsShift3Zero = wpp.Shift3 == 0;
                            machineList1.Add(machine);
                        }
                    }
                }

                machineList1.ForEach(x => x.IsSelectedDate = true);
            }

            //Machine code tidak ada di WPP today+1, isi semua mesin tapi warna grey
            if (wppList2.Count == 0)
            {
                machineList2 = machines.DeserializeToMachineList().Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();
            }
            else
            {
                List<string> wppMachineCodeList = wppList2.Where(x => !string.IsNullOrEmpty(x.Packer)).Select(x => x.Packer).Distinct().ToList();
                var temp = wppList2.Where(x => !string.IsNullOrEmpty(x.Maker)).Select(x => x.Maker).Distinct().ToList();
                if (temp.Count > 0)
                    wppMachineCodeList.AddRange(temp);

                // reinstante
                machineList = machines.DeserializeToMachineList();

                foreach (var item in wppMachineCodeList)
                {
                    var wpp = wppList2.Where(x => x.Packer.Trim() == item.Trim() || x.Maker.Trim() == item.Trim()).First();

                    var machine = machineList.Where(x => x.Code.ToLower() == item.ToLower()).FirstOrDefault();
                    if (machine != null)
                    {
                        var tempMachine = machineList2.Where(x => x.Code == machine.Code).FirstOrDefault();
                        if (tempMachine != null)
                        {
                            tempMachine.IsExistInWpp = true;
                            tempMachine.IsShift1Zero = wpp.Shift1 == 0;
                            tempMachine.IsShift2Zero = wpp.Shift2 == 0;
                            tempMachine.IsShift3Zero = wpp.Shift3 == 0;
                        }
                        else
                        {
                            //machine red
                            machine.IsExistInWpp = true;
                            machine.IsShift1Zero = wpp.Shift1 == 0;
                            machine.IsShift2Zero = wpp.Shift2 == 0;
                            machine.IsShift3Zero = wpp.Shift3 == 0;
                            machineList2.Add(machine);
                        }
                    }
                }
            }

            string machineAllocations = _machineAllocationAppService.GetAll();
            List<MachineAllocationModel> machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();

            foreach (var item in machineList1)
            {
                var allocations = machineAllocationList.Where(x => x.MachineCode == item.Code).ToList();

                bool isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList1.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "1").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift1FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;

                isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList1.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "2").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift2FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;

                isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList1.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "3").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift3FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;
            }

            List<MachineModel> machineSPList1 = machineList1.Where(x => x.Location != null && x.Location.Contains("SP")).ToList();
            machineSPList1.ForEach(x => x.IsSP = true);
            List<MachineModel> machinePPList1 = machineList1.Where(x => x.Location != null && x.Location.Contains("PP")).ToList();
            machinePPList1.ForEach(x => x.IsPP = true);

            foreach (var item in machineList2)
            {
                var allocations = machineAllocationList.Where(x => x.MachineCode == item.Code).ToList();

                bool isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList2.Where(x => x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "1").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift1FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;

                isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList2.Where(x => x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "2").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift2FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;

                isFullyAssigned = true;
                foreach (var alloc in allocations)
                {
                    var tempList = mppModelList2.Where(x => x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "3").ToList();
                    double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                    double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                    if ((allocCount + allocCountLS) < alloc.Value)
                    {
                        isFullyAssigned = false;
                        break;
                    }
                }

                item.IsShift3FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;
            }

            List<MachineModel> machineSPList2 = machineList2.Where(x => x.Location != null && x.Location.Contains("SP")).ToList();
            machineSPList2.ForEach(x => x.IsSP = true);
            List<MachineModel> machinePPList2 = machineList2.Where(x => x.Location != null && x.Location.Contains("PP")).ToList();
            machinePPList2.ForEach(x => x.IsPP = true);

            List<MachineModel> result = new List<MachineModel>();
            if (machineSPList1.Count > 0)
                result.AddRange(machineSPList1);
            if (machinePPList1.Count > 0)
                result.AddRange(machinePPList1);
            if (machineSPList2.Count > 0)
                result.AddRange(machineSPList2);
            if (machinePPList2.Count > 0)
                result.AddRange(machinePPList2);

            double selectedDateShift1 = machineSPList1.Count(x => x.IsShift1FullyAssigned);
            double selectedDateShift2 = machineSPList1.Count(x => x.IsShift2FullyAssigned);
            double selectedDateShift3 = machineSPList1.Count(x => x.IsShift3FullyAssigned);

            var full = machineSPList1.Where(x => x.IsShift1FullyAssigned).FirstOrDefault();


            double nextDateShift1 = machineSPList2.Count(x => x.IsShift1FullyAssigned);
            double nextDateShift2 = machineSPList2.Count(x => x.IsShift2FullyAssigned);
            double nextDateShift3 = machineSPList2.Count(x => x.IsShift3FullyAssigned);

            double machine1Count = machineList1.Count();
            double machine2Count = machineList2.Count();

            double selectedDateShift1Percentage = selectedDateShift1 == 0 ? 0 : Math.Round(selectedDateShift1 / machine1Count, 2) * 100;
            double selectedDateShift2Percentage = selectedDateShift2 == 0 ? 0 : Math.Round(selectedDateShift2 / machine1Count, 2) * 100;
            double selectedDateShift3Percentage = selectedDateShift3 == 0 ? 0 : Math.Round(selectedDateShift3 / machine1Count, 2) * 100;
            double selectedDateTotalFullyAssigned = selectedDateShift1 + selectedDateShift1 + selectedDateShift1;
            double selectedDateOverall = selectedDateTotalFullyAssigned == 0 ? 0 : Math.Round(selectedDateTotalFullyAssigned / (machine1Count * 3), 2) * 100;

            double nextDateShift1Percentage = nextDateShift1 == 0 ? 0 : Math.Round(nextDateShift1 / machine2Count, 2) * 100;
            double nextDateShift2Percentage = nextDateShift2 == 0 ? 0 : Math.Round(nextDateShift2 / machine2Count, 2) * 100;
            double nextDateShift3Percentage = nextDateShift3 == 0 ? 0 : Math.Round(nextDateShift3 / machine2Count, 2) * 100;
            double nextDateTotalFullyAssigned = nextDateShift1 + nextDateShift2 + nextDateShift3;
            double nextDateOverall = nextDateTotalFullyAssigned == 0 ? 0 : Math.Round(nextDateTotalFullyAssigned / (machine2Count * 3), 2) * 100;

            MppMachineModel model = new MppMachineModel();
            model.MachineList = result;
            model.CurrentDateMachineCount = machineSPList1.Count;
            model.NextDateMachineCount = machineSPList2.Count;
            model.CurrentDate = dateFilter;
            model.NextDate = dateFL.AddDays(1).ToString("dd-MMM-yy");

            model.SelectedDateShift1Percentage = selectedDateShift1 == 0 ? "SHIFT 1 (0 %)" : "SHIFT 1 (" + selectedDateShift1Percentage.ToString() + " %)";
            model.SelectedDateShift2Percentage = selectedDateShift2 == 0 ? "SHIFT 2 (0 %)" : "SHIFT 2 (" + selectedDateShift2Percentage.ToString() + " %)";
            model.SelectedDateShift3Percentage = selectedDateShift3 == 0 ? "SHIFT 3 (0 %)" : "SHIFT 3 (" + selectedDateShift3Percentage.ToString() + " %)";
            model.SelectedDateOverallPercentage = selectedDateOverall == 0 ? "OVERALL (0 %)" : "OVERALL  (" + selectedDateOverall.ToString() + " %)";
            model.NextDateShift1Percentage = nextDateShift1 == 0 ? "SHIFT 1 (0 %)" : "SHIFT 1 (" + nextDateShift1Percentage.ToString() + " %)";
            model.NextDateShift2Percentage = nextDateShift2 == 0 ? "SHIFT 2 (0 %)" : "SHIFT 2 (" + nextDateShift2Percentage.ToString() + " %)";
            model.NextDateShift3Percentage = nextDateShift3 == 0 ? "SHIFT 3 (0 %)" : "SHIFT 3 (" + nextDateShift3Percentage.ToString() + " %)";
            model.NextDateOverallPercentage = nextDateOverall == 0 ? "OVERALL (0 %)" : "OVERALL (" + nextDateOverall.ToString() + " %)";


            return Json(model, JsonRequestBehavior.AllowGet);
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
                string mppList = _mppAppService.GetAll(true);
                List<MppModel> result = mppList.DeserializeToMppList().OrderBy(x => x.StartDate).ToList();

                int recordsTotal = result.Count();

                // Search     - Correction 231019
                if (!string.IsNullOrEmpty(searchValue))
                {
                    result = result.Where(m => (m.JobTitle != null ? m.JobTitle.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.EmployeeName != null ? m.EmployeeName.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.EmployeeMachine != null ? m.EmployeeMachine.ToLower().Contains(searchValue.ToLower()) : false)
                                                ).ToList();
                }

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
                {
                    if (sortColumnDir == "asc")
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "startdate":
                                result = result.OrderBy(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                result = result.OrderBy(x => x.EndDate).ToList();
                                break;
                            case "date":
                                result = result.OrderBy(x => x.Date).ToList();
                                break;
                            case "jobtitle":
                                result = result.OrderBy(x => x.JobTitle).ToList();
                                break;
                            case "employeename":
                                result = result.OrderBy(x => x.EmployeeName).ToList();
                                break;
                            case "employeemachine":
                                result = result.OrderBy(x => x.EmployeeMachine).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (sortColumn.ToLower())
                        {
                            case "startdate":
                                result = result.OrderByDescending(x => x.StartDate).ToList();
                                break;
                            case "enddate":
                                result = result.OrderByDescending(x => x.EndDate).ToList();
                                break;
                            case "date":
                                result = result.OrderByDescending(x => x.Date).ToList();
                                break;
                            case "jobtitle":
                                result = result.OrderByDescending(x => x.JobTitle).ToList();
                                break;
                            case "employeename":
                                result = result.OrderByDescending(x => x.EmployeeName).ToList();
                                break;
                            case "employeemachine":
                                result = result.OrderByDescending(x => x.EmployeeMachine).ToList();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // total number of rows count     
                int recordsFiltered = result.Count();

                // Paging     
                List<MppModel> data = result.Skip(skip).Take(pageSize).ToList();

                string holidays = _calendarHolidayAppService.GetAll(true);
                List<CalendarHolidayModel> holidayList = holidays.DeserializeToCalendarHolidayList().OrderBy(x => x.Date).ToList();

                Dictionary<long, List<long>> locationIDListMap = new Dictionary<long, List<long>>();
                Dictionary<long, string> locationMap = new Dictionary<long, string>();

                // check if it has wpp defined
                foreach (MppModel item in data)
                {
                    List<QueryFilter> filter = new List<QueryFilter>();
                    filter.Add(new QueryFilter("LocationID", item.LocationID.ToString()));
                    filter.Add(new QueryFilter("Date", item.Date.ToString("yyyy-MM-dd")));

                    string wpp = _wppStpAppService.Find(filter);
                    List<WppStpModel> wppList = wpp.DeserializeToWppStpList();

                    if (wppList.Count > 0)
                    {
                        item.IsWPPExist = true;
                    }

                    List<long> listOfID = new List<long>();
                    if (locationIDListMap.ContainsKey(item.LocationID))
                    {
                        locationIDListMap.TryGetValue(item.LocationID, out listOfID);
                    }
                    else
                    {
                        listOfID = GetParentIDList(item.LocationID);
                        locationIDListMap.Add(item.LocationID, listOfID);
                    }

                    if (holidayList.Any(x => x.Date == item.Date && listOfID.Any(y => y == x.LocationID)))
                    {
                        item.IsHoliday = true;
                    }

                    if (item.LocationID > 0)
                    {
                        if (locationMap.ContainsKey(item.LocationID))
                        {
                            string loc;
                            locationMap.TryGetValue(item.LocationID, out loc);
                            item.Location = loc;
                        }
                        else
                        {
                            item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
                            locationMap.Add(item.LocationID, item.Location);
                        }
                    }
                }

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                ViewBag.ErrorMessage = UIResources.LoadDataFailed;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<MppModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetAllStatus(string dateFilter, string endDateFilter, long locID, string locType, string shift, long groupTypeID, long pcID)
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
            DateTime dateFL = DateTime.Parse(dateFilter);

            List<MppModel> result = GetStatusData(dateFL, endDateFilter, locationIdList, shift, groupTypeID, locID);
            if (result.Count > 0)
            {
                result[0].SumAvailable = result.Where(x => x.StatusMPP == "Available").Count();
                result[0].SumHalfAssigned = result.Where(x => x.StatusMPP == "Half-Assigned").Count();
                result[0].SumFullyAssigned = result.Where(x => x.StatusMPP == "Fully-Assigned").Count();
                result[0].SumOverload = result.Where(x => x.StatusMPP == "Overload").Count();
            }

            int recordsTotal = result.Count();

            if (!string.IsNullOrEmpty(searchValue))
            {
                result = result.Where(m => (m.StatusMPP != null ? m.StatusMPP.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.EmployeeName != null ? m.EmployeeName.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.EmployeeID != null ? m.EmployeeID.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.EmployeeMachine != null ? m.EmployeeMachine.ToLower().Contains(searchValue.ToLower()) : false)
                                            ).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "date":
                            result = result.OrderBy(x => x.Date).ToList();
                            break;
                        case "shift":
                            result = result.OrderBy(x => x.Shift).ToList();
                            break;
                        case "location":
                            result = result.OrderBy(x => x.Location).ToList();
                            break;
                        case "employeeid":
                            result = result.OrderBy(x => x.EmployeeID).ToList();
                            break;
                        case "statusmpp":
                            result = result.OrderBy(x => x.StatusMPP).ToList();
                            break;
                        case "employeename":
                            result = result.OrderBy(x => x.EmployeeName).ToList();
                            break;
                        case "employeemachine":
                            result = result.OrderBy(x => x.EmployeeMachine).ToList();
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (sortColumn.ToLower())
                    {
                        case "date":
                            result = result.OrderByDescending(x => x.Date).ToList();
                            break;
                        case "shift":
                            result = result.OrderByDescending(x => x.Shift).ToList();
                            break;
                        case "location":
                            result = result.OrderByDescending(x => x.Location).ToList();
                            break;
                        case "employeeid":
                            result = result.OrderByDescending(x => x.EmployeeID).ToList();
                            break;
                        case "statusmpp":
                            result = result.OrderByDescending(x => x.StatusMPP).ToList();
                            break;
                        case "employeename":
                            result = result.OrderByDescending(x => x.EmployeeName).ToList();
                            break;
                        case "employeemachine":
                            result = result.OrderByDescending(x => x.EmployeeMachine).ToList();
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

        private List<MppModel> GetStatusData(DateTime dateFL, string endDateFilter, List<long> locationIdList, string shift, long groupTypeID, long locID)
        {
            // Getting all data			
            #region Get User List
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("GroupTypeID", groupTypeID));
            if (locID != 0)
                filters.Add(new QueryFilter("LocationID", locID));

            if (string.IsNullOrEmpty(endDateFilter))
            {
                filters.Add(new QueryFilter("Date", dateFL.Date.ToString()));
            }
            else
            {
                DateTime endDateFL = DateTime.Parse(endDateFilter);
                if (dateFL == endDateFL)
                {
                    endDateFilter = null;
                    filters.Add(new QueryFilter("Date", dateFL.Date.ToString()));
                }
                else
                {
                    filters.Add(new QueryFilter("Date", dateFL.ToString(), Operator.GreaterThanOrEqual));
                    filters.Add(new QueryFilter("Date", endDateFL.ToString(), Operator.LessThanOrEqualTo));
                }
            }

            List<CalendarModel> calendaModelList = _calendarAppService.Find(filters).DeserializeToCalendarList();

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList().Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

            string groupType = GetGroupType(groupTypeID);

            string emps = _empAppService.GetAll();
            List<EmployeeModel> empList = emps.DeserializeToEmployeeList();

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

            List<UserModel> validUserList = new List<UserModel>();
            foreach (var item in userList)
            {
                var tempEmp = empList.Where(x => x.EmployeeID.Trim() == item.EmployeeID).FirstOrDefault();
                if (tempEmp != null)
                {
                    JobTitleModel jt = jobTitleList.Where(x => x.ID == item.JobTitleID).FirstOrDefault();
                    item.RoleName = jt == null ? string.Empty : jt.RoleName;
                    item.GroupType = tempEmp.GroupType;
                    item.GroupName = tempEmp.GroupName;
                    item.FullName = tempEmp.FullName;

                    validUserList.Add(item);
                }
            }
            #endregion

            ICollection<QueryFilter> filterMpp = new List<QueryFilter>();
            filterMpp.Add(new QueryFilter("GroupType", groupType));
            if (string.IsNullOrEmpty(endDateFilter))
            {
                filterMpp.Add(new QueryFilter("Date", dateFL.Date.ToString()));
            }
            else
            {
                DateTime endDateFL = DateTime.Parse(endDateFilter);
                filterMpp.Add(new QueryFilter("Date", dateFL.ToString(), Operator.GreaterThanOrEqual));
                filterMpp.Add(new QueryFilter("Date", endDateFL.ToString(), Operator.LessThanOrEqualTo));
            }
            if (shift == "1" || shift == "2" || shift == "3" || shift == "NS")
                filterMpp.Add(new QueryFilter("Shift", shift.ToString()));

            string mpps = _mppAppService.Find(filterMpp);
            List<MppModel> mppList = mpps.DeserializeToMppList();
            mppList = mppList.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();

            List<MppModel> result = new List<MppModel>();
            foreach (var locationID in locationIdList)
            {
                var userTempList = validUserList.Where(x => x.LocationID == locationID).ToList();

                if (string.IsNullOrEmpty(endDateFilter) && calendaModelList.Count > 0)
                {
                    if (shift == "1")
                    {
                        userTempList = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendaModelList[0].Shift1).ToList();
                    }
                    else if (shift == "2")
                    {
                        userTempList = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendaModelList[0].Shift2).ToList();
                    }
                    else if (shift == "3")
                    {
                        userTempList = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendaModelList[0].Shift3).ToList();
                    }
                }

                if (string.IsNullOrEmpty(endDateFilter))
                {
                    foreach (var user in userTempList)
                    {
                        var tempMpp = mppList.Where(x => x.EmployeeID == user.EmployeeID).ToList();
                        if (tempMpp.Count == 0)
                        {
                            MppModel newMpp = new MppModel();
                            newMpp.Date = dateFL;
                            newMpp.LocationID = locationID;
                            newMpp.RoleName = user.RoleName;
                            newMpp.EmployeeID = user.EmployeeID;
                            newMpp.EmployeeName = user.FullName;
                            newMpp.Location = user.Location;
                            newMpp.Shift = shift == "0" ? "-" : shift;
                            newMpp.StatusMPP = user.GroupName == null ? "Unlisted in Group" : "Available";

                            result.Add(newMpp);
                        }
                        else
                        {
                            foreach (var item in tempMpp)
                            {
                                MppModel newMpp = new MppModel();
                                newMpp.Date = dateFL;
                                newMpp.LocationID = locationID;
                                newMpp.RoleName = user.RoleName;
                                newMpp.EmployeeID = user.EmployeeID;
                                newMpp.EmployeeName = user.FullName;
                                newMpp.Location = user.Location;
                                newMpp.StatusMPP = "Assigned";
                                newMpp.EmployeeMachine = item.EmployeeMachine;
                                newMpp.Shift = item.Shift;
                                newMpp.ID = item.ID;

                                result.Add(newMpp);
                            }
                        }
                    }
                }
                else
                {
                    DateTime endDateFL = DateTime.Parse(endDateFilter);
                    for (var day = dateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                    {
                        var calendarModel = calendaModelList.Where(x => x.Date == day).FirstOrDefault();
                        //if (calendarModel == null && shift != "0")
                        //    continue;

                        List<UserModel> userList2 = userTempList;

                        if (calendarModel != null)
                        {
                            if (shift == "1")
                            {
                                userList2 = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendarModel.Shift1).ToList();
                            }
                            else if (shift == "2")
                            {
                                userList2 = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendarModel.Shift2).ToList();
                            }
                            else if (shift == "3")
                            {
                                userList2 = userTempList.Where(x => x.GroupName != null && x.GroupName.Trim() == calendarModel.Shift3).ToList();
                            }
                        }

                        foreach (var user2 in userList2)
                        {
                            var tempMpp = mppList.Where(x => x.EmployeeID == user2.EmployeeID && x.Date == day).ToList();
                            if (tempMpp.Count == 0)
                            {
                                MppModel newMpp = new MppModel();
                                newMpp.Date = day;
                                newMpp.LocationID = locationID;
                                newMpp.RoleName = user2.RoleName;
                                newMpp.EmployeeID = user2.EmployeeID;
                                newMpp.EmployeeName = user2.FullName;
                                newMpp.Location = user2.Location;
                                newMpp.Shift = shift == "0" ? "-" : shift;
                                newMpp.StatusMPP = user2.GroupName == null ? "Unlisted in Group" : "Available";

                                result.Add(newMpp);
                            }
                            else
                            {
                                foreach (var item in tempMpp)
                                {
                                    MppModel newMpp = new MppModel();
                                    newMpp.Date = day;
                                    newMpp.LocationID = locationID;
                                    newMpp.RoleName = user2.RoleName;
                                    newMpp.EmployeeID = user2.EmployeeID;
                                    newMpp.EmployeeName = user2.FullName;
                                    newMpp.Location = user2.Location;
                                    newMpp.StatusMPP = "Assigned";
                                    newMpp.EmployeeMachine = item.EmployeeMachine;
                                    newMpp.Shift = item.Shift;
                                    newMpp.ID = item.ID;

                                    result.Add(newMpp);
                                }
                            }
                        }
                    }
                }
            }

            string machineAllocationList = _machineAllocationAppService.GetAll(true);
            List<MachineAllocationModel> machineAllocationModelList = machineAllocationList.DeserializeToMachineAllocationList();

            var assignedList = result.Where(x => x.StatusMPP == "Assigned").ToList();
            List<string> empIDList = result.Select(x => x.EmployeeID).Distinct().ToList();

            List<MppModel> tempList = new List<MppModel>();
            foreach (var item in empIDList)
            {
                if (shift == "1")
                {
                    tempList = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "1").ToList();
                    ValidateStatus(machineAllocationModelList, tempList);
                }
                else if (shift == "2")
                {
                    tempList = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "2").ToList();
                    ValidateStatus(machineAllocationModelList, tempList);
                }
                else if (shift == "3")
                {
                    tempList = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "3").ToList();
                    ValidateStatus(machineAllocationModelList, tempList);
                }
                else
                {
                    var tempList1 = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "1").ToList();
                    ValidateStatus(machineAllocationModelList, tempList1);
                    var tempList2 = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "2").ToList();
                    ValidateStatus(machineAllocationModelList, tempList2);
                    var tempList3 = assignedList.Where(x => x.EmployeeID == item && x.Shift != null && x.Shift.Trim() == "3").ToList();
                    ValidateStatus(machineAllocationModelList, tempList3);
                }
            }

            result = result.Where(x => x.StatusMPP != "Unlisted in Group" && x.StatusMPP != "Unallocated Machine").ToList();

            return result;
        }

        [HttpPost]
        public ActionResult GetAllStatusDashboard(string location, bool isDaily = true, long gtype = 1)
        {
            List<long> locationIdList = new List<long>();
            long locID = 0;
            if (!string.IsNullOrEmpty(location) && location != "All")
            {
                locID = _locationAppService.GetLocationID("ID-" + location);
                locationIdList = _locationAppService.GetLocIDListByLocType(locID, "productioncenter");
            }
            else
            {
                locationIdList = _locationAppService.GetLocIDListByLocType(1, "country");
            }

            DateTime dateFL = DateTime.Now;
            string endDateFL = string.Empty;

            if (!isDaily)
            {
                dateFL = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
                endDateFL = dateFL.AddDays(6).ToString();
            }

            int shift = Helper.GetCurrentShift();
            if (shift >= 3)
                dateFL = dateFL.AddDays(-1);
            List<MppModel> result = GetStatusData(dateFL, endDateFL, locationIdList, shift == 4 ? "3" : shift.ToString(), gtype, locID);

            MppDashboardModel modelDashboard = new MppDashboardModel();
            if (result.Count > 0)
            {
                modelDashboard.Available = result.Where(x => x.StatusMPP == "Available").Count();
                modelDashboard.HalfAssigned = result.Where(x => x.StatusMPP == "Half-Assigned").Count();
                modelDashboard.FullyAssigned = result.Where(x => x.StatusMPP == "Fully-Assigned").Count();
                modelDashboard.Overload = result.Where(x => x.StatusMPP == "Overload").Count();
                modelDashboard.Total = modelDashboard.Available + modelDashboard.HalfAssigned + modelDashboard.FullyAssigned + modelDashboard.Overload;
            }

            // Returning Json Data    
            return Json(modelDashboard, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetAllMachineDashboard(string location, bool isDaily = true)
        {
            List<long> locationIDList = new List<long>();
            if (!string.IsNullOrEmpty(location) && location != "All")
            {
                var pcID = _locationAppService.GetLocationID("ID-" + location);
                locationIDList = _locationAppService.GetLocIDListByLocType(pcID, "productioncenter");
            }
            else
            {
                locationIDList = _locationAppService.GetLocIDListByLocType(1, "country");
            }

            DateTime dateFL = DateTime.Now;
            int shift = Helper.GetCurrentShift();
            if (shift == 4)
            {
                dateFL = dateFL.AddDays(-1);
            }

            List<QueryFilter> filter = new List<QueryFilter>();
            if (isDaily)
            {
                filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd")));
                filter.Add(new QueryFilter("Shift", shift == 4 ? "3" : shift.ToString()));
            }
            else
            {
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);

                filter.Add(new QueryFilter("Date", monday.ToString(), Operator.GreaterThanOrEqual));
                filter.Add(new QueryFilter("Date", sunday.ToString(), Operator.LessThanOrEqualTo));
            }

            string mppList = _mppAppService.Find(filter);
            List<MppModel> mppModelList = mppList.DeserializeToMppList();
            if (location != "All")
            {
                mppModelList = mppModelList.Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();
            }

            filter = new List<QueryFilter>();
            if (isDaily)
            {
                filter.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd")));
            }
            else
            {
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
                var sunday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);

                filter.Add(new QueryFilter("Date", monday.ToString(), Operator.GreaterThanOrEqual));
                filter.Add(new QueryFilter("Date", sunday.ToString(), Operator.LessThanOrEqualTo));
            }

            string wpps = _wppStpAppService.Find(filter);
            List<WppStpModel> wppList1 = wpps.DeserializeToWppStpList();

            if (location != "All")
            {
                wppList1 = wppList1.Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();
            }

            List<MachineModel> machineListResult = new List<MachineModel>();

            List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList();

            if (wppList1.Count == 0)
            {
                machineListResult = machineList.Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();
            }
            else
            {
                List<string> wppMachineCodeList = wppList1.Where(x => !string.IsNullOrEmpty(x.Packer)).Select(x => x.Packer).Distinct().ToList();
                var temp = wppList1.Where(x => !string.IsNullOrEmpty(x.Maker)).Select(x => x.Maker).Distinct().ToList();
                if (temp.Count > 0)
                    wppMachineCodeList.AddRange(temp);

                foreach (var item in wppMachineCodeList)
                {
                    var wpp = wppList1.Where(x => x.Packer.Trim() == item.Trim() || x.Maker.Trim() == item.Trim()).First();

                    var machine = machineList.Where(x => x.Code.ToLower() == item.ToLower()).FirstOrDefault();
                    if (machine != null)
                    {
                        var tempMachine = machineListResult.Where(x => x.Code == machine.Code).FirstOrDefault();
                        if (tempMachine != null)
                        {
                            tempMachine.IsExistInWpp = true;
                            tempMachine.IsShift1Zero = wpp.Shift1 == 0;
                            tempMachine.IsShift2Zero = wpp.Shift2 == 0;
                            tempMachine.IsShift3Zero = wpp.Shift3 == 0;
                        }
                        else
                        {
                            //machine red
                            machine.IsExistInWpp = true;
                            machine.IsShift1Zero = wpp.Shift1 == 0;
                            machine.IsShift2Zero = wpp.Shift2 == 0;
                            machine.IsShift3Zero = wpp.Shift3 == 0;
                            machineListResult.Add(machine);
                        }
                    }
                }
            }

            string machineAllocations = _machineAllocationAppService.GetAll();
            List<MachineAllocationModel> machineAllocationList = machineAllocations.DeserializeToMachineAllocationList();

            foreach (var item in machineListResult)
            {
                var allocations = machineAllocationList.Where(x => x.MachineCode == item.Code).ToList();

                bool isFullyAssigned = true;
                if (shift == 1)
                {
                    foreach (var alloc in allocations)
                    {
                        var tempList = mppModelList.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "1").ToList();
                        double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                        double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                        if ((allocCount + allocCountLS) < alloc.Value)
                        {
                            isFullyAssigned = false;
                            break;
                        }
                    }

                    item.IsShift1FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;
                }
                else if (shift == 2)
                {
                    isFullyAssigned = true;
                    foreach (var alloc in allocations)
                    {
                        var tempList = mppModelList.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "2").ToList();
                        double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                        double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                        if ((allocCount + allocCountLS) < alloc.Value)
                        {
                            isFullyAssigned = false;
                            break;
                        }
                    }

                    item.IsShift2FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;
                }
                else
                {
                    isFullyAssigned = true;
                    foreach (var alloc in allocations)
                    {
                        var tempList = mppModelList.Where(x => x.EmployeeMachine != null && x.EmployeeMachine.Contains(alloc.MachineCode) && x.JobTitle.Contains(alloc.MachineCategory) && x.Shift.Trim() == "3").ToList();
                        double allocCount = tempList.Count(x => x.StatusMPP == "Normal");
                        double allocCountLS = tempList.Count(x => x.StatusMPP != "Normal") * 1.5;
                        if ((allocCount + allocCountLS) < alloc.Value)
                        {
                            isFullyAssigned = false;
                            break;
                        }
                    }

                    item.IsShift3FullyAssigned = allocations.Count == 0 ? false : isFullyAssigned;
                }
            }

            MppDashboardModel model = new MppDashboardModel();
            List<MachineModel> noWPP = machineListResult.Where(x => !x.IsExistInWpp).ToList();
            model.NotExistInWPP = noWPP.Count;
            machineListResult = machineListResult.Except(noWPP).ToList();

            List<MachineModel> unallocated = new List<MachineModel>();
            if (shift == 1)
                unallocated = machineListResult.Where(x => x.IsShift1Zero).ToList();
            else if (shift == 2)
                unallocated = machineListResult.Where(x => x.IsShift2Zero).ToList();
            else
                unallocated = machineListResult.Where(x => x.IsShift3Zero).ToList();

            model.Unallocated = unallocated.Count;
            machineListResult = machineListResult.Except(unallocated).ToList();

            List<MachineModel> fullAssigned = new List<MachineModel>();
            if (shift == 1)
                fullAssigned = machineListResult.Where(x => x.IsShift1FullyAssigned).ToList();
            else if (shift == 2)
                fullAssigned = machineListResult.Where(x => x.IsShift2FullyAssigned).ToList();
            else
                fullAssigned = machineListResult.Where(x => x.IsShift3FullyAssigned).ToList();

            model.FullyAssigned = fullAssigned.Count;
            machineListResult = machineListResult.Except(fullAssigned).ToList();

            List<MachineModel> halfAssigned = new List<MachineModel>();
            if (shift == 1)
                halfAssigned = machineListResult.Where(x => !x.IsShift1Zero).ToList();
            else if (shift == 2)
                halfAssigned = machineListResult.Where(x => !x.IsShift2Zero).ToList();
            else
                halfAssigned = machineListResult.Where(x => !x.IsShift3Zero).ToList();

            model.HalfAssigned = halfAssigned.Count;

            return Json(model, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GenerateExcel(string dateFilter, long locID, string locType)
        {
            try
            {
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);

                DateTime dateFL = DateTime.Parse(dateFilter);

                // Getting all data    			
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", dateFL.ToString()));

                string mppList = _mppAppService.Find(filters);
                List<MppModel> mppModelList = mppList.DeserializeToMppList().ToList();
                mppModelList = mppModelList.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();

                if (mppModelList.Count == 0)
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                    return RedirectToAction("Index");
                }

                byte[] excelData = ExcelGenerator.ExportMPPByDate(mppModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=MPP-Machine-List.xlsx");
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

        public ActionResult GenerateExcelDashboard(string startDateParam, string endDateParam)
        {
            try
            {
                List<MppSummaryModel> result = GetMPPSummaryData(startDateParam, endDateParam);

                if (result.Count == 0)
                {
                    SetFalseTempData(UIResources.NoDataInSelectedCriteria);
                    return RedirectToAction("Dashboard");
                }

                byte[] excelData = ExcelGenerator.ExportMPPDashboard(result, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-MPP-Dashboard.xlsx");
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
        public ActionResult GetAllMPPWithParam(string startDateFilter, string endDateFilter, long locID, string locType)
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

            DateTime startDateFL = DateTime.Parse(startDateFilter);
            DateTime endDateFL = DateTime.Parse(endDateFilter);

            // Getting all data    			
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", startDateFL.ToString(), Operator.GreaterThanOrEqual));
            filters.Add(new QueryFilter("Date", endDateFL.ToString(), Operator.LessThanOrEqualTo));

            string mppList = _mppAppService.Find(filters);
            List<MppModel> result = mppList.DeserializeToMppList().Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();

            int recordsTotal = result.Count();

            // Search     - Correction 231019
            if (!string.IsNullOrEmpty(searchValue))
            {
                result = result.Where(m => (m.JobTitle != null ? m.JobTitle.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.EmployeeName != null ? m.EmployeeName.ToLower().Contains(searchValue.ToLower()) : false) ||
                                           (m.EmployeeMachine != null ? m.EmployeeMachine.ToLower().Contains(searchValue.ToLower()) : false)
                                            ).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "startdate":
                            result = result.OrderBy(x => x.StartDate).ToList();
                            break;
                        case "enddate":
                            result = result.OrderBy(x => x.EndDate).ToList();
                            break;
                        case "date":
                            result = result.OrderBy(x => x.Date).ToList();
                            break;
                        case "jobtitle":
                            result = result.OrderBy(x => x.JobTitle).ToList();
                            break;
                        case "employeename":
                            result = result.OrderBy(x => x.EmployeeName).ToList();
                            break;
                        case "employeemachine":
                            result = result.OrderBy(x => x.EmployeeMachine).ToList();
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (sortColumn.ToLower())
                    {
                        case "startdate":
                            result = result.OrderByDescending(x => x.StartDate).ToList();
                            break;
                        case "enddate":
                            result = result.OrderByDescending(x => x.EndDate).ToList();
                            break;
                        case "date":
                            result = result.OrderByDescending(x => x.Date).ToList();
                            break;
                        case "jobtitle":
                            result = result.OrderByDescending(x => x.JobTitle).ToList();
                            break;
                        case "employeename":
                            result = result.OrderByDescending(x => x.EmployeeName).ToList();
                            break;
                        case "employeemachine":
                            result = result.OrderByDescending(x => x.EmployeeMachine).ToList();
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

        [HttpPost]
        public ActionResult GetMPPSummary(string startDateParam, string endDateParam)
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

            List<MppSummaryModel> result = GetMPPSummaryData(startDateParam, endDateParam);

            int recordsTotal = result.Count();

            // Search     - Correction 231019
            if (!string.IsNullOrEmpty(searchValue))
            {
                result = result.Where(m => (m.JobTitle != null ? m.JobTitle.ToLower().Contains(searchValue.ToLower()) : false)).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "jobtitle":
                            result = result.OrderBy(x => x.JobTitle).ToList();
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (sortColumn.ToLower())
                    {
                        case "jobtitle":
                            result = result.OrderByDescending(x => x.JobTitle).ToList();
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

        [HttpPost]
        public ActionResult GetAllChanges()
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
            string changes = _mppChangesAppService.GetAll(true);

            List<MppChangesModel> result = changes.DeserializeToMppChangesList().ToList();
            //foreach (var item in result)
            //{
            //    item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
            //}

            int recordsTotal = result.Count();

            // Search    
            if (!string.IsNullOrEmpty(searchValue))
            {
                result = result.Where(m => m.OldValue.ToLower().Contains(searchValue.ToLower()) ||
                                           m.NewValue.ToLower().Contains(searchValue.ToLower())).ToList();
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "old":
                            result = result.OrderBy(x => x.OldValue).ToList();
                            break;
                        case "new":
                            result = result.OrderBy(x => x.NewValue).ToList();
                            break;
                        case "field":
                            result = result.OrderBy(x => x.FieldName).ToList();
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
                            result = result.OrderByDescending(x => x.ID).ToList();
                            break;
                        case "old":
                            result = result.OrderByDescending(x => x.OldValue).ToList();
                            break;
                        case "new":
                            result = result.OrderByDescending(x => x.NewValue).ToList();
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

        [HttpPost]
        public ActionResult GenerateExcel(MppModel model)
        {
            try
            {
                UserModel user = (UserModel)Session["UserLogon"];
                if (!user.LocationID.HasValue)
                {
                    SetFalseTempData("Location for the logged user is invalid");
                    return RedirectToAction("Index");
                }

                long locationID = user.LocationID.Value;
                string gtype = _refDetailAppService.GetById(model.GroupTypeID, true);
                ReferenceDetailModel gTypeModel = gtype.DeserializeToRefDetail();

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", locationID.ToString()));
                filters.Add(new QueryFilter("GroupType", gTypeModel.Code));
                filters.Add(new QueryFilter("Date", model.StartDateCalendar.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", model.EndDateCalendar.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string mpps = _mppAppService.Find(filters);
                List<MppModel> mppList = mpps.DeserializeToMppList();
                if (mppList.Count == 0)
                {
                    return DownloadTemplate();
                }
                else
                {
                    DateTime startDate = model.StartDateCalendar.Value;
                    DateTime endDate = model.EndDateCalendar.Value;

                    byte[] excelData = ExcelGenerator.ExportMPPTemplate(mppList, startDate, endDate, gTypeModel.Code);

                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename=Template-MPP.xlsx");
                    Response.BinaryWrite(excelData);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult DownloadTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateMpp.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateMpp.xlsx");
                }
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Input");
        }

        [HttpPost]
        public ActionResult OvertimeAjax(long? locID, string locType, string filterDate = null)
        {
            DateTime startDate = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            if (!string.IsNullOrEmpty(filterDate))
            {
                DateTime filDate = DateTime.ParseExact(filterDate, "MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
                startDate = new DateTime(filDate.Year, 1, 1);
                endDate = filDate.AddMonths(1).AddDays(-1);
            }

            List<long> locationIDList = _locationAppService.GetLocIDListByLocType(locID.Value, locType);

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", startDate.ToString(), Operator.GreaterThanOrEqual));
            filters.Add(new QueryFilter("Date", endDate.ToString(), Operator.LessThanOrEqualTo));

            string empOverTimeList = _employeeOvertimeAppService.Find(filters);
            List<EmployeeOvertimeModel> result = empOverTimeList.DeserializeToEmployeeOvertimeList();

            if (!string.IsNullOrEmpty(locType) && locType != "country")
            {
                result = result.Where(x => !string.IsNullOrEmpty(x.Location) && locationIDList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
            }

            filters.Add(new QueryFilter("Date", startDate.ToString(), Operator.GreaterThanOrEqual));
            filters.Add(new QueryFilter("Date", endDate.ToString(), Operator.LessThanOrEqualTo));

            List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();
            mppList = mppList.Where(x => locationIDList.Any(y => y == x.LocationID)).ToList();

            var totalManPower = mppList.Count();

            string ytdYear = endDate.Year.ToString();
            List<int> monthList = result.Select(x => x.Date.Month).Distinct().ToList();
            List<string> overtimeCategoryList = result.Select(x => x.OvertimeCategory).Distinct().ToList();

            List<EmployeeOvertimeModel> resultList = new List<EmployeeOvertimeModel>();

            OvertimeRootModel rootModel = new OvertimeRootModel();
            List<string> monthListString = new List<string>();

            foreach (var month in monthList)
            {
                var manPowerPerMonth = mppList.Where(x => x.Date.Month == month).Count();
                OvertimeParentModel parent = new OvertimeParentModel();
                parent.Month = new DateTime(DateTime.Now.Year, month, 1).ToString("MMMM");
                parent.YTDYear = ytdYear;
                monthListString.Add(parent.Month);

                foreach (var category in overtimeCategoryList)
                {
                    OvertimeModel child = new OvertimeModel();
                    if (string.IsNullOrEmpty(category))
                    {
                        child.OvertimeCategory = "(blank)";
                    }
                    else if (category == "Back Up Leave" || category == "Backup Leave")
                    {
                        child.OvertimeCategory = "Backup Leave";
                    }
                    else
                    {
                        child.OvertimeCategory = category;
                    }

                    child.TotalManPower = manPowerPerMonth;

                    if (string.IsNullOrEmpty(category))
                    {
                        child.TotalOvertime = result.Where(x => string.IsNullOrEmpty(x.OvertimeCategory) && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Rework"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Other"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Daily"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Emergency"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category == "Back Up Leave" || category == "Backup Leave")
                    {
                        child.TotalOvertime = result.Where(x => (x.OvertimeCategory == "Backup Leave" || x.OvertimeCategory == "Back Up Leave") && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Maintenance"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category == "Leave")
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Volume"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category.Contains("Training"))
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category == "Project Activity")
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }
                    else if (category == "Project")
                    {
                        child.TotalOvertime = result.Where(x => x.OvertimeCategory == category && x.Date.Month == month).Sum(x => x.Overtime);
                    }

                    child.Percentage = child.TotalOvertime / child.TotalManPower;

                    parent.Children.Add(child);
                }

                parent.Children = parent.Children.OrderBy(x => x.OvertimeCategory).ToList();

                OvertimeModel childTotal = new OvertimeModel();
                childTotal.TotalManPower = parent.Children.Sum(x => x.TotalManPower);
                childTotal.TotalOvertime = result.Where(x => x.Date.Month == month).Sum(x => x.Overtime);
                childTotal.Percentage = parent.Children.Sum(x => x.Percentage);
                childTotal.OvertimeCategory = "Total";

                parent.Children.Add(childTotal);

                var totalAllOvertime = childTotal.TotalOvertime;
                foreach (var child in parent.Children)
                {
                    child.Percentage = (child.TotalOvertime / totalAllOvertime) * 100;
                    child.PercentageStr = String.Format("{0:0.00}", child.Percentage);
                }

                rootModel.Parents.Add(parent);
            }

            if (overtimeCategoryList.Any(x => x == null))
            {
                overtimeCategoryList = overtimeCategoryList.Where(x => !string.IsNullOrEmpty(x)).ToList();
                overtimeCategoryList.Add("(blank)");
                overtimeCategoryList = overtimeCategoryList.OrderBy(x => x).ToList();
            }

            if (overtimeCategoryList.Any(x => x == "Back Up Leave"))
            {
                overtimeCategoryList = overtimeCategoryList.Where(x => x != "Back Up Leave").ToList();
                if (!overtimeCategoryList.Any(x => x == "Backup Leave"))
                {
                    overtimeCategoryList.Add("Backup Leave");
                }
            }

            // add yeartd            
            OvertimeParentModel yearTdModel = new OvertimeParentModel();
            yearTdModel.Month = "year";
            yearTdModel.YTDYear = ytdYear;

            foreach (var category in overtimeCategoryList)
            {
                OvertimeModel child = new OvertimeModel();
                child.OvertimeCategory = category;
                child.TotalManPower = totalManPower;

                double totalOvertimeYear = 0;
                foreach (var parent in rootModel.Parents)
                {
                    totalOvertimeYear += parent.Children.Where(x => x.OvertimeCategory == category).Sum(x => x.TotalOvertime);
                }

                child.TotalOvertime = totalOvertimeYear;
                child.Percentage = child.TotalOvertime / child.TotalManPower;

                yearTdModel.Children.Add(child);
            }

            OvertimeModel childTot = new OvertimeModel();
            childTot.Percentage = yearTdModel.Children.Sum(x => x.Percentage);
            childTot.TotalManPower = yearTdModel.Children.Sum(x => x.TotalManPower);
            childTot.TotalOvertime = yearTdModel.Children.Sum(x => x.TotalOvertime);
            childTot.OvertimeCategory = "Total";

            yearTdModel.Children.Add(childTot);

            var totalAllOvertimeParent = childTot.TotalOvertime;
            foreach (var child in yearTdModel.Children)
            {
                child.Percentage = (child.TotalOvertime / totalAllOvertimeParent) * 100;
                child.PercentageStr = String.Format("{0:0.00}", child.Percentage);
            }

            overtimeCategoryList.Add("Total");

            rootModel.Parents.Add(yearTdModel);

            MPPOvertimeModel model = new MPPOvertimeModel();
            model.OvertimeList = rootModel.Parents;
            model.MonthList = monthListString;
            model.CategoryList = overtimeCategoryList;
            model.YearYTD = ytdYear;

            return Json(model, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Overtime()
        {
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

            return View();
        }

        [HttpPost]
        public ActionResult Upload(MppModel model)
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

                    #region Get Header 
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

                    string startDate = dt_.Rows[0][2].ToString();
                    DateTime startDateValue = DateTime.ParseExact(startDate, "yyyyMMdd", CultureInfo.CurrentCulture);
                    string endDate = dt_.Rows[0][4].ToString();
                    DateTime endDateValue = DateTime.ParseExact(endDate, "yyyyMMdd", CultureInfo.CurrentCulture);

                    if (startDateValue > endDateValue)
                    {
                        SetFalseTempData("Start Date must be less than End Date");
                        return RedirectToAction("Index");
                    }

                    if (startDateValue.Date < DateTime.Now.Date)
                    {
                        SetFalseTempData("Start Date must be equal or higher than today");
                        return RedirectToAction("Index");
                    }

                    string grouptypestr = dt_.Rows[0][6].ToString();
                    ReferenceDetailModel groupType = new ReferenceDetailModel();
                    if (string.IsNullOrEmpty(grouptypestr))
                    {
                        SetFalseTempData("Group type is missing. Please check the template");
                        return RedirectToAction("Index");
                    }
                    else
                    {
                        string groupTypes = _referenceAppService.GetDetailAll(ReferenceEnum.Group, true);
                        List<ReferenceDetailModel> groupTypeList = groupTypes.DeserializeToRefDetailList();

                        groupType = groupTypeList.Where(x => x.Code == grouptypestr).FirstOrDefault();
                    }
                    #endregion

                    //string empSkills = _employeeSkillAppService.GetAll(true);
                    //List<UserMachineTypeModel> empSkillList = empSkills.DeserializeToUserMachineTypeList();

                    //string userMachines = _userMachineAppService.GetAll(true);
                    //List<UserMachineModel> userMachineList = userMachines.DeserializeToUserMachineList();

                    //string users = _userAppService.GetAll(true);
                    //List<UserModel> userList = users.DeserializeToUserList();

                    //string jobtitles = _jobTitleAppService.GetAll(true);
                    //List<JobTitleModel> jobTitleList = jobtitles.DeserializeToJobTitleList();

                    string machineList = _machineAppService.GetAll(true);
                    List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

                    string locations = _locationAppService.GetAll(true);
                    List<LocationModel> locationModelList = locations.DeserializeToLocationList();

                    List<long> userLocIDList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");

                    string jt = string.Empty;
                    Dictionary<string, long> machineLocationMap = new Dictionary<string, long>();
                    List<long> machineLocationIDList = new List<long>();
                    List<MPPMiniModel> allMppMiniList = new List<MPPMiniModel>();

                    // grouping the data first
                    for (int index = 1; index < dt_.Rows.Count; index++)
                    {
                        string jobtitleTemp = dt_.Rows[index][0].ToString();
                        if (!string.IsNullOrEmpty(jobtitleTemp))
                        {
                            jt = jobtitleTemp;
                            continue;
                        }

                        string empIdA = dt_.Rows[index][1].ToString();
                        string empNameA = dt_.Rows[index][2].ToString();
                        string empMachineA = dt_.Rows[index][3].ToString();

                        string empIdB = dt_.Rows[index][4].ToString();
                        string empNameB = dt_.Rows[index][5].ToString();
                        string empMachineB = dt_.Rows[index][6].ToString();

                        string empIdC = dt_.Rows[index][7].ToString();
                        string empNameC = dt_.Rows[index][8].ToString();
                        string empMachineC = dt_.Rows[index][9].ToString();

                        string empIdD = dt_.Rows[index][10].ToString();
                        string empNameD = dt_.Rows[index][11].ToString();
                        string empMachineD = dt_.Rows[index][12].ToString();

                        if ((!string.IsNullOrEmpty(empIdA) && empIdA.All(char.IsDigit)) ||
                            (!string.IsNullOrEmpty(empIdB) && empIdB.All(char.IsDigit)) ||
                            (!string.IsNullOrEmpty(empIdC) && empIdC.All(char.IsDigit)) ||
                            (!string.IsNullOrEmpty(empIdD) && empIdD.All(char.IsDigit)))
                        {
                            MPPExcelModel modelExcel = new MPPExcelModel();
                            modelExcel.JobTitle = jt;

                            modelExcel.EmpIdA = empIdA;
                            modelExcel.EmpNameA = empNameA;
                            modelExcel.EmpMachineA = empMachineA;

                            modelExcel.EmpIdB = empIdB;
                            modelExcel.EmpNameB = empNameB;
                            modelExcel.EmpMachineB = empMachineB;

                            modelExcel.EmpIdC = empIdC;
                            modelExcel.EmpNameC = empNameC;
                            modelExcel.EmpMachineC = empMachineC;

                            modelExcel.EmpIdD = empIdD;
                            modelExcel.EmpNameD = empNameD;
                            modelExcel.EmpMachineD = empMachineD;

                            string errorMessage = string.Empty;
                            bool isValid = ValidateMachineLocation(allMppMiniList, modelExcel, machineModelList, userLocIDList, locationModelList, ref machineLocationIDList, out errorMessage);
                            if (!isValid)
                            {
                                SetFalseTempData(errorMessage);
                                return RedirectToAction("Index");
                            }

                            //string errorMessage = string.Empty;
                            //bool isValid = ValidateEmpSkillAndUserMachine(modelExcel, empSkillList, userMachineList, userList, machineModelList, out errorMessage);
                            //if (!isValid)
                            //{
                            //    SetFalseTempData(errorMessage);
                            //    return RedirectToAction("Index");
                            //}                            
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    string holidays = _calendarHolidayAppService.GetAll(true);
                    List<CalendarHolidayModel> resHoliday = holidays.DeserializeToCalendarHolidayList().OrderBy(x => x.Date).ToList();

                    List<MppModel> AddResult = new List<MppModel>();
                    List<MppModel> UpdateResult = new List<MppModel>();
                    List<MppModel> DeleteResult = new List<MppModel>();

                    string empList = _empAppService.GetAll(true);
                    List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

                    List<string> empIDLeaveList = new List<string>();

                    foreach (var locationID in machineLocationIDList)
                    {
                        string location = _locationAppService.GetLocationFullCode(locationID);

                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("LocationID", locationID.ToString()));
                        filters.Add(new QueryFilter("GroupTypeID", groupType.ID.ToString()));
                        filters.Add(new QueryFilter("Date", startDateValue.ToString(), Operator.GreaterThanOrEqual));
                        filters.Add(new QueryFilter("Date", endDateValue.ToString(), Operator.LessThanOrEqualTo));
                        filters.Add(new QueryFilter("IsDeleted", "0"));

                        string calendar = _calendarAppService.Find(filters);
                        List<CalendarModel> calendarList = calendar.DeserializeToCalendarList();

                        if (calendarList.Count > 0)
                        {
                            ICollection<QueryFilter> filterMpps = new List<QueryFilter>();
                            filterMpps.Add(new QueryFilter("LocationID", locationID.ToString()));
                            filterMpps.Add(new QueryFilter("GroupType", groupType.Code.ToString()));
                            filterMpps.Add(new QueryFilter("Date", startDateValue.ToString(), Operator.GreaterThanOrEqual));
                            filterMpps.Add(new QueryFilter("Date", endDateValue.ToString(), Operator.LessThanOrEqualTo));
                            filterMpps.Add(new QueryFilter("IsDeleted", "0"));

                            string oldMpp = _mppAppService.Find(filterMpps);
                            List<MppModel> oldMppList = oldMpp.DeserializeToMppList();

                            string emLeaveList = _employeeLeaveAppService.GetAll();
                            List<EmployeeLeaveModel> emLeaveModelList = emLeaveList.DeserializeToEmployeeLeaveList();

                            for (var day = startDateValue.Date; day.Date <= endDateValue.Date; day = day.AddDays(1))
                            {
                                List<string> employeeIDList = new List<string>();
                                List<MppModel> result = new List<MppModel>();

                                CalendarModel dateCalendar = calendarList.Where(x => x.Date == day.Date).FirstOrDefault();
                                if (dateCalendar != null)
                                {
                                    var tempMppList = allMppMiniList.Where(x => x.EmpMachineLocationID == locationID).ToList();
                                    foreach (var mpp in tempMppList)
                                    {
                                        if (mpp.Group == "A" && (dateCalendar.Shift1 == "A" || dateCalendar.Shift2 == "A" || dateCalendar.Shift3 == "A"))
                                        {
                                            #region Group A
                                            MppModel newMpp = new MppModel();
                                            newMpp.GroupName = "A";
                                            newMpp.GroupType = groupType.Code;
                                            newMpp.Year = day.Date.Year;
                                            newMpp.Week = GetCurrentWeekNumber(day.Date);
                                            newMpp.StartDate = startDateValue;
                                            newMpp.EndDate = endDateValue;
                                            newMpp.Date = day.Date;
                                            newMpp.LocationID = locationID;
                                            newMpp.ModifiedDate = DateTime.Now;
                                            newMpp.ModifiedBy = AccountName;
                                            newMpp.JobTitle = mpp.JobTitle;
                                            newMpp.EmployeeID = mpp.EmpId;
                                            newMpp.EmployeeName = mpp.EmpName;
                                            newMpp.EmployeeMachine = mpp.EmpMachine;

                                            if (dateCalendar.Shift1 == "A")
                                                newMpp.Shift = "1";
                                            else if (dateCalendar.Shift2 == "A")
                                                newMpp.Shift = "2";
                                            else
                                                newMpp.Shift = "3";

                                            if (!string.IsNullOrEmpty(newMpp.EmployeeID) && newMpp.EmployeeID.All(Char.IsDigit))
                                            {
                                                var checkEmpIDA = empModelList.Where(x => x.EmployeeID == newMpp.EmployeeID).FirstOrDefault();
                                                if (checkEmpIDA == null)
                                                {
                                                    SetFalseTempData(string.Format(UIResources.EmployeeNotExist, newMpp.EmployeeID) + " Please check value in B column");
                                                    return RedirectToAction("Index");
                                                }
                                                else if (IsEmployeeOnLeave(newMpp.EmployeeID, day.Date, emLeaveModelList))
                                                {
                                                    empIDLeaveList.Add(newMpp.EmployeeID + " is on leave at " + day.Date.ToString("dd-MMM-yy"));
                                                }
                                            }
                                            else if (!string.IsNullOrEmpty(newMpp.EmployeeName))
                                            {
                                                SetFalseTempData(string.Format(UIResources.EmployeeIdIsMissing, " Please check value in B column"));
                                                return RedirectToAction("Index");
                                            }

                                            result.Add(newMpp);
                                            employeeIDList.Add(newMpp.EmployeeID);
                                            #endregion
                                        }

                                        if (mpp.Group == "B" && (dateCalendar.Shift1 == "B" || dateCalendar.Shift2 == "B" || dateCalendar.Shift3 == "B"))
                                        {
                                            #region Group B
                                            MppModel newMpp = new MppModel();
                                            newMpp.GroupName = "B";
                                            newMpp.GroupType = groupType.Code;
                                            newMpp.Year = day.Date.Year;
                                            newMpp.Week = GetCurrentWeekNumber(day.Date);
                                            newMpp.StartDate = startDateValue;
                                            newMpp.EndDate = endDateValue;
                                            newMpp.Date = day.Date;
                                            newMpp.LocationID = locationID;
                                            newMpp.ModifiedDate = DateTime.Now;
                                            newMpp.ModifiedBy = AccountName;
                                            newMpp.JobTitle = mpp.JobTitle;
                                            newMpp.EmployeeID = mpp.EmpId;
                                            newMpp.EmployeeName = mpp.EmpName;
                                            newMpp.EmployeeMachine = mpp.EmpMachine;

                                            if (dateCalendar.Shift1 == "B")
                                                newMpp.Shift = "1";
                                            else if (dateCalendar.Shift2 == "B")
                                                newMpp.Shift = "2";
                                            else
                                                newMpp.Shift = "3";

                                            if (!string.IsNullOrEmpty(newMpp.EmployeeID) && newMpp.EmployeeID.All(Char.IsDigit))
                                            {
                                                var checkEmpIDB = empModelList.Where(x => x.EmployeeID == newMpp.EmployeeID).FirstOrDefault();
                                                if (checkEmpIDB == null)
                                                {
                                                    SetFalseTempData(string.Format(UIResources.EmployeeNotExist, newMpp.EmployeeID) + " Please check value in E column");
                                                    return RedirectToAction("Index");
                                                }
                                                else if (IsEmployeeOnLeave(newMpp.EmployeeID, day.Date, emLeaveModelList))
                                                {
                                                    empIDLeaveList.Add(newMpp.EmployeeID + " is on leave at " + day.Date.ToString("dd-MMM-yy"));
                                                }
                                            }
                                            else if (!string.IsNullOrEmpty(newMpp.EmployeeName))
                                            {
                                                SetFalseTempData(string.Format(UIResources.EmployeeIdIsMissing, " Please check value in E column"));
                                                return RedirectToAction("Index");
                                            }

                                            result.Add(newMpp);
                                            employeeIDList.Add(newMpp.EmployeeID);
                                            #endregion
                                        }

                                        if (mpp.Group == "C" && (dateCalendar.Shift1 == "C" || dateCalendar.Shift2 == "C" || dateCalendar.Shift3 == "C"))
                                        {
                                            #region Group C
                                            MppModel newMpp = new MppModel();
                                            newMpp.GroupName = "C";
                                            newMpp.GroupType = groupType.Code;
                                            newMpp.Year = day.Date.Year;
                                            newMpp.Week = GetCurrentWeekNumber(day.Date);
                                            newMpp.StartDate = startDateValue;
                                            newMpp.EndDate = endDateValue;
                                            newMpp.Date = day.Date;
                                            newMpp.LocationID = locationID;
                                            newMpp.ModifiedDate = DateTime.Now;
                                            newMpp.ModifiedBy = AccountName;
                                            newMpp.JobTitle = mpp.JobTitle;
                                            newMpp.EmployeeID = mpp.EmpId;
                                            newMpp.EmployeeName = mpp.EmpName;
                                            newMpp.EmployeeMachine = mpp.EmpMachine;

                                            if (dateCalendar.Shift1 == "C")
                                                newMpp.Shift = "1";
                                            else if (dateCalendar.Shift2 == "C")
                                                newMpp.Shift = "2";
                                            else
                                                newMpp.Shift = "3";

                                            if (!string.IsNullOrEmpty(newMpp.EmployeeID) && newMpp.EmployeeID.All(Char.IsDigit))
                                            {
                                                var checkEmpIDC = empModelList.Where(x => x.EmployeeID == newMpp.EmployeeID).FirstOrDefault();
                                                if (checkEmpIDC == null)
                                                {
                                                    SetFalseTempData(string.Format(UIResources.EmployeeNotExist, newMpp.EmployeeID) + " Please check value in H column");
                                                    return RedirectToAction("Index");
                                                }
                                                else if (IsEmployeeOnLeave(newMpp.EmployeeID, day.Date, emLeaveModelList))
                                                {
                                                    empIDLeaveList.Add(newMpp.EmployeeID + " is on leave at " + day.Date.ToString("dd-MMM-yy"));
                                                }
                                            }
                                            else if (!string.IsNullOrEmpty(newMpp.EmployeeName))
                                            {
                                                SetFalseTempData(string.Format(UIResources.EmployeeIdIsMissing, " Please check value in H column"));
                                                return RedirectToAction("Index");
                                            }

                                            result.Add(newMpp);
                                            employeeIDList.Add(newMpp.EmployeeID);
                                            #endregion
                                        }

                                        if (mpp.Group == "D" && (dateCalendar.Shift1 == "D" || dateCalendar.Shift2 == "D" || dateCalendar.Shift3 == "D"))
                                        {
                                            #region Group D
                                            if (!mpp.JobTitle.Equals("SUPPORT"))
                                            {
                                                MppModel newMpp = new MppModel();
                                                newMpp.GroupName = "D";
                                                newMpp.GroupType = groupType.Code;
                                                newMpp.Year = day.Date.Year;
                                                newMpp.Week = GetCurrentWeekNumber(day.Date);
                                                newMpp.StartDate = startDateValue;
                                                newMpp.EndDate = endDateValue;
                                                newMpp.Date = day.Date;
                                                newMpp.LocationID = locationID;
                                                newMpp.ModifiedDate = DateTime.Now;
                                                newMpp.ModifiedBy = AccountName;
                                                newMpp.JobTitle = mpp.JobTitle;
                                                newMpp.EmployeeID = mpp.EmpId;
                                                newMpp.EmployeeName = mpp.EmpName;
                                                newMpp.EmployeeMachine = mpp.EmpMachine;

                                                if (dateCalendar.Shift1 == "D")
                                                    newMpp.Shift = "1";
                                                else if (dateCalendar.Shift2 == "D")
                                                    newMpp.Shift = "2";
                                                else
                                                    newMpp.Shift = "3";

                                                if (!string.IsNullOrEmpty(newMpp.EmployeeID) && newMpp.EmployeeID.All(Char.IsDigit))
                                                {
                                                    var checkEmpIDD = empModelList.Where(x => x.EmployeeID == newMpp.EmployeeID).FirstOrDefault();
                                                    if (checkEmpIDD == null)
                                                    {
                                                        SetFalseTempData(string.Format(UIResources.EmployeeNotExist, newMpp.EmployeeID) + " Please check value in K column");
                                                        return RedirectToAction("Index");
                                                    }
                                                    else if (IsEmployeeOnLeave(newMpp.EmployeeID, day.Date, emLeaveModelList))
                                                    {
                                                        empIDLeaveList.Add(newMpp.EmployeeID + " is on leave at " + day.Date.ToString("dd-MMM-yy"));
                                                    }
                                                }
                                                else if (!string.IsNullOrEmpty(newMpp.EmployeeName))
                                                {
                                                    SetFalseTempData(string.Format(UIResources.EmployeeIdIsMissing, " Please check value in K column"));
                                                    return RedirectToAction("Index");
                                                }

                                                result.Add(newMpp);
                                                employeeIDList.Add(newMpp.EmployeeID);
                                            }
                                            #endregion
                                        }
                                    }
                                }

                                result = result.OrderBy(x => x.Date).ThenBy(x => x.Shift).ToList();

                                List<MppModel> filteredOldMppList = oldMppList.Where(x => x.Date.Date == day.Date).ToList();

                                foreach (var item in result)
                                {
                                    item.StatusMPP = "Normal";
                                    item.Location = location;

                                    if (filteredOldMppList.Count > 0)
                                    {
                                        List<MppModel> duplicateMppObjects = filteredOldMppList.Where(x => x.EmployeeMachine == item.EmployeeMachine && x.Shift.Trim() == item.Shift.Trim()).ToList();
                                        if (duplicateMppObjects.Count > 0)
                                        {
                                            foreach (var dup in duplicateMppObjects)
                                            {
                                                if (!DeleteResult.Any(x => x.ID == dup.ID))
                                                    DeleteResult.Add(dup);
                                            }

                                            AddResult.Add(item);
                                            continue;
                                        }

                                        MppModel oldMppObject = filteredOldMppList.Where(x => x.EmployeeID == item.EmployeeID && x.Shift.Trim() == item.Shift.Trim()).FirstOrDefault();
                                        if (oldMppObject != null)
                                        {
                                            item.ID = oldMppObject.ID;
                                            item.ModifiedBy = AccountName;
                                            item.ModifiedDate = DateTime.Now;

                                            UpdateResult.Add(item);

                                            continue;
                                        }
                                    }

                                    AddResult.Add(item);
                                }
                            }
                        }
                        else
                        {
                            SetFalseTempData("There is no calendar defined on the selected date or location");
                            return RedirectToAction("Index");
                        }
                    }

                    if (AddUpdateDeleteMPP(AddResult, UpdateResult, DeleteResult))
                    {
                        if (empIDLeaveList.Count > 0)
                            SetTrueTempData(UIResources.UploadSucceed + string.Join(",", empIDLeaveList));
                        else
                            SetTrueTempData(UIResources.UploadSucceed);
                    }
                    else
                    {
                        SetFalseTempData("Upload failed. Please try it again");
                    }
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
        #endregion

        #region ::Private Methods::
        private List<UserModel> ValidateUserByMachineCode(List<UserModel> userList, List<UserMachineTypeModel> empSkillList, List<UserMachineModel> userMachineList, MachineModel machine)
        {
            List<UserModel> result = new List<UserModel>();

            foreach (var user in userList)
            {
                var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
                if (uskills.Any(x => x.MachineTypeID == machine.MachineTypeID))
                {
                    result.Add(user);
                    continue;
                }

                var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();
                if (umachines.Any(x => x.MachineID == machine.ID))
                {
                    result.Add(user);
                    continue;
                }
            }

            return result;
        }

        private static string GetCompetencyFromTrainings(List<ReferenceDetailModel> machineTypeReferenceList, List<TrainingModel> trainingList, UserModel user)
        {
            var trainingEmpList = trainingList.Where(x => x.EmployeeID == user.EmployeeID).ToList();

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

            return machineList;
        }

        private static void ValidateStatus(List<MachineAllocationModel> machineAllocationModelList, List<MppModel> tempList)
        {
            if (tempList.Count > 0)
            {
                double totalAllocation = 0;
                foreach (var emp in tempList)
                {
                    var ma = machineAllocationModelList.Where(x => x.MachineCode == emp.EmployeeMachine && x.MachineCategory == emp.RoleName).FirstOrDefault();
                    if (ma != null && ma.Value.HasValue)
                    {
                        totalAllocation += ma.Value.Value;
                    }
                }

                if (totalAllocation == 0)
                {
                    tempList.ForEach(x => x.StatusMPP = "Unallocated Machine");
                }
                else if (totalAllocation < 1)
                {
                    tempList.ForEach(x => x.StatusMPP = "Half-Assigned");
                }
                else if (totalAllocation > 1)
                {
                    tempList.ForEach(x => x.StatusMPP = "Overload");
                }
                else if (totalAllocation == 1)
                {
                    tempList.ForEach(x => x.StatusMPP = "Fully-Assigned");
                }
            }
        }

        private void UpdateJobTitle(List<ManPowerModel> mpList)
        {
            Dictionary<long, string> jobTitleListMap = new Dictionary<long, string>();

            foreach (var item in mpList)
            {
                if (jobTitleListMap.ContainsKey(item.JobTitleID))
                {
                    string jt = "";
                    jobTitleListMap.TryGetValue(item.JobTitleID, out jt);
                    item.JobTitle = jt;
                }
                else
                {
                    string jt = _jobTitleAppService.GetById(item.JobTitleID);
                    JobTitleModel jtModel = jt.DeserializeToJobTitle();
                    item.JobTitle = jtModel.Title;
                    jobTitleListMap.Add(item.JobTitleID, jtModel.Title);
                }
            }
        }

        private string GetGroupType(long groupTypeID)
        {
            string gtype = _refDetailAppService.GetById(groupTypeID, true);
            ReferenceDetailModel gTypeModel = gtype.DeserializeToRefDetail();

            return gTypeModel.Code;
        }

        private string GetGroupName(long locID, string shift, DateTime dateFL)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("LocationID", locID.ToString()));
            filters.Add(new QueryFilter("Date", dateFL.ToString("yyyy-MM-dd")));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string calendar = _calendarAppService.Find(filters);
            List<CalendarModel> calendarList = calendar.DeserializeToCalendarList();

            string groupname = "A";

            if (calendarList.Count > 0)
            {
                foreach (var cl in calendarList)
                {
                    if (shift == "1")
                    {
                        groupname = cl.Shift1;
                        break;
                    }
                    else if (shift == "2")
                    {
                        groupname = cl.Shift2;
                        break;
                    }
                    else if (shift == "3")
                    {
                        groupname = cl.Shift3;
                        break;
                    }
                }
            }

            return groupname;
        }

        private bool ValidateMachineLocation(List<MPPMiniModel> mppMiniList, MPPExcelModel modelExcel, List<MachineModel> machineList, List<long> userLocIDList, List<LocationModel> locationList, ref List<long> machineLocationIDList, out string errorMessage)
        {
            errorMessage = string.Empty;

            if (!string.IsNullOrEmpty(modelExcel.EmpIdA))
            {
                if (string.IsNullOrEmpty(modelExcel.EmpMachineA))
                {
                    errorMessage = modelExcel.EmpIdA + " does not have machine code";
                    return false;
                }
                else if (mppMiniList.Any(x => x.EmpId == modelExcel.EmpIdA))
                {
                    errorMessage = modelExcel.EmpIdA + " double assignment";
                    return false;
                }
                else if (modelExcel.EmpMachineA.Contains(','))
                {
                    string[] machineCodes = modelExcel.EmpMachineA.Replace(" ", "").Split(',');
                    foreach (var mcode in machineCodes)
                    {
                        var machine = machineList.Where(x => x.Code == mcode).FirstOrDefault();
                        if (machine != null && machine.LocationID.HasValue)
                        {
                            if (userLocIDList.Any(x => x == machine.LocationID.Value))
                            {
                                MPPMiniModel miniMpp = new MPPMiniModel
                                {
                                    Group = "A",
                                    EmpId = modelExcel.EmpIdA,
                                    EmpName = modelExcel.EmpNameA,
                                    EmpMachine = mcode,
                                    EmpMachineLocationID = machine.LocationID.Value,
                                    JobTitle = modelExcel.JobTitle
                                };

                                mppMiniList.Add(miniMpp);

                                if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                    machineLocationIDList.Add(machine.LocationID.Value);
                            }
                            else
                            {
                                errorMessage = mcode + " is located in another production center";
                                return false;
                            }
                        }
                        else
                        {
                            errorMessage = mcode + " does not have location ID";
                            return false;
                        }
                    }
                }
                else
                {
                    var machine = machineList.Where(x => x.Code == modelExcel.EmpMachineA.Trim()).FirstOrDefault();
                    if (machine != null && machine.LocationID.HasValue)
                    {
                        if (userLocIDList.Any(x => x == machine.LocationID.Value))
                        {
                            MPPMiniModel miniMpp = new MPPMiniModel
                            {
                                Group = "A",
                                EmpId = modelExcel.EmpIdA,
                                EmpName = modelExcel.EmpNameA,
                                EmpMachine = modelExcel.EmpMachineA,
                                EmpMachineLocationID = machine.LocationID.Value,
                                JobTitle = modelExcel.JobTitle
                            };

                            mppMiniList.Add(miniMpp);

                            if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                machineLocationIDList.Add(machine.LocationID.Value);
                        }
                        else
                        {
                            errorMessage = modelExcel.EmpMachineA + " is located in another production center";
                            return false;
                        }
                    }
                    else
                    {
                        errorMessage = modelExcel.EmpMachineA + " does not have location ID";
                        return false;
                    }
                }
            }

            if (!string.IsNullOrEmpty(modelExcel.EmpIdB))
            {
                if (string.IsNullOrEmpty(modelExcel.EmpMachineB))
                {
                    errorMessage = modelExcel.EmpIdB + " does not have machine code";
                    return false;
                }
                else if (mppMiniList.Any(x => x.EmpId == modelExcel.EmpIdB))
                {
                    errorMessage = modelExcel.EmpIdB + " double assignment";
                    return false;
                }
                else if (modelExcel.EmpMachineB.Contains(','))
                {
                    string[] machineCodes = modelExcel.EmpMachineB.Replace(" ", "").Split(',');
                    foreach (var mcode in machineCodes)
                    {
                        var machine = machineList.Where(x => x.Code == mcode).FirstOrDefault();
                        if (machine != null && machine.LocationID.HasValue)
                        {
                            if (userLocIDList.Any(x => x == machine.LocationID.Value))
                            {
                                MPPMiniModel miniMpp = new MPPMiniModel
                                {
                                    Group = "B",
                                    EmpId = modelExcel.EmpIdB,
                                    EmpName = modelExcel.EmpNameB,
                                    EmpMachine = mcode,
                                    EmpMachineLocationID = machine.LocationID.Value,
                                    JobTitle = modelExcel.JobTitle
                                };

                                mppMiniList.Add(miniMpp);

                                if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                    machineLocationIDList.Add(machine.LocationID.Value);
                            }
                            else
                            {
                                errorMessage = mcode + " is located in another production center";
                                return false;
                            }
                        }
                        else
                        {
                            errorMessage = mcode + " does not have location ID";
                            return false;
                        }
                    }
                }
                else
                {
                    var machine = machineList.Where(x => x.Code == modelExcel.EmpMachineB.Trim()).FirstOrDefault();
                    if (machine != null && machine.LocationID.HasValue)
                    {
                        if (userLocIDList.Any(x => x == machine.LocationID.Value))
                        {
                            MPPMiniModel miniMpp = new MPPMiniModel
                            {
                                Group = "B",
                                EmpId = modelExcel.EmpIdB,
                                EmpName = modelExcel.EmpNameB,
                                EmpMachine = modelExcel.EmpMachineB,
                                EmpMachineLocationID = machine.LocationID.Value,
                                JobTitle = modelExcel.JobTitle
                            };

                            mppMiniList.Add(miniMpp);

                            if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                machineLocationIDList.Add(machine.LocationID.Value);
                        }
                        else
                        {
                            errorMessage = modelExcel.EmpMachineB + " is located in another production center";
                            return false;
                        }
                    }
                    else
                    {
                        errorMessage = modelExcel.EmpMachineB + " does not have location ID";
                        return false;
                    }
                }
            }

            if (!string.IsNullOrEmpty(modelExcel.EmpIdC))
            {
                if (string.IsNullOrEmpty(modelExcel.EmpMachineC))
                {
                    errorMessage = modelExcel.EmpIdC + " does not have machine code";
                    return false;
                }
                else if (mppMiniList.Any(x => x.EmpId == modelExcel.EmpIdC))
                {
                    errorMessage = modelExcel.EmpIdC + " double assignment";
                    return false;
                }
                else if (modelExcel.EmpMachineC.Contains(','))
                {
                    string[] machineCodes = modelExcel.EmpMachineC.Replace(" ", "").Split(',');
                    foreach (var mcode in machineCodes)
                    {
                        var machine = machineList.Where(x => x.Code == mcode).FirstOrDefault();
                        if (machine != null && machine.LocationID.HasValue)
                        {
                            if (userLocIDList.Any(x => x == machine.LocationID.Value))
                            {
                                MPPMiniModel miniMpp = new MPPMiniModel
                                {
                                    Group = "C",
                                    EmpId = modelExcel.EmpIdC,
                                    EmpName = modelExcel.EmpNameC,
                                    EmpMachine = mcode,
                                    EmpMachineLocationID = machine.LocationID.Value,
                                    JobTitle = modelExcel.JobTitle
                                };

                                mppMiniList.Add(miniMpp);

                                if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                    machineLocationIDList.Add(machine.LocationID.Value);
                            }
                            else
                            {
                                errorMessage = mcode + " is located in another production center";
                                return false;
                            }
                        }
                        else
                        {
                            errorMessage = mcode + " does not have location ID";
                            return false;
                        }
                    }
                }
                else
                {
                    var machine = machineList.Where(x => x.Code == modelExcel.EmpMachineC.Trim()).FirstOrDefault();
                    if (machine != null && machine.LocationID.HasValue)
                    {
                        if (userLocIDList.Any(x => x == machine.LocationID.Value))
                        {
                            MPPMiniModel miniMpp = new MPPMiniModel
                            {
                                Group = "C",
                                EmpId = modelExcel.EmpIdC,
                                EmpName = modelExcel.EmpNameC,
                                EmpMachine = modelExcel.EmpMachineC,
                                EmpMachineLocationID = machine.LocationID.Value,
                                JobTitle = modelExcel.JobTitle
                            };

                            mppMiniList.Add(miniMpp);

                            if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                machineLocationIDList.Add(machine.LocationID.Value);
                        }
                        else
                        {
                            errorMessage = modelExcel.EmpMachineC + " is located in another production center";
                            return false;
                        }
                    }
                    else
                    {
                        errorMessage = modelExcel.EmpMachineC + " does not have location ID";
                        return false;
                    }
                }
            }

            if (!string.IsNullOrEmpty(modelExcel.EmpIdD))
            {
                if (string.IsNullOrEmpty(modelExcel.EmpMachineD))
                {
                    errorMessage = modelExcel.EmpIdD + " does not have machine code";
                    return false;
                }
                else if (mppMiniList.Any(x => x.EmpId == modelExcel.EmpIdD))
                {
                    errorMessage = modelExcel.EmpIdD + " double assignment";
                    return false;
                }
                else if (modelExcel.EmpMachineD.Contains(','))
                {
                    string[] machineCodes = modelExcel.EmpMachineD.Replace(" ", "").Split(',');
                    foreach (var mcode in machineCodes)
                    {
                        var machine = machineList.Where(x => x.Code == mcode).FirstOrDefault();
                        if (machine != null && machine.LocationID.HasValue)
                        {
                            if (userLocIDList.Any(x => x == machine.LocationID.Value))
                            {
                                MPPMiniModel miniMpp = new MPPMiniModel
                                {
                                    Group = "D",
                                    EmpId = modelExcel.EmpIdD,
                                    EmpName = modelExcel.EmpNameD,
                                    EmpMachine = mcode,
                                    EmpMachineLocationID = machine.LocationID.Value,
                                    JobTitle = modelExcel.JobTitle
                                };

                                mppMiniList.Add(miniMpp);

                                if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                    machineLocationIDList.Add(machine.LocationID.Value);
                            }
                            else
                            {
                                errorMessage = mcode + " is located in another production center";
                                return false;
                            }
                        }
                        else
                        {
                            errorMessage = mcode + " does not have location ID";
                            return false;
                        }
                    }
                }
                else
                {
                    var machine = machineList.Where(x => x.Code == modelExcel.EmpMachineD.Trim()).FirstOrDefault();
                    if (machine != null && machine.LocationID.HasValue)
                    {
                        if (userLocIDList.Any(x => x == machine.LocationID.Value))
                        {
                            MPPMiniModel miniMpp = new MPPMiniModel
                            {
                                Group = "D",
                                EmpId = modelExcel.EmpIdD,
                                EmpName = modelExcel.EmpNameD,
                                EmpMachine = modelExcel.EmpMachineD,
                                EmpMachineLocationID = machine.LocationID.Value,
                                JobTitle = modelExcel.JobTitle
                            };

                            mppMiniList.Add(miniMpp);

                            if (!machineLocationIDList.Any(x => x == machine.LocationID.Value))
                                machineLocationIDList.Add(machine.LocationID.Value);
                        }
                        else
                        {
                            errorMessage = modelExcel.EmpMachineD + " is located in another production center";
                            return false;
                        }
                    }
                    else
                    {
                        errorMessage = modelExcel.EmpMachineD + " does not have location ID";
                        return false;
                    }
                }
            }

            return true;
        }

        private List<MppSummaryModel> GetMPPSummaryData(string startDateParam, string endDateParam)
        {
            // Getting all data    			
            string mppList = _mppAppService.GetAll(true);
            List<MppModel> mppModeliList = mppList.DeserializeToMppList();

            string empList = _empAppService.GetAll();
            List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

            DateTime startDate = DateTime.Parse(startDateParam);
            DateTime endDate = DateTime.Parse(endDateParam);

            List<MppSummaryModel> result = new List<MppSummaryModel>();

            int empTotal = empModelList.Count();
            int prodtechTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Production Technician")).Count();
            int foremanTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Foreman")).Count();
            int mechanicTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Mechanic")).Count();
            int electricianTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Electrician")).Count();
            int teamleaderTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Team Leader")).Count();
            int reliefTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Team Leader")).Count();
            int supportTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Support")).Count();
            int generalworkerTotal = empModelList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("General Worker")).Count();
            int otherTotal = empTotal - (prodtechTotal + foremanTotal + mechanicTotal + electricianTotal + teamleaderTotal + reliefTotal + supportTotal + generalworkerTotal);

            for (DateTime day = startDate; day.Date <= endDate; day = day.AddDays(1))
            {
                int prodtech = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "PRODTECH" && x.Date == day).Count();
                int foreman = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "FOREMAN" && x.Date == day).Count();
                int mechanic = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "MECHANIC" && x.Date == day).Count();
                int electrician = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "ELECTRICIAN" && x.Date == day).Count();
                int teamleader = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "TEAMLEADER" && x.Date == day).Count();
                int relief = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "RELIEF" && x.Date == day).Count();
                int support = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "SUPPORT" && x.Date == day).Count();
                int generalworker = mppModeliList.Where(x => x.JobTitle != null && x.JobTitle == "GENERALWORKER" && x.Date == day).Count();
                int other = mppModeliList.Count() - (prodtech + foreman + mechanic + electrician + teamleader + relief + support + generalworker);

                #region ::Sum by Title::
                MppSummaryModel prodTechModel = new MppSummaryModel
                {
                    JobTitle = "PRODTECH",
                    Date = day,
                    Total = prodtechTotal,
                    Assigned = prodtech,
                    Idle = prodtechTotal - prodtech
                };

                MppSummaryModel foremanModel = new MppSummaryModel
                {
                    JobTitle = "FOREMAN",
                    Date = day,
                    Total = foremanTotal,
                    Assigned = foreman,
                    Idle = foremanTotal - foreman
                };

                MppSummaryModel mechanicModel = new MppSummaryModel
                {
                    JobTitle = "MECHANIC",
                    Date = day,
                    Total = mechanicTotal,
                    Assigned = mechanic,
                    Idle = mechanicTotal - mechanic
                };

                MppSummaryModel electricianModel = new MppSummaryModel
                {
                    JobTitle = "ELECTRICIAN",
                    Date = day,
                    Total = electricianTotal,
                    Assigned = electrician,
                    Idle = electricianTotal - electrician
                };

                MppSummaryModel teamLeadModel = new MppSummaryModel
                {
                    JobTitle = "TEAMLEAD",
                    Date = day,
                    Total = teamleaderTotal,
                    Assigned = teamleader,
                    Idle = teamleaderTotal - teamleader
                };

                MppSummaryModel reliefModel = new MppSummaryModel
                {
                    JobTitle = "RELIEF",
                    Date = day,
                    Total = reliefTotal,
                    Assigned = relief,
                    Idle = reliefTotal - relief
                };

                MppSummaryModel supportModel = new MppSummaryModel
                {
                    JobTitle = "SUPPORT",
                    Date = day,
                    Total = supportTotal,
                    Assigned = support,
                    Idle = supportTotal - support
                };

                MppSummaryModel generalModel = new MppSummaryModel
                {
                    JobTitle = "GENERALWORKER",
                    Date = day,
                    Total = generalworkerTotal,
                    Assigned = generalworker,
                    Idle = generalworkerTotal - generalworker
                };

                MppSummaryModel otherModel = new MppSummaryModel
                {
                    JobTitle = "OTHER",
                    Date = day,
                    Total = otherTotal,
                    Assigned = other,
                    Idle = otherTotal - other
                };
                #endregion

                result.Add(prodTechModel);
                result.Add(foremanModel);
                result.Add(mechanicModel);
                result.Add(electricianModel);
                result.Add(teamLeadModel);
                result.Add(reliefModel);
                result.Add(supportModel);
                result.Add(generalModel);
                result.Add(otherModel);
            }

            return result;
        }

        private MppModel GetIndexModel()
        {
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.MachineList = DropDownHelper.BuildEmptyList();
            ViewBag.UserList = DropDownHelper.BuildEmptyList();
            ViewBag.JobTitleList = DropDownHelper.BuildEmptyList();
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupType(_refDetailAppService);
            ViewBag.StatusMppList = DropDownHelper.BindDropDownStatusMpp();
            ViewBag.UserLS1List = DropDownHelper.BuildEmptyList();
            ViewBag.UserLS2List = DropDownHelper.BuildEmptyList();

            MppModel model = new MppModel();
            model.Access = GetAccess(WebConstants.MenuSlug.MPP, _menuService);
            model.PcID = AccountProdCenterID;
            model.DepID = AccountDepartmentID;
            model.SubDepID = AccountLocationID;

            return model;
        }

        private bool ValidateEmpSkillAndUserMachine(MPPExcelModel mpp, List<UserMachineTypeModel> empSkillList, List<UserMachineModel> userMachineList, List<UserModel> userList, List<MachineModel> machineList, out string errorMessage)
        {
            errorMessage = string.Empty;

            if (!string.IsNullOrEmpty(mpp.EmpIdA))
            {
                var user = userList.Where(x => x.EmployeeID == mpp.EmpIdA).FirstOrDefault();
                if (user != null)
                {
                    string[] machines = mpp.EmpMachineA.Split(',');
                    foreach (var machineCode in machines)
                    {
                        var machine = machineList.Where(x => x.Code == machineCode.Trim()).FirstOrDefault();
                        if (machine != null)
                        {
                            var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
                            if (!uskills.Any(x => x.MachineTypeID == machine.MachineTypeID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdA + " in Employee Skill";
                                return false;
                            }

                            var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();
                            if (!umachines.Any(x => x.MachineID == machine.ID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdA + " in User Machines";
                                return false;
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(mpp.EmpIdB))
            {
                var user = userList.Where(x => x.EmployeeID == mpp.EmpIdB).FirstOrDefault();
                if (user != null)
                {
                    string[] machines = mpp.EmpMachineB.Split(',');
                    foreach (var machineCode in machines)
                    {
                        var machine = machineList.Where(x => x.Code == machineCode.Trim()).FirstOrDefault();
                        if (machine != null)
                        {
                            var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
                            if (!uskills.Any(x => x.MachineTypeID == machine.MachineTypeID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdB + " in Employee Skill";
                                return false;
                            }

                            var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();
                            if (!umachines.Any(x => x.MachineID == machine.ID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdB + " in User Machines";
                                return false;
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(mpp.EmpIdC))
            {
                var user = userList.Where(x => x.EmployeeID == mpp.EmpIdC).FirstOrDefault();
                if (user != null)
                {
                    string[] machines = mpp.EmpMachineC.Split(',');
                    foreach (var machineCode in machines)
                    {
                        var machine = machineList.Where(x => x.Code == machineCode.Trim()).FirstOrDefault();
                        if (machine != null)
                        {
                            var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
                            if (!uskills.Any(x => x.MachineTypeID == machine.MachineTypeID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdC + " in Employee Skill";
                                return false;
                            }

                            var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();
                            if (!umachines.Any(x => x.MachineID == machine.ID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdC + " in User Machines";
                                return false;
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(mpp.EmpIdD))
            {
                var user = userList.Where(x => x.EmployeeID == mpp.EmpIdD).FirstOrDefault();
                if (user != null)
                {
                    string[] machines = mpp.EmpMachineD.Split(',');
                    foreach (var machineCode in machines)
                    {
                        var machine = machineList.Where(x => x.Code == machineCode.Trim()).FirstOrDefault();
                        if (machine != null)
                        {
                            var uskills = empSkillList.Where(x => x.UserID == user.ID).ToList();
                            if (!uskills.Any(x => x.MachineTypeID == machine.MachineTypeID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdD + " in Employee Skill";
                                return false;
                            }

                            var umachines = userMachineList.Where(x => x.UserID == user.ID).ToList();
                            if (!umachines.Any(x => x.MachineID == machine.ID))
                            {
                                errorMessage = machineCode + " is not listed for " + mpp.EmpIdD + " in User Machines";
                                return false;
                            }
                        }
                    }
                }
            }

            return true;
        }

        private MppModel GetMpp(int id)
        {
            string mpp = _mppAppService.GetById(id, true);
            MppModel mppModel = mpp.DeserializeToMpp();

            return mppModel;
        }

        private List<long> GetParentIDList(long locID)
        {
            List<long> result = new List<long>();

            LocationModel model = GetLocation(locID);
            if (model.ParentID != 0)
            {
                result.Add(model.ID);
                model = GetLocation(model.ParentID);
                if (model.ParentID != 0)
                {
                    result.Add(model.ID);
                    model = GetLocation(model.ParentID);
                    if (model.ParentID != 0)
                    {
                        result.Add(model.ID);
                        model = GetLocation(model.ParentID);
                        result.Add(model.ID);
                    }
                    else
                    {
                        result.Add(model.ID);
                    }
                }
                else
                {
                    result.Add(model.ID);
                }
            }
            else
            {
                result.Add(model.ID);
            }

            return result;
        }

        private LocationModel GetLocation(long locationID)
        {
            string location = _locationAppService.GetById(locationID, true);
            LocationModel locationModel = location.DeserializeToLocation();

            return locationModel;
        }

        private bool AddUpdateDeleteMPP(List<MppModel> addResult, List<MppModel> updateResult, List<MppModel> deleteResult)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            string addMPP = "INSERT [dbo].[Mpps] ([StartDate], [EndDate], [Year], [Week], [Date], [Shift], [StatusMPP], [JobTitle], [EmployeeID], [EmployeeName], [EmployeeMachine], [LocationID], [Location], [GroupType], [GroupName], " +
                            "[IsDeleted], [ModifiedBy], [ModifiedDate], [WPPID]) VALUES(@StartDate, @EndDate, @Year, @Week, @Date, @Shift, @StatusMPP, @JobTitle, @EmployeeID, @EmployeeName, @EmployeeMachine, @LocationID, @Location, " +
                             "@GroupType, @GroupName, @IsDeleted, @ModifiedBy, @ModifiedDate, @WPPID)";
            string updateMPP = "UPDATE [dbo].[Mpps] SET [StartDate] = @StartDate, [EndDate] = @EndDate, [Year] = @Year, [Week] = @Week, [Date] = @Date, [Shift] = @Shift, [StatusMPP] = @StatusMPP, [JobTitle] = @JobTitle, [EmployeeID] = @EmployeeID, " +
                               "[EmployeeName] = @EmployeeName, [EmployeeMachine] = @EmployeeMachine, [LocationID] = @LocationID, [Location] = @Location, [GroupType] = @GroupType, [GroupName] = @GroupName, " +
                               "[ModifiedBy] = @ModifiedBy, [ModifiedDate] = @ModifiedDate WHERE([ID] = @ID)";
            string deleteMPP = "DELETE FROM [dbo].[Mpps] WHERE([ID] = @ID)";

            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    foreach (var mpp in addResult)
                    {
                        SqlCommand command = new SqlCommand(addMPP, connection, transaction);
                        command.Parameters.Add("@StartDate", SqlDbType.Date).Value = mpp.StartDate;
                        command.Parameters.Add("@EndDate", SqlDbType.Date).Value = mpp.EndDate;
                        command.Parameters.Add("@Year", SqlDbType.Int).Value = mpp.Year;
                        command.Parameters.Add("@Week", SqlDbType.Int).Value = mpp.Week;
                        command.Parameters.Add("@Date", SqlDbType.Date).Value = mpp.Date;
                        command.Parameters.Add("@Shift", SqlDbType.Char).Value = mpp.Shift;
                        command.Parameters.Add("@StatusMPP", SqlDbType.VarChar).Value = mpp.StatusMPP;
                        command.Parameters.Add("@JobTitle", SqlDbType.VarChar).Value = mpp.JobTitle;
                        command.Parameters.Add("@EmployeeID", SqlDbType.Char).Value = mpp.EmployeeID;
                        command.Parameters.Add("@EmployeeName", SqlDbType.VarChar).Value = mpp.EmployeeName;
                        command.Parameters.Add("@EmployeeMachine", SqlDbType.VarChar).Value = mpp.EmployeeMachine;
                        command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = mpp.LocationID;
                        command.Parameters.Add("@Location", SqlDbType.VarChar).Value = mpp.Location;
                        command.Parameters.Add("@GroupType", SqlDbType.VarChar).Value = mpp.GroupType;
                        command.Parameters.Add("@GroupName", SqlDbType.VarChar).Value = mpp.GroupName;
                        command.Parameters.Add("@IsDeleted", SqlDbType.Bit).Value = mpp.IsDeleted;
                        command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = mpp.ModifiedBy;
                        command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = mpp.ModifiedDate;
                        command.Parameters.Add("@WPPID", SqlDbType.BigInt).Value = mpp.WPPID;
                        command.ExecuteNonQuery();
                    }

                    foreach (var mpp in updateResult)
                    {
                        SqlCommand command = new SqlCommand(updateMPP, connection, transaction);
                        command.Parameters.Add("@StartDate", SqlDbType.Date).Value = mpp.StartDate;
                        command.Parameters.Add("@EndDate", SqlDbType.Date).Value = mpp.EndDate;
                        command.Parameters.Add("@Year", SqlDbType.Int).Value = mpp.Year;
                        command.Parameters.Add("@Week", SqlDbType.Int).Value = mpp.Week;
                        command.Parameters.Add("@Date", SqlDbType.Date).Value = mpp.Date;
                        command.Parameters.Add("@Shift", SqlDbType.Char).Value = mpp.Shift;
                        command.Parameters.Add("@StatusMPP", SqlDbType.VarChar).Value = mpp.StatusMPP;
                        command.Parameters.Add("@JobTitle", SqlDbType.VarChar).Value = mpp.JobTitle;
                        command.Parameters.Add("@EmployeeID", SqlDbType.Char).Value = mpp.EmployeeID;
                        command.Parameters.Add("@EmployeeName", SqlDbType.VarChar).Value = mpp.EmployeeName;
                        command.Parameters.Add("@EmployeeMachine", SqlDbType.VarChar).Value = mpp.EmployeeMachine;
                        command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = mpp.LocationID;
                        command.Parameters.Add("@Location", SqlDbType.VarChar).Value = mpp.Location;
                        command.Parameters.Add("@GroupType", SqlDbType.VarChar).Value = mpp.GroupType;
                        command.Parameters.Add("@GroupName", SqlDbType.VarChar).Value = mpp.GroupName;
                        command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = mpp.ModifiedBy;
                        command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = mpp.ModifiedDate;
                        command.Parameters.Add("@ID", SqlDbType.BigInt).Value = mpp.ID;
                        command.ExecuteNonQuery();
                    }

                    foreach (var mpp in deleteResult)
                    {
                        SqlCommand command = new SqlCommand(deleteMPP, connection, transaction);
                        command.Parameters.Add("@ID", SqlDbType.BigInt).Value = mpp.ID;
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

        private string GetJobTitle(long id)
        {
            if (id > 0)
            {
                string jt = _jobTitleAppService.GetById(id, true);
                if (string.IsNullOrEmpty(jt))
                    return null;

                return jt.DeserializeToJobTitle().Title;
            }

            return null;
        }

        private string GetMachineName(long id)
        {
            if (id > 0)
            {
                string machine = _machineAppService.GetById(id, true);
                if (string.IsNullOrEmpty(machine))
                    return null;

                return machine.DeserializeToMachine().Code;
            }

            return null;
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
        #endregion

        #region ::Helper::
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

        public int GetCurrentWeekNumber(DateTime dateTime)
        {
            var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dateTime, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weeknum;
        }

        public string GetEmployeeID(long id)
        {
            string res = _empAppService.GetEmployeeIdByID(id);
            return res;
        }

        public bool IsMachineExist(string machineCode, long locationId, List<MachineModel> machineModelList)
        {
            return machineModelList.Any(x => x.Code.Trim().ToLower() == machineCode.Trim().ToLower() && x.LocationID == locationId);
        }

        public bool IsEmployeeOnLeave(string empID, DateTime date, List<EmployeeLeaveModel> emLeaveModelList)
        {
            bool isLeave = false;

            var resultEm = emLeaveModelList.Where(x => x.EmployeeID.Trim() == empID.Trim() && x.StartDate >= date && x.EndDate <= date);

            if (resultEm.Count() > 0)
            {
                isLeave = true;
            }
            else
            {
                isLeave = false;
            }
            return isLeave;
        }

        public bool IsMultiMachine(string machineCode)
        {
            return machineCode.Contains(',');
        }
        #endregion
    }
}
