#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using System.Globalization;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.Script.Serialization;
#endregion


namespace Fast.Web.Controllers
{
    public class BaseController<TEntity> : Controller where TEntity : class
    {
        #region ::Account Properties::

        protected UserModel Account
        {
            get
            {
                return Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
            }
        }

        protected long AccountID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.ID;
            }
        }

        protected string AccountEmail
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.Email;
            }
        }

        protected string AccountRole
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.RoleName;
            }
        }

        protected List<string> AccountRoleList
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.RoleNames;
            }
        }

        protected string AccountJobTitle
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.JobTitle;
            }
        }

        protected string SupervisorEmail
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.SupervisorEmail;
            }
        }
        protected string AccountSpvName
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.SupervisorName;
            }
        }

        protected string AccountName
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.UserName;
            }
        }

        protected string AccountEmployeeID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.EmployeeID;
            }
        }

        protected string AccountSpvEmployeeID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.SupervisorID;
            }
        }


        protected long AccountSpvUserID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.SupervisorUserID;
            }
        }

        protected bool AccountIsAdmin
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.IsAdmin;
            }
        }

        protected long AccountLocationID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.LocationID.HasValue ? user.LocationID.Value : 0;
            }
        }

        protected string AccountLocation
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.Location;
            }
        }

        protected long AccountDepartmentID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.DepartmentID;
            }
        }

        protected long AccountProdCenterID
        {
            get
            {
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                return user.ProdCenterID;
            }
        }
        #endregion

        #region ::Get User Access::
        protected AccessRightDBModel GetAccess(string menuSlug, IMenuAppService _menuService)
        {
            AccessRightDBModel result = new AccessRightDBModel();
            List<AccessRightDBModel> accessModel = (List<AccessRightDBModel>)HttpContext.Session["AuthAccess"];

            if (AccountIsAdmin)
            {
                result.Read = true;
                result.Write = true;
                result.Print = true;
                result.IsAdmin = true;
            }
            else
            {
                if (accessModel != null)
                {
                    string usermenu = _menuService.GetBy("PageSlug", menuSlug, true);
                    MenuModel userMenuModel = usermenu.DeserializeToMenu();

                    var temp = accessModel.Where(x => x.MenuID == userMenuModel.ID).FirstOrDefault();
                    if (temp != null)
                    {
                        result = temp;
                    }
                }
            }

            return result;
        }
        #endregion

        #region ::Others::

        protected TEntity GetModel(string value)
        {
            return string.IsNullOrEmpty(value) ? null : JsonConvert.DeserializeObject<TEntity>(value);
        }

        protected HttpStatusCodeResult BadRequestResponse
        {
            get
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
        }

        protected void SetFalseTempData(string errorMessage)
        {
            TempData["Result"] = false;
            TempData["ErrorMessage"] = errorMessage;
        }

        protected void SetTrueTempData(string message)
        {
            TempData["Result"] = true;
            TempData["Message"] = message;
        }

        protected void GetTempData()
        {
            ViewBag.Result = TempData["Result"];
            ViewBag.ErrorMessage = TempData["ErrorMessage"];
            ViewBag.Message = TempData["Message"];
        }

        protected object BaseResult
        {
            get
            {
                return TempData["Result"];
            }
        }

        protected object BaseErrorMessage
        {
            get
            {
                return TempData["ErrorMessage"];
            }
        }

        protected string GetShift()
        {
            int result;
            if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 6 && Int32.Parse(DateTime.Now.Hour.ToString()) < 14)
            {
                result = 1;
            }
            else if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 14 && Int32.Parse(DateTime.Now.Hour.ToString()) < 22)
            {
                result = 2;
            }
            else
            {
                result = 3;
            }

            return result.ToString();
        }
        #endregion

        #region ::LPH Header::
        protected LPHHeaderModel LPHHeader(IWeeksAppService _weekService, IMppAppService _mppService, IWppAppService _wppService)
        {
            LPHHeaderModel result = new LPHHeaderModel();

            // set date
            result.Date = DateTime.Now;

            // set week
            //List<QueryFilter> weekFilter = new List<QueryFilter>();
            //weekFilter.Add(new QueryFilter("StartDate", result.Date.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
            //weekFilter.Add(new QueryFilter("EndDate", result.Date.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
            //string week = _weekService.Get(weekFilter);
            //WeeksModel weekModel = week.DeserializeToWeek();
            result.Week = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            // get shift based on current time
            if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 6 && Int32.Parse(DateTime.Now.Hour.ToString()) < 14)
            {
                result.Shift = 1;
                result.Start = 6;
                result.Stop = 14;
            }
            else if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 14 && Int32.Parse(DateTime.Now.Hour.ToString()) < 22)
            {
                result.Shift = 2;
                result.Start = 14;
                result.Stop = 22;
            }
            else
            {
                result.Shift = 3;
                result.Start = 22;
                result.Stop = 6;
                if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 0 && Int32.Parse(DateTime.Now.Hour.ToString()) < 6)
                {
                    result.Date = DateTime.Now.AddDays(-1);
                }

            }

            // set machine from MPP
            List<QueryFilter> mppFilter = new List<QueryFilter>();
            mppFilter.Add(new QueryFilter("Date", result.Date.ToString("yyyy-MM-dd")));
            mppFilter.Add(new QueryFilter("Week", result.Week));
            mppFilter.Add(new QueryFilter("Shift", result.Shift.ToString()));
            mppFilter.Add(new QueryFilter("LocationID", AccountLocationID.ToString()));

            //mppFilter.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString(), Operator.Equals, Operation.Or));
            //mppFilter.Add(new QueryFilter("LocationID", AccountProdCenterID.ToString(), Operator.Equals, Operation.Or));
            try
            {
                string mpps = _mppService.Find(mppFilter);
                List<MppModel> mppModelList = mpps.DeserializeToMppList();

                MppModel temp_selected = mppModelList.Where(x => (x.JobTitle == "PRODTECH" || x.JobTitle == "GENERALWORKER") && x.EmployeeID == AccountEmployeeID).FirstOrDefault();

                if (temp_selected != null && mppModelList.Count > 0)
                {
                    mppModelList = mppModelList.Where(x => x.EmployeeMachine == temp_selected.EmployeeMachine).ToList();

                    // set group
                    result.Group = mppModelList[0].GroupName;
                    MppModel prodtech = mppModelList.Where(x => x.JobTitle == "PRODTECH").FirstOrDefault();
                    MppModel foreman = mppModelList.Where(x => x.JobTitle == "FOREMAN").FirstOrDefault();
                    MppModel mechanic = mppModelList.Where(x => x.JobTitle == "MECHANIC").FirstOrDefault();
                    MppModel electrician = mppModelList.Where(x => x.JobTitle == "ELECTRICIAN" || x.JobTitle == "ELECTRIC").FirstOrDefault();
                    MppModel teamleader = mppModelList.Where(x => x.JobTitle == "TEAMLEADER" || x.JobTitle == "SUPERVISOR" || x.JobTitle == "FOREMAN").FirstOrDefault();
                    MppModel relief = mppModelList.Where(x => x.JobTitle == "RELIEF").FirstOrDefault();
                    MppModel support = mppModelList.Where(x => x.JobTitle == "SUPPORT").FirstOrDefault();
                    MppModel generalworker = mppModelList.Where(x => x.JobTitle == "GENERALWORKER" || x.JobTitle == "GW").FirstOrDefault();
                    MppModel other = mppModelList.Where(x => x.JobTitle == "OTHER").FirstOrDefault();

                    // set prodtech mechanic electrician
                    result.ProdTech = prodtech == null ? string.Empty : prodtech.EmployeeID;
                    result.Foreman = foreman == null ? string.Empty : foreman.EmployeeID;
                    result.Mechanic = mechanic == null ? string.Empty : mechanic.EmployeeID;
                    result.Electrician = electrician == null ? string.Empty : electrician.EmployeeID;
                    result.TeamLeader = teamleader == null ? string.Empty : teamleader.EmployeeID;
                    result.Relief = relief == null ? string.Empty : relief.EmployeeID;
                    result.Support = support == null ? string.Empty : support.EmployeeID;
                    result.GeneralWorker = generalworker == null ? string.Empty : generalworker.EmployeeID;
                    result.Other = other == null ? string.Empty : other.EmployeeID;

                    result.Machines = new List<string>();
                    ////machine from mpps BATAL
                    //foreach (var mpp in mppModelList)
                    //{
                    //	if (mpp.EmployeeMachine != null && mpp.EmployeeMachine.Trim() != "")
                    //	{
                    //		var TagIds = mpp.EmployeeMachine.Split(',');

                    //		foreach (var machine in TagIds)
                    //		{
                    //			result.Machines.Add(machine);
                    //		}
                    //	}

                    //}
                }

                // set brand from WPP
                //location level
                List<QueryFilter> wppFilter = new List<QueryFilter>();
                wppFilter.Add(new QueryFilter("Date", result.Date.ToString("yyyy-MM-dd")));
                wppFilter.Add(new QueryFilter("LocationID", AccountLocationID.ToString()));

                string wpp = _wppService.Find(wppFilter);
                List<WppModel> wppModel = wpp.DeserializeToWppList();
                foreach (var dwpp in wppModel)
                {
                    if (!string.IsNullOrEmpty(dwpp.Brand))
                        result.Brands.Add(dwpp.Brand);

                    if (!string.IsNullOrEmpty(dwpp.Maker))
                        result.Machines.Add(dwpp.Maker);

                    if (!string.IsNullOrEmpty(dwpp.Packer))
                        result.Machines.Add(dwpp.Packer);
                }

                if (AccountLocationID != AccountDepartmentID)
                {
                    //department level
                    wppFilter = new List<QueryFilter>();
                    wppFilter.Add(new QueryFilter("Date", result.Date.ToString("yyyy-MM-dd")));
                    wppFilter.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                    wpp = _wppService.Find(wppFilter);
                    wppModel = wpp.DeserializeToWppList();
                    foreach (var dwpp in wppModel)
                    {
                        if (!string.IsNullOrEmpty(dwpp.Brand))
                            result.Brands.Add(dwpp.Brand);

                        if (!string.IsNullOrEmpty(dwpp.Maker))
                            result.Machines.Add(dwpp.Maker);

                        if (!string.IsNullOrEmpty(dwpp.Packer))
                            result.Machines.Add(dwpp.Packer);
                    }
                }

                if (AccountLocationID != AccountProdCenterID)
                {
                    //prod center level
                    wppFilter = new List<QueryFilter>();
                    wppFilter.Add(new QueryFilter("Date", result.Date.ToString("yyyy-MM-dd")));
                    wppFilter.Add(new QueryFilter("LocationID", AccountProdCenterID.ToString()));
                    wpp = _wppService.Find(wppFilter);
                    wppModel = wpp.DeserializeToWppList();
                    foreach (var dwpp in wppModel)
                    {
                        if (!string.IsNullOrEmpty(dwpp.Brand))
                            result.Brands.Add(dwpp.Brand);

                        if (!string.IsNullOrEmpty(dwpp.Maker))
                            result.Machines.Add(dwpp.Maker);

                        if (!string.IsNullOrEmpty(dwpp.Packer))
                            result.Machines.Add(dwpp.Packer);
                    }
                }

                result.Brands = result.Brands.Distinct().ToList();
                result.Machines = result.Machines.Distinct().ToList();

                return result;
            }
            catch (Exception e)
            {
                var lala = e.InnerException;
                var yeye = e;
                return result;
            }
        }
        #endregion
        protected void LoadFlag(long LPHID, int flag)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    SqlCommand command = new SqlCommand("UPDATE LPHSubmissions SET Flag = @Flag WHERE LPHID = @LPHID ;", connection, transaction);
                    command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = LPHID;
                    command.Parameters.Add("@Flag", SqlDbType.Int).Value = flag;
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " RandomFlag: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                    transaction.Rollback();
                }
            }

        }

        protected void LoadFlagPrimary(long LPHID, int flag)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    SqlCommand command = new SqlCommand("UPDATE PPLPHSubmissions SET Flag = @Flag WHERE LPHID = @LPHID ;", connection, transaction);
                    command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = LPHID;
                    command.Parameters.Add("@Flag", SqlDbType.Int).Value = flag;
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " RandomFlagPP: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                    transaction.Rollback();
                }
            }

        }

        #region ::Load LPH Header SP::
        protected void LoadLPHHeader(
            LPHHeaderModel model,
            ILocationAppService _locationAppService,
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            IUserRoleAppService _userRoleAppService,
            IMachineAppService _machineAppService,
            IBrandAppService _brandAppService,
            IJobTitleAppService _jobTitleAppService,
            IWppAppService _wppAppService,
            IReferenceAppService _referenceAppService = null,
            int flagdef = 0,
            long locationID = 0
            )
        {
            var ProdCenter = AccountProdCenterID;
            var Department = AccountDepartmentID;
            if (locationID != 0)
            {
                string location = _locationAppService.GetById(locationID);
                LocationModel locationModel = location.DeserializeToLocation();
                if (locationModel.ParentID == 1)
                {
                    ProdCenter = locationModel.ID;
                }
                else
                {
                    string loc = _locationAppService.GetById(locationModel.ParentID);
                    LocationModel locModel = loc.DeserializeToLocation();
                    if (locModel.ParentID == 1)
                    {
                        Department = locationModel.ID;
                        ProdCenter = locModel.ID;
                    }
                    else
                    {
                        string dep = _locationAppService.GetById(locModel.ID);
                        LocationModel depModel = dep.DeserializeToLocation();
                        Department = depModel.ID;
                        ProdCenter = depModel.ParentID;
                    }
                }
            }

            List<long> locationPcIdList = _locationAppService.GetLocIDListByLocType(ProdCenter, "productioncenter");
            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            if (flagdef == 1) //maker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, true, false);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 2) //packer
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 3) //filter
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "FT");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 3).OrderBy(x => x.Code);
            }
            else if (flagdef == 4) //casepacker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "CP");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 5) //ripper
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "RI");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0).OrderBy(x => x.Code);
            }
            else if (flagdef == 6) //laser
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "LS");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0).OrderBy(x => x.Code);
            }
            else if (flagdef == 7) //menthol
            {
                //FA saja yg muncul , dua param terakhir getbrandlist (true)
                //nonFA saja yg muncul , dua param terakhir getbrandlist (false, 7)
                //all code muncul, dua param terakhir getbrandlist (false, 0)
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "MT");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 7).OrderBy(x => x.Code);
            }
            else if (flagdef == 8) //robot paletizer
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "RP");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0).OrderBy(x => x.Code);
            }
            else if (flagdef == 9) //tdcpacker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "TD");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 10) //gw general
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model.Machines, false, false);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0).OrderBy(x => x.Code);
            }
            else //other XXX
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model.Machines, false, false, "XXX");
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0).OrderBy(x => x.Code);
            }

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            string vProdtech = ConfigurationManager.AppSettings["Role1"].ToString();
            string vMechanic = ConfigurationManager.AppSettings["Role2"].ToString();
            string vElectrician = ConfigurationManager.AppSettings["Role3"].ToString();
            string vSupervisor = ConfigurationManager.AppSettings["Role4"].ToString();
            string vGW = ConfigurationManager.AppSettings["Role5"].ToString();


            // chanif: tak buat null semua second paramnya, biar list nya tetap muncul
            List<string> electricianPositionList = jobTitleList.Where(x => x.RoleName.Equals(vElectrician, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.ElectricSelected = String.IsNullOrEmpty(model.Electrician) ? "" : model.Electrician;
            ViewBag.Electrics = LPHSPHelper.GetElectricianList(employeeList, "", electricianPositionList, userList, locationIdList, extraEmployeeRoles);

            ViewBag.ProdtechSelected = String.IsNullOrEmpty(model.ProdTech) ? "" : model.ProdTech;
            List<string> prodTechPositionList = jobTitleList.Where(x => x.RoleName.Equals(vProdtech, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.Prodtechs = LPHSPHelper.GetProdTechList(employeeList, "", prodTechPositionList, userList, locationIdList, extraEmployeeRoles);

            ViewBag.ForemanSelected = String.IsNullOrEmpty(model.Foreman) ? "" : model.Foreman;
            ViewBag.Foremans = LPHSPHelper.GetForemanList(employeeList, "", userList, locationIdList, extraEmployeeRoles);

            List<string> mechanicPositionList = jobTitleList.Where(x => x.RoleName.Equals(vMechanic, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.MechanicSelected = String.IsNullOrEmpty(model.Mechanic) ? "" : model.Mechanic;
            ViewBag.Mechanics = LPHSPHelper.GetMechanicList(employeeList, "", mechanicPositionList, userList, locationIdList, extraEmployeeRoles);

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals(vSupervisor, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.LeaderSelected = String.IsNullOrEmpty(model.TeamLeader) ? "" : model.TeamLeader;
            ViewBag.Leaders = LPHSPHelper.GetTeamLeadList(employeeList, "", leaderPositionList, userList, locationIdList, extraEmployeeRoles);

            List<string> gwPositionList = jobTitleList.Where(x => x.RoleName.Equals(vGW, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.GeneralWorkerSelected = String.IsNullOrEmpty(model.GeneralWorker) ? "" : model.GeneralWorker;
            ViewBag.GeneralWorker = LPHSPHelper.GetGeneralWorkerList(employeeList, "", gwPositionList, userList, locationIdList, extraEmployeeRoles);

            string daysAgo = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd");
            string daysLater = DateTime.Now.AddDays(3).ToString("yyyy-MM-dd");

            List<QueryFilter> wppFilter = new List<QueryFilter>();
            wppFilter.Add(new QueryFilter("Date", daysLater, Operator.LessThanOrEqualTo));
            wppFilter.Add(new QueryFilter("Date", daysAgo, Operator.GreaterThanOrEqual));

            //PO Number
            string wppString = _wppAppService.Find(wppFilter);
            List<WppModel> wppModels = wppString.DeserializeToWppList();
            wppModels = wppModels.Where(x => !string.IsNullOrEmpty(x.PONumber) && locationPcIdList.Any(y => y == x.LocationID)).Distinct().ToList();
            wppModels = wppModels.GroupBy(x => x.PONumber.TrimEnd()).Select(y => y.First()).OrderBy(x => x.PONumber).ToList();

            ViewBag.WPP = wppModels;



            if (_referenceAppService != null)
                ViewBag.MachineTypes = LPHSPHelper.GetMachineTypeList(_machineAppService, _referenceAppService, locationPcIdList, model.Machines);
        }
        #endregion

        #region ::Load LPH Header Approval SP::
        protected void LoadLPHHeaderApproval(
            LPHHeaderModel model,
            ILocationAppService _locationAppService,
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            IUserRoleAppService _userRoleAppService,
            IMachineAppService _machineAppService,
            IBrandAppService _brandAppService,
            IJobTitleAppService _jobTitleAppService,
            IWppAppService _wppAppService,
            IReferenceAppService _referenceAppService = null,
            int flagdef = 0,
            long locationID = 0
            )
        {
            var ProdCenter = AccountProdCenterID;
            var Department = AccountDepartmentID;
            if (locationID != 0)
            {
                string location = _locationAppService.GetById(locationID);
                LocationModel locationModel = location.DeserializeToLocation();
                if (locationModel.ParentID == 1)
                {
                    ProdCenter = locationModel.ID;
                }
                else
                {
                    string loc = _locationAppService.GetById(locationModel.ParentID);
                    LocationModel locModel = loc.DeserializeToLocation();
                    if (locModel.ParentID == 1)
                    {
                        Department = locationModel.ID;
                        ProdCenter = locModel.ID;
                    }
                    else
                    {
                        string dep = _locationAppService.GetById(locModel.ID);
                        LocationModel depModel = dep.DeserializeToLocation();
                        Department = depModel.ID;
                        ProdCenter = depModel.ParentID;
                    }
                }
            }

            List<long> locationPcIdList = _locationAppService.GetLocIDListByLocType(ProdCenter, "productioncenter");
            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            if (flagdef == 1) //maker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, true, false, "MK", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 2) //packer
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, true, "PC", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 3) //filter
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "FT", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 3, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 4) //casepacker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "CP", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 5) //ripper
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "RI", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 6) //laser
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "LS", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 7) //menthol
            {
                //FA saja yg muncul , dua param terakhir getbrandlist (true)
                //nonFA saja yg muncul , dua param terakhir getbrandlist (false, 7)
                //all code muncul, dua param terakhir getbrandlist (false, 0)
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "MT", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 7, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 8) //robot paletizer
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "RP", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 9) //tdcpacker
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationPcIdList, model.Machines, false, false, "TD", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, true, 0, true).OrderBy(x => x.Code);
            }
            else if (flagdef == 10) //gw general
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model.Machines, false, false, "", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0, true).OrderBy(x => x.Code);
            }
            else //other XXX
            {
                ViewBag.Machines = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model.Machines, false, false, "XXX", true);
                ViewBag.Brand = LPHSPHelper.GetBrandList(_brandAppService, locationPcIdList, model.Brands, false, 0, true).OrderBy(x => x.Code);
            }

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            string vProdtech = ConfigurationManager.AppSettings["Role1"].ToString();
            string vMechanic = ConfigurationManager.AppSettings["Role2"].ToString();
            string vElectrician = ConfigurationManager.AppSettings["Role3"].ToString();
            string vSupervisor = ConfigurationManager.AppSettings["Role4"].ToString();
            string vGW = ConfigurationManager.AppSettings["Role5"].ToString();


            // chanif: tak buat null semua second paramnya, biar list nya tetap muncul
            List<string> electricianPositionList = jobTitleList.Where(x => x.RoleName.Equals(vElectrician, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.ElectricSelected = String.IsNullOrEmpty(model.Electrician) ? "" : model.Electrician;
            ViewBag.Electrics = LPHSPHelper.GetElectricianList(employeeList, "", electricianPositionList, userList, locationIdList, extraEmployeeRoles, true);

            ViewBag.ProdtechSelected = String.IsNullOrEmpty(model.ProdTech) ? "" : model.ProdTech;
            List<string> prodTechPositionList = jobTitleList.Where(x => x.RoleName.Equals(vProdtech, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.Prodtechs = LPHSPHelper.GetProdTechList(employeeList, "", prodTechPositionList, userList, locationIdList, extraEmployeeRoles, true);

            ViewBag.ForemanSelected = String.IsNullOrEmpty(model.Foreman) ? "" : model.Foreman;
            ViewBag.Foremans = LPHSPHelper.GetForemanList(employeeList, "", userList, locationIdList, extraEmployeeRoles);

            List<string> mechanicPositionList = jobTitleList.Where(x => x.RoleName.Equals(vMechanic, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.MechanicSelected = String.IsNullOrEmpty(model.Mechanic) ? "" : model.Mechanic;
            ViewBag.Mechanics = LPHSPHelper.GetMechanicList(employeeList, "", mechanicPositionList, userList, locationIdList, extraEmployeeRoles, true);

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals(vSupervisor, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.LeaderSelected = String.IsNullOrEmpty(model.TeamLeader) ? "" : model.TeamLeader;
            ViewBag.Leaders = LPHSPHelper.GetTeamLeadList(employeeList, "", leaderPositionList, userList, locationIdList, extraEmployeeRoles, true);

            List<string> gwPositionList = jobTitleList.Where(x => x.RoleName.Equals(vGW, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.GeneralWorkerSelected = String.IsNullOrEmpty(model.GeneralWorker) ? "" : model.GeneralWorker;
            ViewBag.GeneralWorker = LPHSPHelper.GetGeneralWorkerList(employeeList, "", gwPositionList, userList, locationIdList, extraEmployeeRoles, true);

            string daysAgo = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd");
            string daysLater = DateTime.Now.AddDays(3).ToString("yyyy-MM-dd");

            List<QueryFilter> wppFilter = new List<QueryFilter>();
            wppFilter.Add(new QueryFilter("Date", daysLater, Operator.LessThanOrEqualTo));
            wppFilter.Add(new QueryFilter("Date", daysAgo, Operator.GreaterThanOrEqual));

            //PO Number
            string wppString = _wppAppService.Find(wppFilter);
            List<WppModel> wppModels = wppString.DeserializeToWppList();
            //old code
            //wppModels = wppModels.Where(x => !string.IsNullOrEmpty(x.PONumber) && locationPcIdList.Any(y => y == x.LocationID)).Distinct().ToList();

            //nolocation needed for approval
            wppModels = wppModels.Where(x => !string.IsNullOrEmpty(x.PONumber)).Distinct().ToList();
            wppModels = wppModels.GroupBy(x => x.PONumber.TrimEnd()).Select(y => y.First()).OrderBy(x => x.PONumber).ToList();

            ViewBag.WPP = wppModels;

            if (_referenceAppService != null)
                ViewBag.MachineTypes = LPHSPHelper.GetMachineTypeList(_machineAppService, _referenceAppService, locationPcIdList, model.Machines);
        }
        #endregion


        #region ::Load LPH Header PP::
        protected void LoadLPHHeaderPP(
            ILocationAppService _locationAppService,
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            IUserRoleAppService _userRoleAppService,
            IJobTitleAppService _jobTitleAppService,
            long locationID = 0
            )
        {
            var ProdCenter = AccountProdCenterID;
            var Department = AccountDepartmentID;
            if (locationID != 0)
            {
                string location = _locationAppService.GetById(locationID);
                LocationModel locationModel = location.DeserializeToLocation();
                if (locationModel.ParentID == 1)
                {
                    ProdCenter = locationModel.ID;
                }
                else
                {
                    string loc = _locationAppService.GetById(locationModel.ParentID);
                    LocationModel locModel = loc.DeserializeToLocation();
                    if (locModel.ParentID == 1)
                    {
                        Department = locationModel.ID;
                        ProdCenter = locModel.ID;
                    }
                    else
                    {
                        string dep = _locationAppService.GetById(locModel.ID);
                        LocationModel depModel = dep.DeserializeToLocation();
                        Department = depModel.ID;
                        ProdCenter = depModel.ParentID;
                    }
                }
            }

            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            string vProdtech = ConfigurationManager.AppSettings["Role1"].ToString();
            string vSupervisor = ConfigurationManager.AppSettings["Role4"].ToString();

            List<string> prodTechPositionList = jobTitleList.Where(x => x.RoleName.Equals(vProdtech, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.Prodtechs = LPHSPHelper.GetProdTechList(employeeList, null, prodTechPositionList, userList, locationIdList, extraEmployeeRoles);

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals(vSupervisor, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            var TL = LPHSPHelper.GetTeamLeadList(employeeList, null, leaderPositionList, userList, locationIdList, extraEmployeeRoles);
            var Fore = LPHSPHelper.GetForemanList(employeeList, null, userList, locationIdList, extraEmployeeRoles);
            ViewBag.Leaders = TL.Concat(Fore).Distinct().ToList();

            if (ProdCenter == 4 || ProdCenter == 5)
            {
                ViewBag.isWest = 1;
            }
            else
            {
                ViewBag.isWest = 0;
            }
        }
        #endregion

        #region ::Load LPH Header Approval PP::
        protected void LoadLPHHeaderPPApproval(
            ILocationAppService _locationAppService,
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            IUserRoleAppService _userRoleAppService,
            IJobTitleAppService _jobTitleAppService,
            long locationID = 0
            )
        {
            var ProdCenter = AccountProdCenterID;
            var Department = AccountDepartmentID;
            if (locationID != 0)
            {
                string location = _locationAppService.GetById(locationID);
                LocationModel locationModel = location.DeserializeToLocation();
                if (locationModel.ParentID == 1)
                {
                    ProdCenter = locationModel.ID;
                }
                else
                {
                    string loc = _locationAppService.GetById(locationModel.ParentID);
                    LocationModel locModel = loc.DeserializeToLocation();
                    if (locModel.ParentID == 1)
                    {
                        Department = locationModel.ID;
                        ProdCenter = locModel.ID;
                    }
                    else
                    {
                        string dep = _locationAppService.GetById(locModel.ID);
                        LocationModel depModel = dep.DeserializeToLocation();
                        Department = depModel.ID;
                        ProdCenter = depModel.ParentID;
                    }
                }
            }

            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            string vProdtech = ConfigurationManager.AppSettings["Role1"].ToString();
            string vSupervisor = ConfigurationManager.AppSettings["Role4"].ToString();

            List<string> prodTechPositionList = jobTitleList.Where(x => x.RoleName.Equals(vProdtech, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            ViewBag.Prodtechs = LPHSPHelper.GetProdTechList(employeeList, null, prodTechPositionList, userList, locationIdList, extraEmployeeRoles, true);

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals(vSupervisor, StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            var TL = LPHSPHelper.GetTeamLeadList(employeeList, null, leaderPositionList, userList, locationIdList, extraEmployeeRoles, true);
            var Fore = LPHSPHelper.GetForemanList(employeeList, null, userList, locationIdList, extraEmployeeRoles, true);
            ViewBag.Leaders = TL.Concat(Fore).Distinct().ToList();

            if (ProdCenter == 4 || ProdCenter == 5)
            {
                ViewBag.isWest = 1;
            }
            else
            {
                ViewBag.isWest = 0;
            }
        }
        #endregion

        #region ::Create LPH SP::
        protected long CreateLPH(
            ILPHAppService _lphAppService,
            ILPHLocationsAppService _lphLocationsAppService,
            ILPHExtrasAppService _lphExtrasAppService,
            ILPHSubmissionsAppService _lphSubmissionsAppService,
            ILPHApprovalsAppService _lphApprovalAppService,
            ILPHComponentsAppService _lphComponentsAppService,
            ILPHValuesAppService _lphValuesAppService,
            List<LPHExtrasModel> detailExtras,
            LPHSubmissionsModel detailSubmission,
            List<LPHComponentsModel> detailComponent,
            List<LPHValuesModel> detailValue,
            string controllerName)
        {
            #region Session
            UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
            long userAccountID = user.ID;
            string userEmpID = user.EmployeeID;
            long userSupervisorUserID = user.SupervisorUserID;
            string userSupervisorEmpID = user.SupervisorID;
            long userLocationID = user.LocationID.HasValue ? user.LocationID.Value : 0;
            string userAccountName = user.UserName;
            #endregion

            #region LPH			
            LPHModel model = new LPHModel
            {
                MenuTitle = controllerName,
                Header = controllerName,
                Type = "SP",
                LocationID = userLocationID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                IsDeleted = false // chanif: karena create awal sudah pasti jadi draft; edit: jadi acuan untuk lihat lph dihapus/atau tidak
            };

            long lphID = _lphAppService.AddModel(model);
            #endregion

            #region LPH Locations			
            LPHLocationsModel lphLocModel = new LPHLocationsModel
            {
                LPHID = lphID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                LocationID = userLocationID
            };

            _lphLocationsAppService.AddModel(lphLocModel);
            #endregion

            #region LPH Extras			
            if (detailExtras != null)
            {
                foreach (var item in detailExtras)
                {
                    if (item.ValueType.Contains("ImageURL"))
                        item.Value = GetValueOnCreate(controllerName, item);
                    item.LPHID = lphID;
                    item.Date = DateTime.Now;
                    item.UserID = userAccountID;
                    item.ModifiedBy = userAccountName;
                    item.ModifiedDate = DateTime.Now;
                    item.LocationID = userLocationID;
                    item.SubShift = 1;
                }

                if (detailExtras.Count > 0)
                {
                    string newExtras = JsonHelper<LPHExtrasModel>.Serialize(detailExtras);
                    _lphExtrasAppService.AddRange(newExtras);
                }
            }
            #endregion

            #region LPH Submissions			
            DateTime tanggal = DateTime.Now;
            DateTime yesterday = tanggal.AddDays(-1);

            string submit = _lphSubmissionsAppService.FindBy("LocationID", userLocationID, true);
            List<LPHSubmissionsModel> submitModel = submit.DeserializeToLPHSubmissionsList();
            submitModel = submitModel.Where(x => x.LPHHeader == model.Header && x.Shift.Trim() == detailSubmission.Shift && x.ModifiedDate > tanggal.AddHours(-24) && x.Date > yesterday).OrderByDescending(x => x.ID).ToList();

            int subshift = 1;
            if (submitModel != null && submitModel.Count() > 0)
            {
                subshift = int.Parse(submitModel[0].SubShift.ToString()) + 1;
            }

            LPHSubmissionsModel lphSubmissionsModel = new LPHSubmissionsModel
            {
                LPHID = lphID,
                LPHHeader = model.Header,
                Date = DateTime.Now,
                Shift = detailSubmission.Shift,
                SubShift = subshift,
                UserID = userAccountID,
                LocationID = userLocationID,
                StartTime = detailSubmission.StartTime.ToString(),
                EndTime = detailSubmission.EndTime.ToString(),
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                IsDeleted = true // chanif: karena create awal sudah pasti jadi draft. supaya tidak muncul di history dan approval, isdeleted harus true
            };

            var submissionID = _lphSubmissionsAppService.AddModel(lphSubmissionsModel);
            #endregion

            #region LPH Approvals
            LPHApprovalsModel modelApp = new LPHApprovalsModel
            {
                LPHSubmissionID = submissionID,
                UserID = userAccountID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                Status = "Draft",
                Date = DateTime.Now,
                LocationID = userLocationID,
                ApproverID = userSupervisorUserID,
                ApproverEmployeeID = userSupervisorEmpID,
                UserEmployeeID = userEmpID,
                Shift = detailSubmission.Shift
            };

            var appID = _lphApprovalAppService.AddModel(modelApp);
            if (appID == 0)
            {
                // recreate using ado
                //ExecuteQuery("INSERT INTO LPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status,ApproverID,ApproverEmployeeID) VALUES ('"
                //    + submissionID + "','" + DateTime.Now + "','" + detailSubmission.Shift + "','" + userAccountID + "','" + userEmpID + "','" + userLocationID + "','" + "Draft" + "','" + userSupervisorUserID + "','" + userSupervisorEmpID + "');");
                ExecuteQuery(string.Format("INSERT INTO LPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status,ApproverID,ApproverEmployeeID) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','Draft','{6}','{7}');", submissionID, DateTime.Now, detailSubmission.Shift, userAccountID, userEmpID, userLocationID, userSupervisorUserID, userSupervisorEmpID));
            }
            #endregion

            #region LPH Components			
            if (detailComponent != null)
            {
                // update fields
                detailComponent.ForEach(x => x.LPHID = lphID);
                detailComponent.ForEach(x => x.ModifiedBy = userAccountName);
                detailComponent.ForEach(x => x.ModifiedDate = DateTime.Now);

                // bulk insert
                _lphComponentsAppService.AddRangeModel(detailComponent);

                // get data with ID
                detailComponent = _lphComponentsAppService.FindBy("LPHID", lphID, true).DeserializeToLPHComponentList();
            }
            #endregion

            #region LPH Values
            if (detailComponent.Count() == detailValue.Count())
            {
                List<LPHValuesModel> valueList = new List<LPHValuesModel>();

                for (int i = 0; i < detailComponent.Count(); i++)
                {
                    LPHValuesModel modelValue = new LPHValuesModel();
                    modelValue.Value = GetValue(controllerName, detailValue[i]);
                    modelValue.ValueType = detailValue[i].ValueType;
                    modelValue.LPHComponentID = detailComponent[i].ID;
                    modelValue.ValueDate = DateTime.Now;
                    modelValue.SubmissionID = submissionID;
                    modelValue.ModifiedBy = userAccountName;
                    modelValue.ModifiedDate = DateTime.Now;

                    valueList.Add(modelValue);
                }

                _lphValuesAppService.AddRangeModel(valueList);
            }
            #endregion

            return lphID;
        }
        #endregion

        #region ::Edit LPH SP::
        protected async Task EditLPH(
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            ILPHSubmissionsAppService _lphSubmissionsAppService,
            ILPHExtrasAppService _lphExtrasAppService,
            ILPHComponentsAppService _lphComponentsAppService,
            ILPHValuesAppService _lphValuesAppService,
            ILPHValueHistoriesAppService _lphValueHistoriesAppService,
            ILPHApprovalsAppService _lphApprovalAppService,
            long id,
            List<LPHExtrasModel> detailExtras,
            List<LPHValuesModel> detailValue,
            string controllerName,
            int isSubmit = 0, int Special = 0)
        {
            LPHSubmissionsModel submitModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();

            if (submitModel != null && submitModel.IsDeleted && submitModel.Flag == Special || Special == 0) //kalo masih draft baru bisa submit, kalo ngga, dia akan keluar
            {
                #region ::Session::
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                long userAccountID = user.ID;
                string userEmpID = user.EmployeeID;
                long userSupervisorUserID = user.SupervisorUserID;
                string userSupervisorEmpID = user.SupervisorID;
                long userLocationID = user.LocationID.HasValue ? user.LocationID.Value : 0;
                string userAccountName = user.UserName;
                #endregion

                #region ::Update Approval::
                //fery optimizing
                string approval = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID);
                List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().ToList();
                approvalList = approvalList.Where(x => x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
                LPHApprovalsModel apprModel = approvalList.FirstOrDefault();

                if (isSubmit == 1 && submitModel != null && submitModel.IsDeleted)
                {
                    submitModel.IsDeleted = false;
                    _lphSubmissionsAppService.UpdateModel(submitModel);

                    //Do update in approval with status 'submitted' 
                    LPHApprovalsModel approveModel = _lphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID, true).DeserializeToLPHApproval();
                    if (approveModel != null)
                    {

                        approveModel.Status = "Submitted";
                        approveModel.ModifiedBy = AccountName;
                        approveModel.ModifiedDate = DateTime.Now;
                        approveModel.Notes = "";

                        string dataAppSub = JsonHelper<LPHApprovalsModel>.Serialize(approveModel);
                        _lphApprovalAppService.Update(dataAppSub);
                    }
                    else
                    {
                        approveModel = new LPHApprovalsModel
                        {
                            LPHSubmissionID = submitModel.ID,
                            UserID = userAccountID,
                            ModifiedBy = userAccountName,
                            ModifiedDate = DateTime.Now,
                            Status = "Submitted",
                            Date = DateTime.Now,
                            LocationID = userLocationID,
                            ApproverID = userSupervisorUserID,
                            ApproverEmployeeID = userSupervisorEmpID,
                            UserEmployeeID = userEmpID,
                            Shift = submitModel.Shift,
                            Notes = ""
                        };

                        _lphApprovalAppService.AddModel(approveModel);
                    }
                }
                else
                {
                    if (apprModel == null)
                    {
                        apprModel = new LPHApprovalsModel
                        {
                            LPHSubmissionID = submitModel.ID,
                            UserID = userAccountID,
                            ModifiedBy = userAccountName,
                            ModifiedDate = DateTime.Now,
                            Status = "Submitted",
                            Date = DateTime.Now,
                            LocationID = userLocationID,
                            ApproverID = userSupervisorUserID,
                            ApproverEmployeeID = userSupervisorEmpID,
                            UserEmployeeID = userEmpID,
                            Shift = submitModel.Shift,
                            Notes = ""
                        };

                        _lphApprovalAppService.AddModel(apprModel);
                    }
                }
                #endregion

                #region ::Detail Extras::
                //DELETE LPHExtra secepat kilat            						
                ExecuteQuery(string.Format("UPDATE LPHExtras SET IsDeleted = 1 WHERE LPHID = {0};", id));

                if (detailExtras != null)
                {
                    foreach (var item in detailExtras)
                    {
                        if (item.ValueType.Contains("ImageURL"))
                            item.Value = GetValueOnEdit(controllerName, item);
                        item.LPHID = id;
                        item.Date = DateTime.Now;
                        item.UserID = AccountID;
                        item.ModifiedBy = AccountName;
                        item.ModifiedDate = DateTime.Now;
                        item.LocationID = AccountLocationID;
                        item.SubShift = 1;


                        #region querykhusus
                        string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        using (SqlConnection connection = new SqlConnection(strConString))
                        {
                            connection.Open();
                            SqlTransaction transaction = connection.BeginTransaction();

                            try
                            {
                                SqlCommand command = new SqlCommand("INSERT INTO LPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                                command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                                command.ExecuteNonQuery();
                                transaction.Commit();
                            }
                            catch
                            {
                                transaction.Rollback();
                            }
                        }
                        #endregion querykhusus

                    }

                }
                #endregion

                #region ::Update Values::
                string checkDataComps = _lphComponentsAppService.FindBy("LPHID", id, true);
                List<LPHComponentsModel> dataComps = checkDataComps.DeserializeToLPHComponentList().OrderBy(x => x.ID).ToList();

                if (dataComps.Count() == detailValue.Count())
                {
                    Dictionary<string, long> empIDuserIDMap = new Dictionary<string, long>();

                    string values = _lphValuesAppService.FindBy("SubmissionID", submitModel.ID, true);
                    List<LPHValuesModel> valueModelList = values.DeserializeToLPHValueList();
                    string updateValues = string.Empty;

                    for (int i = 0; i < dataComps.Count(); i++)
                    {
                        var componentID = dataComps[i].ID;
                        var valueModel = valueModelList.Where(x => x.LPHComponentID == componentID).FirstOrDefault();

                        if (dataComps[i].ComponentName.ToLower() == "GeneralInfo-DateInfo".ToLower() || dataComps[i].ComponentName.ToLower() == "Date".ToLower())
                        {
                            string dateX = detailValue[i].Value;
                            string tempDate = dateX.Trim();
                            tempDate = tempDate.Remove(0, 4);
                            tempDate = tempDate.Trim();
                            DateTime newDate = DateTime.ParseExact(tempDate, "dd-MMM-yy", CultureInfo.CurrentCulture);

                            LPHSubmissionsModel submModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
                            submModel.Date = newDate;
                            submModel.ModifiedBy = AccountName;
                            submModel.ModifiedDate = DateTime.Now;

                            string dataSub = JsonHelper<LPHSubmissionsModel>.Serialize(submModel);
                            _lphSubmissionsAppService.Update(dataSub);
                        }

                        if (dataComps[i].ComponentName.ToLower() == "generalinfo-shift" || dataComps[i].ComponentName.ToLower() == "shift")
                        {
                            if (valueModel.Value != detailValue[i].Value)
                            {
                                var shiftX = detailValue[i].Value;

                                string app = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID.ToString(), true);
                                List<LPHApprovalsModel> approveModel_list = app.DeserializeToLPHApprovalList();

                                string updateQuery = string.Empty;
                                foreach (var item in approveModel_list)
                                {
                                    item.Shift = shiftX;
                                    item.ModifiedBy = AccountName;
                                    item.ModifiedDate = DateTime.Now;
                                    updateQuery += string.Format("UPDATE LPHApprovals SET Shift = '{0}', ModifiedBy = '{1}', ModifiedDate = '{2}' WHERE ID = {3};", shiftX, userSupervisorEmpID, AccountName, DateTime.Now, item.ID);
                                }

                                if (!string.IsNullOrEmpty(updateQuery))
                                {
                                    ExecuteQuery(updateQuery);
                                }

                                LPHSubmissionsModel submModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
                                submModel.Shift = shiftX;
                                submModel.ModifiedBy = AccountName;
                                submModel.ModifiedDate = DateTime.Now;

                                string dataSub = JsonHelper<LPHSubmissionsModel>.Serialize(submModel);
                                _lphSubmissionsAppService.Update(dataSub);
                            }
                        }

                        if (dataComps[i].ComponentName.ToLower() == "teamleader" || dataComps[i].ComponentName.ToLower() == "generalinfo-teamleader" || dataComps[i].ComponentName.ToLower() == "supervisor")
                        {
                            if (valueModel.Value != detailValue[i].Value || userSupervisorEmpID != detailValue[i].Value)
                            {
                                var oldUserSupervisorEmpID = userSupervisorEmpID;
                                userSupervisorEmpID = detailValue[i].Value ?? AccountSpvEmployeeID;

                                if (empIDuserIDMap.ContainsKey(userSupervisorEmpID))
                                {
                                    empIDuserIDMap.TryGetValue(userSupervisorEmpID, out userSupervisorUserID);
                                }
                                else
                                {
                                    userSupervisorUserID = GetUserIDByEmployeeId(userSupervisorEmpID, _userAppService);
                                    empIDuserIDMap.Add(userSupervisorEmpID, userSupervisorUserID);
                                }

                                //Do update in approval 'Approver' 
                                string app = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID.ToString(), true);
                                List<LPHApprovalsModel> approveModel_list = app.DeserializeToLPHApprovalList();

                                string updateQuery = string.Empty;

                                foreach (var item in approveModel_list)
                                {
                                    if (item.UserEmployeeID.Equals(valueModel.Value) || item.UserEmployeeID.Equals(oldUserSupervisorEmpID))
                                    {
                                        updateQuery += string.Format("UPDATE LPHApprovals SET UserID = {0}, UserEmployeeID = '{1}', ModifiedBy = '{2}', ModifiedDate = '{3}' WHERE ID = {4};",
                                                                      userSupervisorUserID, userSupervisorEmpID, AccountName, DateTime.Now, item.ID);
                                    }
                                    else
                                    {
                                        updateQuery += string.Format("UPDATE LPHApprovals SET ApproverID = {0}, ApproverEmployeeID = '{1}', ModifiedBy = '{2}', ModifiedDate = '{3}' WHERE ID = {4};",
                                                                      userSupervisorUserID, userSupervisorEmpID, AccountName, DateTime.Now, item.ID);
                                    }
                                }

                                if (!string.IsNullOrEmpty(updateQuery))
                                {
                                    ExecuteQuery(updateQuery);
                                }

                                string dataApproval2 = _lphApprovalAppService.GetBy("LPHSubmissionID", submitModel.ID, true);
                                LPHApprovalsModel modelApp2 = dataApproval2.DeserializeToLPHApproval();

                                if (isSubmit == 1 && modelApp2.Status == "Submitted")
                                {
                                    LPHApprovalsModel mApp = new LPHApprovalsModel
                                    {
                                        LPHSubmissionID = submitModel.ID,
                                        UserID = userSupervisorUserID,
                                        ModifiedBy = userAccountName,
                                        ModifiedDate = DateTime.Now,
                                        Status = "",
                                        Shift = modelApp2.Shift,
                                        Date = DateTime.Now,
                                        LocationID = modelApp2.LocationID,
                                        UserEmployeeID = userSupervisorEmpID,
                                    };

                                    var appID = _lphApprovalAppService.AddModel(mApp);
                                    if (appID == 0)
                                    {
                                        // recreate using ado
                                        //ExecuteQuery("INSERT INTO LPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status) VALUES ('"
                                        //    + submitModel.ID + "','" + DateTime.Now + "','" + modelApp2.Shift + "','" + userSupervisorUserID + "','" + userSupervisorEmpID + "','" + modelApp2.LocationID + "','" + "" + "');");
                                        ExecuteQuery(string.Format("INSERT INTO LPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status,ApproverID,ApproverEmployeeID) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','Draft','{6}','{7}');", submitModel.ID, DateTime.Now, modelApp2.Shift, userAccountID, userEmpID, userLocationID, userSupervisorUserID, userSupervisorEmpID));
                                    }
                                }
                            }
                        }

                        if (valueModel != null ? valueModel.Value != detailValue[i].Value : detailValue[i].Value != null)
                        {
                            if (detailValue[i].ValueType == "ImageURL")
                            {
                                //if (detailValue[i].Value != null ? detailValue[i].Value.Length > 200 : false)
                                if (detailValue[i].Value != null ? detailValue[i].Value.ToString().ToLower().StartsWith("data:image") : false)
                                {
                                    detailValue[i].Value = GetFileName(controllerName, detailValue[i], valueModel.Value);
                                }
                                else
                                {
                                    continue;
                                }
                            }

                            LPHValueHistoriesModel model = new LPHValueHistoriesModel
                            {
                                LPHValuesID = valueModel.ID,
                                OldValue = valueModel.Value,
                                NewValue = detailValue[i].Value,
                                UserID = AccountID,
                                ModifiedBy = AccountName,
                                ModifiedDate = DateTime.Now
                            };

                            _lphValueHistoriesAppService.AddModel(model);

                            valueModel.Value = detailValue[i].Value;

                            updateValues += string.Format("UPDATE LPHValues SET Value = '{0}' WHERE ID = {1};", valueModel.Value, valueModel.ID);
                        }
                    }

                    if (!string.IsNullOrEmpty(updateValues))
                        ExecuteQuery(updateValues);
                }
                #endregion

                #region querySaveLog
                string strConString2 = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(strConString2))
                {
                    connection.Open();
                    SqlTransaction transaction = connection.BeginTransaction();

                    try
                    {
                        SqlCommand command = new SqlCommand("INSERT INTO SaveLogs(Date,LPHHeader,ComputerName,IpAddress,LPHID,UserID,IsSubmit) VALUES (@Date,@LPHHeader,@ComputerName,@IpAddress,@LPHID,@UserID,@IsSubmit);", connection, transaction);
                        command.Parameters.Add("@Date", SqlDbType.DateTime).Value = DateTime.Now;
                        command.Parameters.Add("@LPHHeader", SqlDbType.VarChar).Value = submitModel.LPHHeader;
                        command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = submitModel.LPHID;

                        string comName = System.Net.Dns.GetHostName().ToString();
                        string ipList = Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

                        if (!string.IsNullOrEmpty(ipList))
                            ipList = ipList.Split(',')[0];
                        else
                            ipList = Request.ServerVariables["REMOTE_ADDR"];

                        if (!string.IsNullOrEmpty(comName))
                            command.Parameters.Add("@ComputerName", SqlDbType.VarChar).Value = comName;
                        else
                            command.Parameters.Add("@ComputerName", SqlDbType.VarChar).Value = "";

                        if (!string.IsNullOrEmpty(ipList))
                            command.Parameters.Add("@IpAddress", SqlDbType.VarChar).Value = ipList;
                        else
                            command.Parameters.Add("@IpAddress", SqlDbType.VarChar).Value = "";

                        command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = userAccountID;
                        command.Parameters.Add("@IsSubmit", SqlDbType.Int).Value = isSubmit;

                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " SaveLogs: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                        transaction.Rollback();
                    }
                }
                #endregion querySaveLog


                //add new row approval for row approverSPV
                #region ::Update Approval::
                string dataApproval = _lphApprovalAppService.GetBy("LPHSubmissionID", submitModel.ID, true);
                LPHApprovalsModel modelApp = dataApproval.DeserializeToLPHApproval();

                if (isSubmit == 1 && modelApp.Status == "Submitted")
                {
                    LPHApprovalsModel mApp = new LPHApprovalsModel
                    {
                        LPHSubmissionID = submitModel.ID,
                        UserID = modelApp.ApproverID,
                        ModifiedBy = userAccountName,
                        ModifiedDate = DateTime.Now,
                        Status = "",
                        Shift = modelApp.Shift,
                        Date = DateTime.Now,
                        LocationID = modelApp.LocationID,
                        UserEmployeeID = modelApp.ApproverEmployeeID,
                    };

                    _lphApprovalAppService.AddModel(mApp);
                }

                if (modelApp.Status.Equals("Revised"))
                {
                    LPHApprovalsModel NewModelApp = new LPHApprovalsModel
                    {
                        LPHSubmissionID = submitModel.ID,
                        UserID = userAccountID,
                        ModifiedBy = userAccountName,
                        ModifiedDate = DateTime.Now,
                        Status = "Draft",
                        Date = DateTime.Now,
                        LocationID = userLocationID,
                        ApproverID = userSupervisorUserID,
                        ApproverEmployeeID = userSupervisorEmpID,
                        UserEmployeeID = userEmpID
                    };

                    _lphApprovalAppService.AddModel(NewModelApp);
                }
                #endregion

                #region ::Send Email::
                //Email LPH Ketika Submit
                EmployeeModel employeeSupervisor = GetEmployeeByEmployeeId(userSupervisorEmpID, _employeeAppService);
                EmployeeModel employeeUser = GetEmployeeByEmployeeId(user.EmployeeID, _employeeAppService);

                if (isSubmit == 1 && !string.IsNullOrEmpty(employeeSupervisor.Email))
                {
                    await EmailSender.SendEmailLPHPPSubmit(employeeSupervisor.Email, employeeUser.FullName + " (" + userEmpID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHHeader.Replace("Controller", ""), submitModel.LPHID.ToString());
                }
                #endregion
            }
        }
        #endregion

        #region ::Upload Extras Images::
        protected void UploadExtrasImages(long id, List<LPHExtrasModel> detailExtras, string controllerName)
        {
            if (detailExtras != null)
            {
                detailExtras = detailExtras.Where(x => x.ValueType.Contains("ImageURL")).ToList();

                foreach (var item in detailExtras)
                {
                    item.Value = GetValueOnEdit(controllerName, item);
                    item.LPHID = id;
                    item.Date = DateTime.Now;
                    item.UserID = AccountID;
                    item.ModifiedBy = AccountName;
                    item.ModifiedDate = DateTime.Now;
                    item.LocationID = AccountLocationID;
                    item.SubShift = 1;

                    //ExecuteQuery("INSERT INTO LPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES ('" + item.LPHID + "','" + item.HeaderName + "','" + item.FieldName + "','" + item.Value + "','" + item.ValueType + "','" + item.Date + "','" + item.SubShift + "','" + item.UserID + "','0','" + item.ModifiedBy + "','" + item.ModifiedDate + "','" + item.RowNumber + "','" + item.LocationID + "');");
                    //ExecuteQuery(string.Format("INSERT INTO LPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','0','{8}','{9}','{10}','{11}');", item.LPHID, item.HeaderName, item.FieldName, item.Value, item.ValueType, item.Date, item.SubShift, item.UserID, item.ModifiedBy, item.ModifiedDate, item.RowNumber, item.LocationID));
                    #region querykhusus
                    string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    using (SqlConnection connection = new SqlConnection(strConString))
                    {
                        connection.Open();
                        SqlTransaction transaction = connection.BeginTransaction();

                        try
                        {
                            SqlCommand command = new SqlCommand("INSERT INTO LPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                            command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                            command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                            command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                            command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                            command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                            command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                            command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                            command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                            command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                            command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                            command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                            command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                            command.ExecuteNonQuery();
                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                        }
                    }
                    #endregion querykhusus
                }
            }
        }
        #endregion

        #region ::Approve LPH SP::
        protected async Task ApproveLPH(
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            ILPHSubmissionsAppService _lphSubmissionsAppService,
            ILPHExtrasAppService _lphExtrasAppService,
            ILPHComponentsAppService _lphComponentsAppService,
            ILPHValuesAppService _lphValuesAppService,
            ILPHValueHistoriesAppService _lphValueHistoriesAppService,
            ILPHApprovalsAppService _lphApprovalAppService,
            long id,
            string comment,
            List<LPHExtrasModel> detailExtras,
            List<LPHValuesModel> detailValue,
            string controllerName)
        {
            LPHSubmissionsModel submitModel = _lphSubmissionsAppService.GetBy("LPHID", id, true).DeserializeToLPHSubmissions();

            #region ::Detail Extras::
            //DELETE LPHExtra secepat kilat            			
            ExecuteQuery(string.Format("UPDATE LPHExtras SET IsDeleted = 1 WHERE LPHID = {0};", id));

            if (detailExtras != null)
            {
                foreach (var item in detailExtras)
                {
                    if (item.ValueType.Contains("ImageURL"))
                        item.Value = GetValueOnApproval(controllerName, item);
                    item.LPHID = id;
                    item.Date = DateTime.Now;
                    item.UserID = AccountID;
                    item.ModifiedBy = AccountName;
                    item.ModifiedDate = DateTime.Now;
                    item.LocationID = AccountLocationID;
                    item.SubShift = 1;

                    #region querykhusus
                    string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    using (SqlConnection connection = new SqlConnection(strConString))
                    {
                        connection.Open();
                        SqlTransaction transaction = connection.BeginTransaction();

                        try
                        {
                            SqlCommand command = new SqlCommand("INSERT INTO LPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                            command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                            command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                            command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                            command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                            command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                            command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                            command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                            command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                            command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                            command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                            command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                            command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                            command.ExecuteNonQuery();
                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                        }
                    }
                    #endregion querykhusus
                }

            }
            #endregion

            #region ::Update Values::
            string checkDataComps = _lphComponentsAppService.FindBy("LPHID", id, true);
            List<LPHComponentsModel> dataComps = checkDataComps.DeserializeToLPHComponentList().OrderBy(x => x.ID).ToList();

            if (dataComps.Count() == detailValue.Count())
            {
                Dictionary<string, long> empIDuserIDMap = new Dictionary<string, long>();

                string values = _lphValuesAppService.FindBy("SubmissionID", submitModel.ID, true);
                List<LPHValuesModel> valueModelList = values.DeserializeToLPHValueList();
                string updateValues = string.Empty;

                for (int i = 0; i < dataComps.Count(); i++)
                {
                    var componentID = dataComps[i].ID;
                    var valueModel = valueModelList.Where(x => x.LPHComponentID == componentID).FirstOrDefault();

                    #region ::Shift::
                    if (dataComps[i].ComponentName.ToLower() == "generalinfo-shift" || dataComps[i].ComponentName.ToLower() == "shift")
                    {
                        if (valueModel.Value != detailValue[i].Value)
                        {
                            var shiftX = detailValue[i].Value;

                            string app = _lphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<LPHApprovalsModel> approve2Model_list = app.DeserializeToLPHApprovalList();

                            foreach (var item in approve2Model_list)
                            {
                                item.Shift = shiftX;
                                item.ModifiedBy = AccountName;
                                item.ModifiedDate = DateTime.Now;

                                string dataAppSub = JsonHelper<LPHApprovalsModel>.Serialize(item);
                                _lphApprovalAppService.Update(dataAppSub);
                            }

                            LPHSubmissionsModel submModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
                            submModel.Shift = shiftX;
                            submModel.ModifiedBy = AccountName;
                            submModel.ModifiedDate = DateTime.Now;

                            string dataSub = JsonHelper<LPHSubmissionsModel>.Serialize(submModel);
                            _lphSubmissionsAppService.Update(dataSub);
                        }
                    }
                    #endregion

                    #region ::Team Leader::
                    if (dataComps[i].ComponentName.ToLower() == "teamleader" || dataComps[i].ComponentName.ToLower() == "generalinfo-teamleader")
                    {
                        if (valueModel.Value != detailValue[i].Value)
                        {
                            var approverEmployeeID = detailValue[i].Value ?? AccountSpvEmployeeID;

                            //Do update in approval 'Approver' 
                            string app = _lphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<LPHApprovalsModel> approve2Model_list = app.DeserializeToLPHApprovalList();

                            long userID;
                            foreach (var item in approve2Model_list)
                            {
                                if (valueModel.Value != null)
                                {
                                    if (empIDuserIDMap.ContainsKey(approverEmployeeID))
                                    {
                                        empIDuserIDMap.TryGetValue(approverEmployeeID, out userID);
                                    }
                                    else
                                    {
                                        userID = GetUserIDByEmployeeId(approverEmployeeID, _userAppService);
                                        empIDuserIDMap.Add(approverEmployeeID, userID);
                                    }

                                    if (item.ApproverEmployeeID.Equals(valueModel.Value))
                                    {
                                        item.ApproverID = userID;
                                        item.ApproverEmployeeID = approverEmployeeID;
                                    }
                                    else if (item.ApproverEmployeeID != null && !item.UserEmployeeID.Equals(valueModel.Value))
                                    {
                                        item.UserID = userID;
                                        item.UserEmployeeID = approverEmployeeID;
                                    }

                                    item.ModifiedBy = AccountName;
                                    item.ModifiedDate = DateTime.Now;

                                    string dataNewApprover = JsonHelper<LPHApprovalsModel>.Serialize(item);
                                    _lphApprovalAppService.Update(dataNewApprover);
                                    //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 1432 " + dataNewApprover.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                                }
                            }
                        }
                    }
                    #endregion

                    #region ::Image URL::
                    if (valueModel != null ? valueModel.Value != detailValue[i].Value : detailValue[i].Value != null)
                    {
                        if (detailValue[i].ValueType == "ImageURL")
                        {
                            //if (detailValue[i].Value != null ? detailValue[i].Value.Length > 200 : false)
                            if (detailValue[i].Value != null ? detailValue[i].Value.ToString().ToLower().StartsWith("data:image") : false)
                            {
                                detailValue[i].Value = GetFileName(controllerName, detailValue[i], valueModel.Value);
                            }
                            else
                            {
                                continue;
                            }
                        }

                        valueModel.Value = detailValue[i].Value;

                        updateValues += string.Format("UPDATE LPHValues SET Value = '{0}' WHERE ID = {1};", valueModel.Value, valueModel.ID);
                    }
                    #endregion
                }

                if (!string.IsNullOrEmpty(updateValues))
                    ExecuteQuery(updateValues);
            }
            #endregion

            #region ::Update Approval::
            LPHApprovalsModel approveModel = _lphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID).DeserializeToLPHApproval();
            approveModel.Status = "Approved";
            approveModel.ModifiedBy = AccountName;
            approveModel.ModifiedDate = DateTime.Now;
            approveModel.Notes = comment;
            approveModel.ApproverID = AccountID;

            string dataApproval = JsonHelper<LPHApprovalsModel>.Serialize(approveModel);
            _lphApprovalAppService.Update(dataApproval);
            #endregion

            #region ::Send Email::
            // Email Ketika Approve :::::::::::::
            string employeeIDRequestor = GetEmployeeIDByUserID(submitModel.UserID, _userAppService);
            EmployeeModel employeeRequestor = GetEmployeeByEmployeeId(employeeIDRequestor, _employeeAppService);

            if (!string.IsNullOrEmpty(employeeRequestor.Email))
            {
                await EmailSender.SendEmailLPHPPApprove(employeeRequestor.Email, employeeRequestor.FullName + " (" + employeeRequestor.EmployeeID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHID.ToString(), submitModel.LPHHeader.Replace("Controller", ""));
            }
            #endregion
        }
        #endregion

        #region ::Revise LPH SP::
        // Revised use this method (no reject term anymore)
        protected async Task RejectLPH(
            IEmployeeAppService _employeeAppService,
            IUserAppService _userAppService,
            ILPHSubmissionsAppService _lphSubmissionsAppService,
            ILPHApprovalsAppService _lphApprovalAppService,
            long id,
            string notes)
        {
            #region Session
            UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
            long userAccountID = user.ID;
            string userEmpID = user.EmployeeID;
            long userSupervisorUserID = user.SupervisorUserID;
            string userSupervisorEmpID = user.SupervisorID;
            long userLocationID = user.LocationID.HasValue ? user.LocationID.Value : 0;
            string userAccountName = user.UserName;
            #endregion

            #region Submission
            LPHSubmissionsModel submitModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
            submitModel.IsDeleted = true;
            _lphSubmissionsAppService.UpdateModel(submitModel);
            #endregion

            #region Approval
            //fery optimizing			
            string approvalC = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID);
            List<LPHApprovalsModel> approvalCheck = approvalC.DeserializeToLPHApprovalList().OrderBy(x => x.Date).ToList();

            // Update status to 'revised'
            LPHApprovalsModel approveModel = _lphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID).DeserializeToLPHApproval();
            approveModel.Status = "Revised";
            approveModel.ModifiedBy = AccountName;
            approveModel.ModifiedDate = DateTime.Now;
            approveModel.Notes = notes;
            approveModel.ApproverID = AccountID;

            string dataApproval = JsonHelper<LPHApprovalsModel>.Serialize(approveModel);
            _lphApprovalAppService.Update(dataApproval);

            //Create new record with status draft (like first step)
            LPHApprovalsModel modelApp = new LPHApprovalsModel
            {
                LPHSubmissionID = submitModel.ID,
                UserID = approvalCheck.ElementAt(0).UserID,
                ModifiedBy = AccountName,
                ModifiedDate = DateTime.Now,
                Status = "Draft",
                Shift = approvalCheck.ElementAt(0).Shift,
                Date = approvalCheck.ElementAt(0).Date,
                LocationID = approvalCheck.ElementAt(0).LocationID,
                ApproverID = approvalCheck.ElementAt(0).ApproverID,
                ApproverEmployeeID = approvalCheck.ElementAt(0).ApproverEmployeeID,
                UserEmployeeID = approvalCheck.ElementAt(0).UserEmployeeID
            };

            _lphApprovalAppService.AddModel(modelApp);
            #endregion

            #region Send Email
            //Email Revise LPH
            string employeeIDRequestor = GetEmployeeIDByUserID(submitModel.UserID, _userAppService);
            EmployeeModel employeeRequestor = GetEmployeeByEmployeeId(employeeIDRequestor, _employeeAppService);

            EmployeeModel employeeUser = GetEmployeeByEmployeeId(user.EmployeeID, _employeeAppService);
            if (!string.IsNullOrEmpty(employeeRequestor.Email))
            {
                await EmailSender.SendEmailLPHPPRevise(employeeRequestor.Email, employeeUser.FullName + " (" + AccountEmployeeID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHID.ToString(), submitModel.LPHHeader.Replace("Controller", ""), submitModel.LPHID.ToString());
            }
            #endregion
        }
        #endregion

        #region ::Create LPH PP::
        protected long CreateLPHPrimary(
            IPPLPHAppService _ppLphAppService,
            IPPLPHLocationsAppService _ppLphLocationsAppService,
            IPPLPHExtrasAppService _ppLphExtrasAppService,
            IPPLPHSubmissionsAppService _ppLphSubmissionsAppService,
            IPPLPHApprovalsAppService _ppLphApprovalAppService,
            IPPLPHComponentsAppService _ppLphComponentsAppService,
            IPPLPHValuesAppService _ppLphValuesAppService,
            List<PPLPHExtrasModel> detailsExtras,
            List<PPLPHComponentsModel> detailComponent,
            List<PPLPHValuesModel> detailValue,
            string controllerName)
        {
            #region Session
            long userAccountID = AccountID;
            long userLocationID = AccountLocationID;
            string userAccountName = AccountName;
            string currentShift = GetShift();
            #endregion

            #region LPH
            PPLPHModel model = new PPLPHModel
            {
                MenuTitle = controllerName,
                Header = controllerName,
                Type = "PP",
                LocationID = userLocationID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                IsDeleted = false // chanif: karena create awal sudah pasti jadi draft; edit: jadi acuan untuk lihat lph dihapus/atau tidak
            };

            long lphID = _ppLphAppService.AddModel(model);
            #endregion

            #region LPH Locations			
            PPLPHLocationsModel lphLocModel = new PPLPHLocationsModel
            {
                LPHID = lphID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                LocationID = userLocationID
            };

            _ppLphLocationsAppService.AddModel(lphLocModel);
            #endregion

            #region LPH Submissions
            // chanif: define subshift
            DateTime tanggal = DateTime.Now;
            DateTime yesterday = tanggal.AddDays(-1);

            string submit = _ppLphSubmissionsAppService.FindBy("LocationID", AccountLocationID, true);
            List<PPLPHSubmissionsModel> submitModel = string.IsNullOrEmpty(submit) ? new List<PPLPHSubmissionsModel>() : JsonConvert.DeserializeObject<List<PPLPHSubmissionsModel>>(submit);
            submitModel = submitModel.Where(x => x.LPHHeader == model.Header && x.Shift.Trim() == currentShift && x.ModifiedDate > tanggal.AddHours(-24) && x.Date > yesterday).OrderByDescending(x => x.ID).ToList();

            int subshift = 1;
            if (submitModel != null && submitModel.Count() > 0)
            {
                subshift = Int32.Parse(submitModel[0].SubShift.ToString()) + 1;
            }

            PPLPHSubmissionsModel lphSubmissionsModel = new PPLPHSubmissionsModel
            {
                LPHID = lphID,
                LPHHeader = model.Header,
                Date = DateTime.Now,
                Shift = currentShift,
                SubShift = subshift,
                UserID = userAccountID,
                LocationID = userLocationID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                IsDeleted = true // chanif: karena create awal sudah pasti jadi draft. supaya tidak muncul di history dan approval, isdeleted harus true
            };
            var submissionID = _ppLphSubmissionsAppService.AddModel(lphSubmissionsModel);
            #endregion

            #region LPH Approvals
            PPLPHApprovalsModel modelApp = new PPLPHApprovalsModel
            {
                LPHSubmissionID = submissionID,
                UserID = userAccountID,
                ModifiedBy = userAccountName,
                ModifiedDate = DateTime.Now,
                Status = "Draft",
                Date = DateTime.Now,
                LocationID = userLocationID,
                ApproverID = AccountSpvUserID,
                ApproverEmployeeID = AccountSpvEmployeeID,
                UserEmployeeID = AccountEmployeeID,
                Shift = currentShift
            };
            var appID = _ppLphApprovalAppService.AddModel(modelApp);
            if (appID == 0)
            {
                // recreate using ado
                //ExecuteQuery("INSERT INTO PPLPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status,ApproverID,ApproverEmployeeID) VALUES ('"
                //    + submissionID + "','" + DateTime.Now + "','" + currentShift + "','" + userAccountID + "','" + AccountEmployeeID + "','" + userLocationID + "','" + "Draft" + "','" + AccountSpvEmployeeID + "','" + AccountSpvEmployeeID + "');");
                ExecuteQuery(string.Format("INSERT INTO PPLPHApprovals (LPHSubmissionID,Date,Shift,UserID,UserEmployeeID,LocationID,Status,ApproverID,ApproverEmployeeID) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','Draft','{6}','{7}');", submissionID, DateTime.Now, currentShift, userAccountID, AccountEmployeeID, userLocationID, AccountSpvEmployeeID, AccountSpvEmployeeID));
            }

            #endregion

            #region LPH Components			
            if (detailComponent != null)
            {
                detailComponent.ForEach(x => x.LPHID = lphID);
                detailComponent.ForEach(x => x.ModifiedBy = userAccountName);
                detailComponent.ForEach(x => x.ModifiedDate = DateTime.Now);

                _ppLphComponentsAppService.AddRangeModel(detailComponent);

                detailComponent = _ppLphComponentsAppService.FindBy("LPHID", lphID, true).DeserializeToPPLPHComponentList();
            }
            #endregion

            #region LPH Values
            if (detailComponent.Count() == detailValue.Count())
            {
                List<PPLPHValuesModel> valueList = new List<PPLPHValuesModel>();

                for (int i = 0; i < detailComponent.Count(); i++)
                {
                    PPLPHValuesModel modelValue = new PPLPHValuesModel();
                    modelValue.Value = GetValuePP(detailValue[i], modelValue);
                    modelValue.ValueType = detailValue[i].ValueType;
                    modelValue.LPHComponentID = detailComponent.ElementAt(i).ID;
                    modelValue.ValueDate = DateTime.Now;
                    modelValue.SubmissionID = submissionID;
                    modelValue.ModifiedBy = userAccountName;
                    modelValue.ModifiedDate = DateTime.Now;

                    valueList.Add(modelValue);
                }

                _ppLphValuesAppService.AddRangeModel(valueList);
            }
            #endregion

            #region LPH Extras			
            if (detailsExtras != null)
            {
                foreach (var item in detailsExtras.ToList())
                {
                    item.LPHID = lphID;
                    item.Date = DateTime.Now;
                    item.UserID = userAccountID;
                    item.ModifiedBy = userAccountName;
                    item.ModifiedDate = DateTime.Now;
                    item.LocationID = userLocationID;
                    item.SubShift = subshift;
                }

                if (detailsExtras.Count > 0)
                {
                    string newExtras = JsonHelper<PPLPHExtrasModel>.Serialize(detailsExtras);
                    _ppLphExtrasAppService.AddRange(newExtras);
                }
            }
            #endregion

            return lphID;
        }
        #endregion

        #region ::Edit LPH PP::
        protected async Task EditLPHPrimary(
            IUserAppService _userAppService,
            IEmployeeAppService _employeeAppService,
            IPPLPHSubmissionsAppService _ppLphSubmissionsAppService,
            IPPLPHExtrasAppService _ppLphExtrasAppService,
            IPPLPHComponentsAppService _ppLphComponentsAppService,
            IPPLPHValuesAppService _ppLphValuesAppService,
            IPPLPHValueHistoriesAppService _ppLphValueHistoriesAppService,
            IPPLPHApprovalsAppService _ppLphApprovalAppService,
            long id, List<PPLPHExtrasModel> detailsExtras, List<PPLPHValuesModel> detailValue, int isSubmit = 0, int Special = 0)
        {
            PPLPHSubmissionsModel submitModel = _ppLphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToPPLPHSubmissions();
            if (submitModel != null && submitModel.IsDeleted)
            {
                #region Session
                UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
                long userAccountID = user.ID;
                string userEmpID = user.EmployeeID;
                long userSupervisorUserID = user.SupervisorUserID;
                string userSupervisorEmpID = user.SupervisorID;
                long userLocationID = user.LocationID.HasValue ? user.LocationID.Value : 0;
                string userAccountName = user.UserName;
                #endregion

                #region Approval
                string approval = _ppLphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID.ToString());
                List<PPLPHApprovalsModel> approvalList = approval.DeserializeToPPLPHApprovalList().ToList();
                approvalList = approvalList.Where(x => x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
                PPLPHApprovalsModel apprModel = approvalList.FirstOrDefault();

                // chanif: rubah status isdeleted jika draft
                if (isSubmit == 1 && submitModel != null && apprModel.Status.Trim().ToLower() == "draft")
                {
                    submitModel.IsDeleted = false;
                    string submit = JsonHelper<PPLPHSubmissionsModel>.Serialize(submitModel);
                    _ppLphSubmissionsAppService.Update(submit);

                    //Do update in approval with status 'submitted' 				
                    PPLPHApprovalsModel approveModel = _ppLphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID, true).DeserializeToPPLPHApproval();
                    approveModel.Status = "Submitted";
                    approveModel.ModifiedBy = AccountName;
                    approveModel.ModifiedDate = DateTime.Now;
                    approveModel.Notes = "";

                    string dataAppSub = JsonHelper<PPLPHApprovalsModel>.Serialize(approveModel);
                    _ppLphApprovalAppService.Update(dataAppSub);
                }
                #endregion

                #region DetailExtras
                //DELETE LPHExtra secepat kilat            			
                ExecuteQuery(string.Format("UPDATE PPLPHExtras SET IsDeleted=1 WHERE ValueType!='JSON' AND LPHID = {0};", id));
                ExecuteQuery(string.Format("UPDATE PPLPHExtras SET Shift='DELJSON' WHERE ValueType='JSON' AND LPHID = {0};", id));
                ////ExecuteQuery(string.Format("DELETE FROM PPLPHExtras WHERE LPHID = {0};", id));
                bool flagrecover = false;
                if (detailsExtras != null)
                {
                    //checking anti JSON error
                    foreach (var item in detailsExtras)
                    {
                        if (item.ValueType.ToLower() == "json")
                        {
                            if (item.Value.Length > 5 && item.Value.Length <= 30)
                            {
                                flagrecover = true;
                                break;
                            }
                        }
                    }

                    foreach (var item in detailsExtras)
                    {
                        item.LPHID = id;
                        item.Date = DateTime.Now;
                        item.UserID = AccountID;
                        item.ModifiedBy = AccountName;
                        item.ModifiedDate = DateTime.Now;
                        item.LocationID = AccountLocationID;
                        item.SubShift = submitModel.SubShift;
                        if (item.ValueType.ToLower() == "json")
                        {
                            //Helper.LogErrorMessageAddback(DateTime.Now.ToString() + " " + AccountEmployeeID + " EDIT JSON LPHID=" + item.LPHID + " HEADER,FIELD=" + item.HeaderName + "," + item.FieldName + " ROWNUMBER=" + item.RowNumber + " VALUE=" + item.Value, Server.MapPath("~/Uploads/"), AccountEmployeeID);

                            #region querykhusus2
                            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            using (SqlConnection connection = new SqlConnection(strConString))
                            {
                                connection.Open();
                                SqlTransaction transaction = connection.BeginTransaction();

                                try
                                {
                                    SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtrasTemp(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID,Value2) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID,@Value2);", connection, transaction);
                                    command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                    command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                    command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                    command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                    command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                    command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                    command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                    command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                    command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                    command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                    command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                    command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;
                                    command.Parameters.Add("@Value2", SqlDbType.VarChar).Value = item.Value;

                                    command.ExecuteNonQuery();
                                    transaction.Commit();
                                }
                                catch
                                {
                                    transaction.Rollback();
                                }
                            }
                            #endregion querykhusus

                            if (!flagrecover)
                            {
                                #region querykhusus
                                using (SqlConnection connection = new SqlConnection(strConString))
                                {
                                    connection.Open();
                                    SqlTransaction transaction = connection.BeginTransaction();

                                    try
                                    {
                                        SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                                        command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                        command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                        command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                        command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                        command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                        command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                        command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                        command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                        command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                        command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                        command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                        command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                                        command.ExecuteNonQuery();
                                        transaction.Commit();
                                    }
                                    catch
                                    {
                                        transaction.Rollback();
                                    }
                                }
                                #endregion querykhusus
                            }
                        }
                        else
                        {

                            #region querykhusus
                            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            using (SqlConnection connection = new SqlConnection(strConString))
                            {
                                connection.Open();
                                SqlTransaction transaction = connection.BeginTransaction();

                                try
                                {
                                    SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                                    command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                    command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                    command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                    command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                    command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                    command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                    command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                    command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                    command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                    command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                    command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                    command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                                    command.ExecuteNonQuery();
                                    transaction.Commit();
                                }
                                catch
                                {
                                    transaction.Rollback();
                                }
                            }
                            #endregion querykhusus
                        }
                    }

                }
                else
                {
                    Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " detailsExtras == null , LPHID = " + id, Server.MapPath("~/Uploads/"), AccountEmployeeID);
                }

                if (!flagrecover)
                    ExecuteQuery(string.Format("UPDATE PPLPHExtras SET IsDeleted = 1 WHERE Shift='DELJSON' AND LPHID = {0};", id));
                else
                    ExecuteQuery(string.Format("UPDATE PPLPHExtras SET Shift='' WHERE Shift='DELJSON' AND ValueType='JSON' AND LPHID = {0};", id));
                #endregion

                #region Update Values
                string checkDataComps = _ppLphComponentsAppService.FindBy("LPHID", id, true);
                List<PPLPHComponentsModel> dataComps = checkDataComps.DeserializeToPPLPHComponentList();
                dataComps = dataComps.OrderBy(x => x.ID).ToList();

                string shiftX = "";
                string dateX = "";

                if (dataComps.Count() == detailValue.Count())
                {
                    Dictionary<string, long> empIDuserIDMap = new Dictionary<string, long>();

                    string values = _ppLphValuesAppService.FindBy("SubmissionID", submitModel.ID, true);
                    List<PPLPHValuesModel> valueModelList = values.DeserializeToPPLPHValueList();
                    string updateValues = string.Empty;

                    for (int i = 0; i < dataComps.Count(); i++)
                    {
                        var componentID = dataComps[i].ID;
                        var valueModel = valueModelList.Where(x => x.LPHComponentID == componentID).FirstOrDefault();

                        string tempDate = "";
                        if (dataComps[i].ComponentName == "Information-Date")
                        {
                            dateX = detailValue[i].Value;
                            tempDate = dateX.Trim();
                            tempDate = tempDate.Remove(0, 4);
                            tempDate = tempDate.Trim();
                            DateTime newDate = DateTime.ParseExact(tempDate, "dd-MMM-yy", CultureInfo.CurrentCulture);

                            string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<PPLPHApprovalsModel> approveModel_list = app.DeserializeToPPLPHApprovalList();

                            foreach (var item in approveModel_list)
                            {
                                item.Date = newDate;
                                item.ModifiedBy = AccountName;
                                item.ModifiedDate = DateTime.Now;

                                string dataAppSub = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                _ppLphApprovalAppService.Update(dataAppSub);
                            }

                            PPLPHSubmissionsModel submModel = _ppLphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToPPLPHSubmissions();
                            submModel.Date = newDate;
                            submModel.ModifiedBy = AccountName;
                            submModel.ModifiedDate = DateTime.Now;

                            string dataSub = JsonHelper<PPLPHSubmissionsModel>.Serialize(submModel);
                            _ppLphSubmissionsAppService.Update(dataSub);
                        }

                        if (dataComps[i].ComponentName == "Information-Shift")
                        {
                            if (valueModel.Value != detailValue[i].Value)
                            {
                                shiftX = detailValue[i].Value;

                                string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                                List<PPLPHApprovalsModel> approveModel_list = app.DeserializeToPPLPHApprovalList();

                                foreach (var item in approveModel_list)
                                {
                                    item.Shift = shiftX;
                                    item.ModifiedBy = AccountName;
                                    item.ModifiedDate = DateTime.Now;

                                    string dataAppSub = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                    _ppLphApprovalAppService.Update(dataAppSub);
                                }

                                PPLPHSubmissionsModel submModel = _ppLphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToPPLPHSubmissions();
                                submModel.Shift = shiftX;
                                submModel.ModifiedBy = AccountName;
                                submModel.ModifiedDate = DateTime.Now;

                                string dataSub = JsonHelper<PPLPHSubmissionsModel>.Serialize(submModel);
                                _ppLphSubmissionsAppService.Update(dataSub);
                            }
                        }

                        if (dataComps[i].ComponentName == "Information-TeamLeader")
                        {
                            if (valueModel.Value != detailValue[i].Value || userSupervisorEmpID != detailValue[i].Value)
                            {
                                var oldUserSupervisorEmpID = userSupervisorEmpID;
                                userSupervisorEmpID = detailValue[i].Value;
                                userSupervisorUserID = GetUserIDByEmployeeId(userSupervisorEmpID, _userAppService);

                                //Do update in approval 'Approver' 
                                string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                                List<PPLPHApprovalsModel> approveModel_list = app.DeserializeToPPLPHApprovalList();

                                foreach (var item in approveModel_list)
                                {
                                    if (item.UserEmployeeID.Equals(valueModel.Value) || item.UserEmployeeID.Equals(oldUserSupervisorEmpID))
                                    {
                                        item.UserID = userSupervisorUserID;
                                        item.UserEmployeeID = userSupervisorEmpID;
                                    }
                                    else
                                    {
                                        item.ApproverID = userSupervisorUserID;
                                        item.ApproverEmployeeID = userSupervisorEmpID;
                                    }

                                    item.ModifiedBy = AccountName;
                                    item.ModifiedDate = DateTime.Now;

                                    string dataNewApprover = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                    _ppLphApprovalAppService.Update(dataNewApprover);
                                }
                            }
                        }

                        if (valueModel != null ? valueModel.Value != detailValue[i].Value : detailValue[i].Value != null)
                        {
                            if (detailValue[i].ValueType == "ImageURL")
                            {
                                //if (detailValue[i].Value != null ? detailValue[i].Value.Length > 200 : false)
                                if (detailValue[i].Value != null ? detailValue[i].Value.ToString().ToLower().StartsWith("data:image") : false)
                                {
                                    detailValue[i].Value = GetFileNamePP(detailValue[i]);
                                }
                                else
                                {
                                    continue;
                                }
                            }

                            valueModel.Value = detailValue[i].Value;

                            updateValues += string.Format("UPDATE PPLPHValues SET Value = '{0}' WHERE ID = {1};", valueModel.Value, valueModel.ID);
                        }
                    }

                    if (!string.IsNullOrEmpty(updateValues))
                        ExecuteQuery(updateValues);
                }
                #endregion

                #region querySaveLog
                string strConString2 = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(strConString2))
                {
                    connection.Open();
                    SqlTransaction transaction = connection.BeginTransaction();

                    try
                    {
                        SqlCommand command = new SqlCommand("INSERT INTO SaveLogs(Date,LPHHeader,ComputerName,IpAddress,LPHID,UserID,IsSubmit) VALUES (@Date,@LPHHeader,@ComputerName,@IpAddress,@LPHID,@UserID,@IsSubmit);", connection, transaction);
                        command.Parameters.Add("@Date", SqlDbType.DateTime).Value = DateTime.Now;
                        command.Parameters.Add("@LPHHeader", SqlDbType.VarChar).Value = submitModel.LPHHeader;
                        command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = submitModel.LPHID;

                        string comName = System.Net.Dns.GetHostName().ToString();
                        string ipList = Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

                        if (!string.IsNullOrEmpty(ipList))
                            ipList = ipList.Split(',')[0];
                        else
                            ipList = Request.ServerVariables["REMOTE_ADDR"];

                        if (!string.IsNullOrEmpty(comName))
                            command.Parameters.Add("@ComputerName", SqlDbType.VarChar).Value = comName;
                        else
                            command.Parameters.Add("@ComputerName", SqlDbType.VarChar).Value = "";

                        if (!string.IsNullOrEmpty(ipList))
                            command.Parameters.Add("@IpAddress", SqlDbType.VarChar).Value = ipList;
                        else
                            command.Parameters.Add("@IpAddress", SqlDbType.VarChar).Value = "";

                        command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = userAccountID;
                        command.Parameters.Add("@IsSubmit", SqlDbType.Int).Value = isSubmit;

                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " SaveLogs: " + ex.GetAllMessages(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                        transaction.Rollback();
                    }
                }
                #endregion querySaveLog

                #region Approval
                string dataApproval = _ppLphApprovalAppService.GetBy("LPHSubmissionID", submitModel.ID, true);
                PPLPHApprovalsModel modelApp = dataApproval.DeserializeToPPLPHApproval();

                if (isSubmit == 1 && modelApp.Status == "Submitted")
                {
                    //Create new row with blank info
                    PPLPHApprovalsModel mApp = new PPLPHApprovalsModel
                    {
                        LPHSubmissionID = submitModel.ID,
                        UserID = modelApp.ApproverID,
                        ModifiedBy = userAccountName,
                        ModifiedDate = DateTime.Now,
                        Status = "",
                        Shift = modelApp.Shift,
                        Date = DateTime.Now,
                        LocationID = modelApp.LocationID,
                        UserEmployeeID = userSupervisorEmpID,
                    };

                    _ppLphApprovalAppService.AddModel(mApp);
                }

                if (modelApp.Status.Equals("Revised"))
                {
                    PPLPHApprovalsModel NewModelApp = new PPLPHApprovalsModel
                    {
                        LPHSubmissionID = submitModel.ID,
                        UserID = userAccountID,
                        ModifiedBy = userAccountName,
                        ModifiedDate = DateTime.Now,
                        Status = "Draft",
                        Date = DateTime.Now,
                        LocationID = userLocationID,
                        ApproverID = userSupervisorUserID,
                        ApproverEmployeeID = userSupervisorEmpID,
                        UserEmployeeID = userEmpID
                    };

                    _ppLphApprovalAppService.AddModel(NewModelApp);
                }
                #endregion

                #region Send Email
                EmployeeModel employeeSupervisor = GetEmployeeByEmployeeId(userSupervisorEmpID, _employeeAppService);
                EmployeeModel employeeUser = GetEmployeeByEmployeeId(user.EmployeeID, _employeeAppService);

                if (isSubmit == 1 && !string.IsNullOrEmpty(employeeSupervisor.Email))
                {
                    await EmailSender.SendEmailLPHPPSubmit(employeeSupervisor.Email, employeeUser.FullName + " (" + userEmpID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHHeader.Replace("Controller", ""), submitModel.LPHID.ToString());
                }

                #endregion
            }
        }
        #endregion

        #region ::Approve LPH PP::
        protected async Task ApproveLPHPrimary(
            IUserAppService _userAppService,
            IEmployeeAppService _employeeAppService,
            IPPLPHSubmissionsAppService _ppLphSubmissionsAppService,
            IPPLPHExtrasAppService _ppLphExtrasAppService,
            IPPLPHComponentsAppService _ppLphComponentsAppService,
            IPPLPHValuesAppService _ppLphValuesAppService,
            IPPLPHValueHistoriesAppService _ppLphValueHistoriesAppService,
            IPPLPHApprovalsAppService _ppLphApprovalAppService,
            long id, List<PPLPHExtrasModel> detailsExtras, List<PPLPHValuesModel> detailValue, string comment)
        {
            PPLPHSubmissionsModel submitModel = _ppLphSubmissionsAppService.GetBy("LPHID", id.ToString(), true).DeserializeToPPLPHSubmissions();


            #region DetailExtras
            //DELETE LPHExtra secepat kilat            			
            ExecuteQuery(string.Format("UPDATE PPLPHExtras SET IsDeleted = 1 WHERE ValueType!='JSON' AND LPHID = {0};", id));
            ExecuteQuery(string.Format("UPDATE PPLPHExtras SET Shift='DELJSON2' WHERE ValueType='JSON' AND LPHID = {0};", id));

            bool flagrecover = false;
            if (detailsExtras != null)
            {
                //checking anti JSON error
                foreach (var item in detailsExtras)
                {
                    if (item.ValueType.ToLower() == "json")
                    {
                        if (!IsValidJson(item.Value))
                        {
                            flagrecover = true;
                            break;
                        }
                    }
                }

                foreach (var item in detailsExtras)
                {
                    item.LPHID = id;
                    item.Date = DateTime.Now;
                    item.UserID = AccountID;
                    item.ModifiedBy = AccountName;
                    item.ModifiedDate = DateTime.Now;
                    item.LocationID = AccountLocationID;
                    item.SubShift = submitModel.SubShift;
                    if (item.ValueType.ToLower() == "json")
                    {
                        Helper.LogErrorMessageAddback(DateTime.Now.ToString() + " " + AccountEmployeeID + " APPROVE JSON LPHID=" + item.LPHID + " HEADER,FIELD=" + item.HeaderName + "," + item.FieldName + " ROWNUMBER=" + item.RowNumber + " VALUE=" + item.Value, Server.MapPath("~/Uploads/"), AccountEmployeeID);

                        #region querykhusus2
                        string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        using (SqlConnection connection = new SqlConnection(strConString))
                        {
                            connection.Open();
                            SqlTransaction transaction = connection.BeginTransaction();

                            try
                            {
                                SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtrasTemp(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID,Value2) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID,@Value2);", connection, transaction);
                                command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;
                                command.Parameters.Add("@Value2", SqlDbType.VarChar).Value = item.Value;

                                command.ExecuteNonQuery();
                                transaction.Commit();
                            }
                            catch
                            {
                                transaction.Rollback();
                            }
                        }
                        #endregion querykhusus

                        if (!flagrecover)
                        {
                            #region querykhusus
                            using (SqlConnection connection = new SqlConnection(strConString))
                            {
                                connection.Open();
                                SqlTransaction transaction = connection.BeginTransaction();

                                try
                                {
                                    SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                                    command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                    command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                    command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                    command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                    command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                    command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                    command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                    command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                    command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                    command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                    command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                    command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                                    command.ExecuteNonQuery();
                                    transaction.Commit();
                                }
                                catch
                                {
                                    transaction.Rollback();
                                }
                            }
                            #endregion querykhusus
                        }
                    }
                    else
                    {
                        #region querykhusus
                        string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        using (SqlConnection connection = new SqlConnection(strConString))
                        {
                            connection.Open();
                            SqlTransaction transaction = connection.BeginTransaction();

                            try
                            {
                                SqlCommand command = new SqlCommand("INSERT INTO PPLPHExtras(LPHID,HeaderName,FieldName,Value,ValueType,Date,SubShift,UserID,IsDeleted,ModifiedBy,ModifiedDate,RowNumber,LocationID) VALUES (@LPHID,@HeaderName,@FieldName,@Value,@ValueType,@Date,@SubShift,@UserID,'0',@ModifiedBy,@ModifiedDate,@RowNumber,@LocationID);", connection, transaction);
                                command.Parameters.Add("@LPHID", SqlDbType.BigInt).Value = item.LPHID;
                                command.Parameters.Add("@HeaderName", SqlDbType.VarChar).Value = item.HeaderName;
                                command.Parameters.Add("@FieldName", SqlDbType.VarChar).Value = item.FieldName;
                                command.Parameters.Add("@Value", SqlDbType.VarChar).Value = item.Value;
                                command.Parameters.Add("@ValueType", SqlDbType.VarChar).Value = item.ValueType;
                                command.Parameters.Add("@Date", SqlDbType.DateTime).Value = item.Date;
                                command.Parameters.Add("@SubShift", SqlDbType.Int).Value = item.SubShift;
                                command.Parameters.Add("@UserID", SqlDbType.BigInt).Value = item.UserID;
                                command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = item.ModifiedBy;
                                command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = item.ModifiedDate;
                                command.Parameters.Add("@RowNumber", SqlDbType.BigInt).Value = item.RowNumber;
                                command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = item.LocationID;

                                command.ExecuteNonQuery();
                                transaction.Commit();
                            }
                            catch
                            {
                                transaction.Rollback();
                            }
                        }
                        #endregion querykhusus
                    }
                }

            }
            else
            {
                Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " detailsExtras == null , LPHID = " + id, Server.MapPath("~/Uploads/"), AccountEmployeeID);
            }

            if (!flagrecover)
                ExecuteQuery(string.Format("UPDATE PPLPHExtras SET IsDeleted = 1 WHERE Shift='DELJSON2' AND LPHID = {0};", id));
            else
                ExecuteQuery(string.Format("UPDATE PPLPHExtras SET Shift='' WHERE Shift='DELJSON2' AND ValueType='JSON' AND LPHID = {0};", id));
            #endregion            

            #region Update Values
            string checkDataComps = _ppLphComponentsAppService.FindBy("LPHID", id, true);
            List<PPLPHComponentsModel> dataComps = checkDataComps.DeserializeToPPLPHComponentList();
            dataComps = dataComps.OrderBy(x => x.ID).ToList();

            string shiftX = "";
            string dateX = "";
            string updateValues = string.Empty;

            if (dataComps.Count() == detailValue.Count())
            {
                for (int i = 0; i < dataComps.Count(); i++)
                {
                    var componentID = dataComps[i].ID;
                    PPLPHValuesModel valueModel = _ppLphValuesAppService.GetByNoTracking("LPHComponentID", componentID, true).DeserializeToPPLPHValue();

                    string tempDate = "";

                    if (dataComps[i].ComponentName == "Information-Date")
                    {
                        if (valueModel.Value != detailValue[i].Value)
                        {
                            dateX = detailValue[i].Value;
                            tempDate = dateX.Trim();
                            tempDate = tempDate.Remove(0, 4);
                            tempDate = tempDate.Trim();
                            DateTime newDate = DateTime.ParseExact(tempDate, "dd-MMM-yy", CultureInfo.CurrentCulture);

                            string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<PPLPHApprovalsModel> approve2Model_list = app.DeserializeToPPLPHApprovalList();
                            foreach (var item in approve2Model_list)
                            {
                                item.Date = newDate;
                                item.ModifiedBy = AccountName;
                                item.ModifiedDate = DateTime.Now;

                                string dataAppSub = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                _ppLphApprovalAppService.Update(dataAppSub);
                                //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2208 " + dataAppSub.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                            }

                            PPLPHSubmissionsModel submModel = _ppLphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToPPLPHSubmissions();
                            submModel.Date = newDate;
                            submModel.ModifiedBy = AccountName;
                            submModel.ModifiedDate = DateTime.Now;

                            string dataSub = JsonHelper<PPLPHSubmissionsModel>.Serialize(submModel);
                            _ppLphSubmissionsAppService.Update(dataSub);
                            //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2222 " + dataSub.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                        }
                    }

                    if (dataComps[i].ComponentName == "Information-Shift")
                    {
                        if (valueModel.Value != detailValue[i].Value)
                        {
                            shiftX = detailValue[i].Value;

                            string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<PPLPHApprovalsModel> approve2Model_list = app.DeserializeToPPLPHApprovalList();
                            foreach (var item in approve2Model_list)
                            {
                                item.Shift = shiftX;
                                item.ModifiedBy = AccountName;
                                item.ModifiedDate = DateTime.Now;

                                string dataAppSub = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                _ppLphApprovalAppService.Update(dataAppSub);
                                //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2242 " + dataAppSub.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                            }

                            PPLPHSubmissionsModel submModel = _ppLphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToPPLPHSubmissions();
                            submModel.Shift = shiftX;
                            submModel.ModifiedBy = AccountName;
                            submModel.ModifiedDate = DateTime.Now;

                            string dataSub = JsonHelper<PPLPHSubmissionsModel>.Serialize(submModel);
                            _ppLphSubmissionsAppService.Update(dataSub);
                            //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2261 " + dataSub.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                        }
                    }

                    if (dataComps[i].ComponentName == "Information-TeamLeader")
                    {
                        if (valueModel.Value != detailValue[i].Value)
                        {
                            var approverEmployeeID = detailValue[i].Value;

                            //Do update in approval 'Approver' 
                            string app = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submitModel.ID.ToString(), true);
                            List<PPLPHApprovalsModel> approve2Model_list = app.DeserializeToPPLPHApprovalList();
                            foreach (var item in approve2Model_list)
                            {
                                if (valueModel.Value != null)
                                {
                                    if (item.ApproverEmployeeID.Equals(valueModel.Value))
                                    {
                                        item.ApproverID = GetUserIDByEmployeeId(approverEmployeeID, _userAppService);
                                        item.ApproverEmployeeID = approverEmployeeID;
                                    }
                                    else if (item.ApproverEmployeeID != null && !item.UserEmployeeID.Equals(valueModel.Value))
                                    {
                                        item.UserID = GetUserIDByEmployeeId(approverEmployeeID, _userAppService);
                                        item.UserEmployeeID = approverEmployeeID;
                                    }

                                    item.ModifiedBy = AccountName;
                                    item.ModifiedDate = DateTime.Now;

                                    string dataNewApprover = JsonHelper<PPLPHApprovalsModel>.Serialize(item);
                                    _ppLphApprovalAppService.Update(dataNewApprover);
                                    //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2294 " + dataNewApprover.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
                                }
                            }
                        }
                    }

                    if (valueModel != null ? valueModel.Value != detailValue[i].Value : detailValue[i].Value != null)
                    {
                        if (detailValue[i].ValueType == "ImageURL")
                        {
                            //if (detailValue[i].Value != null ? detailValue[i].Value.Length > 200 : false)
                            if (detailValue[i].Value != null ? detailValue[i].Value.ToString().ToLower().StartsWith("data:image") : false)
                            {
                                detailValue[i].Value = GetFileNamePP(detailValue[i]);
                            }
                            else
                            {
                                continue;
                            }
                        }

                        valueModel.Value = detailValue[i].Value;

                        updateValues += string.Format("UPDATE PPLPHValues SET Value = '{0}' WHERE ID = {1};", valueModel.Value, valueModel.ID);
                    }
                }

                if (!string.IsNullOrEmpty(updateValues))
                    ExecuteQuery(updateValues);
            }
            #endregion

            #region Approval
            PPLPHApprovalsModel approveModel = _ppLphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID).DeserializeToPPLPHApproval();
            approveModel.Status = "Approved";
            approveModel.ModifiedBy = AccountName;
            approveModel.ModifiedDate = DateTime.Now;
            approveModel.Notes = comment;
            approveModel.ApproverID = AccountID;

            string dataApproval = JsonHelper<PPLPHApprovalsModel>.Serialize(approveModel);
            _ppLphApprovalAppService.Update(dataApproval);
            #endregion

            #region Send Email
            string employeeIDRequestor = GetEmployeeIDByUserID(submitModel.UserID, _userAppService);
            EmployeeModel employeeRequestor = GetEmployeeByEmployeeId(employeeIDRequestor, _employeeAppService);

            if (!string.IsNullOrEmpty(employeeRequestor.Email))
            {
                await EmailSender.SendEmailLPHPPApprove(employeeRequestor.Email, employeeRequestor.FullName + " (" + employeeRequestor.EmployeeID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHID.ToString(), submitModel.LPHHeader.Replace("Controller", ""));
            }
            #endregion
        }
        #endregion

        #region ::Revise LPH PP::
        // Revised use this method (no reject term anymore)
        protected async Task RejectPPLPH(
            IUserAppService _userAppService,
            IEmployeeAppService _employeeAppService,
            IPPLPHSubmissionsAppService _ppLphSubmissionsAppService,
            IPPLPHApprovalsAppService _ppLphApprovalAppService, long id, string notes)
        {
            PPLPHSubmissionsModel submitModel = _ppLphSubmissionsAppService.GetBy("LPHID", id, true).DeserializeToPPLPHSubmissions();

            #region Session
            UserModel user = Session["UserLogon"] == null ? new UserModel() : (UserModel)Session["UserLogon"];
            long userAccountID = user.ID;
            string userEmpID = user.EmployeeID;
            long userSupervisorUserID = user.SupervisorUserID;
            string userSupervisorEmpID = user.SupervisorID;
            long userLocationID = user.LocationID.HasValue ? user.LocationID.Value : 0;
            string userAccountName = user.UserName;
            #endregion

            #region Approval
            string approvalC = _ppLphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID);
            List<PPLPHApprovalsModel> approvalCheck = approvalC.DeserializeToPPLPHApprovalList().OrderBy(x => x.Date).ToList();

            // Update status to 'revised'
            PPLPHApprovalsModel approveModel = _ppLphApprovalAppService.GetLastByNoTracking("LPHSubmissionID", submitModel.ID).DeserializeToPPLPHApproval();
            approveModel.Status = "Revised";
            approveModel.ModifiedBy = AccountName;
            approveModel.ModifiedDate = DateTime.Now;
            approveModel.Notes = notes;
            approveModel.ApproverID = AccountID;

            string dataApproval = JsonHelper<PPLPHApprovalsModel>.Serialize(approveModel);
            _ppLphApprovalAppService.Update(dataApproval);

            //Create new record with status draft (like first step)
            PPLPHApprovalsModel modelApp = new PPLPHApprovalsModel
            {
                LPHSubmissionID = submitModel.ID,
                UserID = approvalCheck.ElementAt(0).UserID,
                ModifiedBy = AccountName,
                ModifiedDate = DateTime.Now,
                Status = "Draft",
                Shift = approvalCheck.ElementAt(0).Shift,
                Date = approvalCheck.ElementAt(0).Date,
                LocationID = approvalCheck.ElementAt(0).LocationID,
                ApproverID = approvalCheck.ElementAt(0).ApproverID,
                ApproverEmployeeID = approvalCheck.ElementAt(0).ApproverEmployeeID,
                UserEmployeeID = approvalCheck.ElementAt(0).UserEmployeeID
            };

            _ppLphApprovalAppService.AddModel(modelApp);
            //Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID + " BaseController 2434 " + modelApp.ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
            #endregion

            #region Send Email
            string employeeIDRequestor = GetEmployeeIDByUserID(submitModel.UserID, _userAppService);
            EmployeeModel employeeRequestor = GetEmployeeByEmployeeId(employeeIDRequestor, _employeeAppService);
            EmployeeModel employeeUser = GetEmployeeByEmployeeId(user.EmployeeID, _employeeAppService);

            if (!string.IsNullOrEmpty(employeeRequestor.Email))
            {
                await EmailSender.SendEmailLPHPPRevise(employeeRequestor.Email, employeeUser.FullName + " (" + AccountEmployeeID + ")", GetLPHPPNameFromControllerName(submitModel.LPHHeader), submitModel.LPHID.ToString(), submitModel.LPHHeader.Replace("Controller", ""), submitModel.LPHID.ToString());
            }
            #endregion
        }
        #endregion

        #region ::Helpers::
        private string GetValuePP(PPLPHValuesModel detailValue, PPLPHValuesModel modelValue)
        {
            if (detailValue.ValueType == "ImageURL")
            {
                //if (detailValue.Value != null ? detailValue.Value.Length > 200 : false)
                if (detailValue.Value != null ? detailValue.Value.ToLower().StartsWith("data:image") : false)
                {
                    //Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                    //Random r = new Random();
                    //var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";
                    var fileName = Guid.NewGuid().ToString() + ".jpeg";

                    string filePath = Server.MapPath("~/Uploads/lph/pp/") + fileName;

                    var bytes = Convert.FromBase64String(detailValue.Value.Replace("data:image/jpeg;base64,", ""));
                    using (var imageFile = new FileStream(filePath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    modelValue.Value = fileName;
                }
                else
                {
                    modelValue.Value = "_no_image.png";
                }
            }
            else
            {
                modelValue.Value = detailValue.Value;
            }

            return modelValue.Value;
        }

        private string GetFileNamePP(PPLPHValuesModel detailValue)
        {
            //Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            //Random r = new Random();
            //var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";
            var fileName = Guid.NewGuid().ToString() + ".jpeg";

            string filePath = Server.MapPath("~/Uploads/lph/pp/") + fileName;

            var bytes = Convert.FromBase64String(detailValue.Value.Replace("data:image/jpeg;base64,", ""));
            using (var imageFile = new FileStream(filePath, FileMode.Create))
            {
                imageFile.Write(bytes, 0, bytes.Length);
                imageFile.Flush();
            }

            return fileName;
        }

        private string GetValueOnCreate(string controllerName, LPHExtrasModel item)
        {
            if (item.Value != null ? item.Value.Length > 200 : false)
            {
                var t = item.ValueType.Split('/');
                item.ValueType = t[0];
                var fileName = t[1];

                string filePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + fileName;

                var bytes = Convert.FromBase64String(item.Value.Replace("data:image/jpeg;base64,", ""));
                using (var imageFile = new FileStream(filePath, FileMode.Create))
                {
                    imageFile.Write(bytes, 0, bytes.Length);
                    imageFile.Flush();
                }

                item.Value = fileName;
            }
            else
            {
                item.Value = "_no_image.png";
            }

            return item.Value;
        }

        private string GetValueOnEdit(string controllerName, LPHExtrasModel item)
        {
            var t = item.ValueType.Split('/');
            item.ValueType = t[0];
            var fileName = t[1];
            if (t[2] == "New")
            {
                if (item.Value != null ? item.Value.Length > 200 : false)
                {
                    string filePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + fileName;

                    var bytes = Convert.FromBase64String(item.Value.Replace("data:image/jpeg;base64,", ""));
                    using (var imageFile = new FileStream(filePath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    item.Value = fileName;
                }
                else
                {
                    item.Value = "_no_image.png";
                }
            }

            return item.Value;
        }

        private string GetValue(string controllerName, LPHValuesModel detailValue)
        {
            if (detailValue.ValueType == "ImageURL")
            {
                //if (detailValue.Value != null ? detailValue.Value.Length > 200 : false)
                if (detailValue.Value != null ? detailValue.Value.ToString().ToLower().StartsWith("data:image") : false)
                {
                    //Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                    //Random r = new Random();
                    //var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";
                    var fileName = Guid.NewGuid().ToString() + ".jpeg";

                    string filePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + fileName;

                    var bytes = Convert.FromBase64String(detailValue.Value.Replace("data:image/jpeg;base64,", ""));
                    using (var imageFile = new FileStream(filePath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    detailValue.Value = fileName;
                }
                else
                {
                    detailValue.Value = "_no_image.png";
                }
            }

            return detailValue.Value;
        }

        private string GetValueOnApproval(string controllerName, LPHExtrasModel item)
        {
            var t = item.ValueType.Split('/');
            item.ValueType = t[0];
            var fileName = t[1];
            if (t[2] == "New")
            {
                if (item.Value != null ? item.Value.Length > 200 : false)
                {
                    string filePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + fileName;

                    var bytes = Convert.FromBase64String(item.Value.Replace("data:image/jpeg;base64,", ""));

                    using (var imageFile = new FileStream(filePath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    item.Value = fileName;
                }
                else
                {
                    item.Value = "_no_image.png";
                }
            }

            return item.Value;
        }

        private string GetFileName(string controllerName, LPHValuesModel detailValue, string valueModel)
        {
            //Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            //Random r = new Random();
            //var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";
            var fileName = Guid.NewGuid().ToString() + ".jpeg";


            string filePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + fileName;

            var bytes = Convert.FromBase64String(detailValue.Value.Replace("data:image/jpeg;base64,", ""));

            using (var imageFile = new FileStream(filePath, FileMode.Create))
            {
                imageFile.Write(bytes, 0, bytes.Length);
                imageFile.Flush();
            }

            //delete old file
            //if (valueModel != "_no_image.png")
            //{
            //    string oldFilePath = Server.MapPath("~/Uploads/lph/" + controllerName.ToLower().Replace("controller", "") + "/") + valueModel;
            //    if (System.IO.File.Exists(oldFilePath))
            //    {
            //        System.IO.File.Delete(oldFilePath);
            //    }
            //}

            return fileName;
        }

        public long GetUserIDByEmployeeId(string empID, IUserAppService _userAppService)
        {
            string findUser = _userAppService.GetBy("EmployeeID", empID, false);
            UserModel userModel = findUser.DeserializeToUser();
            return userModel.ID;
        }
        public string GetEmployeeIDByUserID(long userID, IUserAppService _userAppService)
        {
            string findUser = _userAppService.GetBy("ID", userID, false);
            UserModel userModel = findUser.DeserializeToUser();
            return userModel.EmployeeID;
        }
        public EmployeeModel GetEmployeeByEmployeeId(string empID, IEmployeeAppService _employeeAppService)
        {
            string findEmployee = _employeeAppService.GetBy("EmployeeID", empID, false);
            EmployeeModel employeeModel = findEmployee.DeserializeToEmployee();
            return employeeModel;
        }
        public string GetLPHPPNameFromControllerName(string controllerName)
        {
            if (controllerName == "LPHPrimaryKretekLineAddback")
                return "Kretek Line - Addback";
            else if (controllerName == "LPHPrimaryDiet")
                return "Intermediate Line - DIET";
            else if (controllerName == "LPHPrimaryCloveInfeedConditioning")
                return "Intermediate Line - Clove Feeding & DCCC";
            else if (controllerName == "LPHPrimaryCSFCutDryPacking")
                return "Intermediate Line - CSF Cut Dry & Packing";
            else if (controllerName == "LPHPrimaryCSFInfeedConditioning")
                return "Intermediate Line - CSF Feeding & DCCC";
            else if (controllerName == "LPHPrimaryCloveCutDryPacking")
                return "Intermediate Line - Clove Cut Dry & Packing";
            else if (controllerName == "LPHPrimaryCloveCutDryPacking")
                return "Intermediate Line - RTC";
            else if (controllerName == "LPHPrimaryKitchen")
                return "Intermediate Line - Casing Kitchen";
            else if (controllerName == "LPHPrimaryWhiteLineOTP")
                return "White Line OTP - Process Note";
            else if (controllerName == "LPHPrimaryKretekLineFeeding")
                return "Kretek Line - Feeding KR & RJ";
            else if (controllerName == "LPHPrimaryKretekLineConditioning")
                return "Kretek Line - DCCC KR & RJ";
            else if (controllerName == "LPHPrimaryKretekLineCuttingDrying")
                return "Kretek Line - Cut Dry";
            else if (controllerName == "LPHPrimaryKretekLinePacking")
                return "Kretek Line - Packing";
            else if (controllerName == "LPHPrimaryCresFeedingConditioning")
                return "Kretek Line - CRES Feeding & DCCC";
            else if (controllerName == "LPHPrimaryCresDryingPacking")
                return "Kretek Line - CRES Cut Dry & Packing";
            else if (controllerName == "LPHPrimaryWhiteLineFeedingWhite")
                return "White Line PMID - Feeding White";
            else if (controllerName == "LPHPrimaryWhiteLineDCCC")
                return "White Line PMID - DCCC";
            else if (controllerName == "LPHPrimaryWhiteLineCuttingFTD")
                return "White Line PMID - Cutting + FTD";
            else if (controllerName == "LPHPrimaryWhiteLineAddback")
                return "White Line PMID - Addback";
            else if (controllerName == "LPHPrimaryWhiteLinePackingWhite")
                return "White Line PMID - Packing White";
            else if (controllerName == "LPHPrimaryWhiteLineFeedingSPM")
                return "White Line PMID - Feeding SPM";
            else if (controllerName == "LPHPrimaryISWhiteFeeding")
                return "White Line PMID - Feeding IS White";
            else if (controllerName == "LPHPrimaryISWhiteCutDry")
                return "White Line PMID - Cut Dry IS White";
            else
                return controllerName.Replace("Controller", "");
        }
        public void ExecuteQuery(string query)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    SqlCommand command = new SqlCommand(query, connection, transaction);
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                }
            }
        }

        public bool ExecuteQuerySuper(string query)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    SqlCommand command = new SqlCommand(query, connection, transaction);
                    command.ExecuteNonQuery();
                    transaction.Commit();
                    return true;
                }
                catch
                {
                    transaction.Rollback();
                    return false;
                }
            }
        }
        public bool IsValidJson(string strInput)
        {
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                dynamic result = serializer.DeserializeObject(strInput);
                return true;
            }
            catch { return false; }
        }

        public void refGroup(IReferenceAppService _referenceAppService, IReferenceDetailAppService _referenceDetailAppService)
        {
            string RefGroup = _referenceAppService.GetBy("Name", "GroupName");
            ReferenceModel referenceGroup = RefGroup.DeserializeToReference();
            ViewBag.referenceGroup = null;
            if (referenceGroup != null)
            {
                string stringRefGroup = _referenceDetailAppService.FindBy("ReferenceID", referenceGroup.ID);

                List<ReferenceDetailModel> referenceDetails = stringRefGroup.DeserializeToRefDetailList();
                ViewBag.referenceGroup = referenceDetails.Where(x => x.Code != "-").ToList();
            }
        }
        #endregion
    }
}