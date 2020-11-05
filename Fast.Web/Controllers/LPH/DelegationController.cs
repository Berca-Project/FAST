#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.LPH;
using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Fast.Web.Resources;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.Script.Serialization;

#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("delegation")]
	public class DelegationController : BaseController<LPHModel>
	{
		#region ::Services::
		private readonly ILPHAppService _lphAppService;
		private readonly ILPHApprovalsAppService _lphApprovalAppService;
		private readonly ILPHComponentsAppService _lphComponentsAppService;
		private readonly ILPHLocationsAppService _lphLocationsAppService;
		private readonly ILPHValuesAppService _lphValuesAppService;
		private readonly ILPHValueHistoriesAppService _lphValueHistoriesAppService;
		private readonly ILPHExtrasAppService _lphExtrasAppService;
		private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IMenuAppService _menuAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IJobTitleAppService _jobTitleAppService;

		#endregion

		#region ::Constructor::
		public DelegationController(
		  ILPHAppService lphAppService,
		  ILPHComponentsAppService lphComponentsAppService,
		  ILPHLocationsAppService lphLocationsAppService,
		  ILPHValuesAppService lphValuesAppService,
		  ILPHApprovalsAppService lphApprovalsAppService,
		  ILPHValueHistoriesAppService lphValueHistoriesAppService,
		  ILPHExtrasAppService lphExtrasAppService,
		  ILoggerAppService logger,
		  IReferenceAppService referenceAppService,
		  ILPHSubmissionsAppService lPHSubmissionsAppService,
		  ILocationAppService locationAppService,
		  IEmployeeAppService employeeAppService,
		  IMenuAppService menuAppService,
		  IUserAppService userAppService,
		  IUserRoleAppService userRoleAppService,
          IJobTitleAppService jobTitleAppService)
		{
			_lphAppService = lphAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphLocationsAppService = lphLocationsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphValueHistoriesAppService = lphValueHistoriesAppService;
			_lphExtrasAppService = lphExtrasAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_menuAppService = menuAppService;
			_userAppService = userAppService;
			_userRoleAppService = userRoleAppService;
            _jobTitleAppService = jobTitleAppService;
		}
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
		{
			LPHModel model = new LPHModel();
            ViewBag.isSupervisor = isSupervisor();
            ViewBag.isSuperadmin = isSuperadmin();
            ListSupervisor();
            ListSupervisorPP();
            ListNotApproveLPHPP();
            ListNotApproveLPHSP();
			
			ViewBag.Me = AccountID;
            ViewBag.EmpIDMe = AccountEmployeeID;

            string findEmployee = _employeeAppService.GetBy("EmployeeID", AccountEmployeeID, false);
            EmployeeModel emp = findEmployee.DeserializeToEmployee();

            ViewBag.EmpNameMe = emp.FullName;

            return View(model);
		}
        private bool isSuperadmin()
        {
            return AccountRoleList.Contains("SUPERADMIN") || AccountRoleList.Contains("Superadmin");
        }
        private bool isSupervisor()
        {
            return AccountRoleList.Contains("SUPERVISOR") || AccountRoleList.Contains("Supervisor");
        }
        public void ListSupervisor()
        {
            var Department = AccountDepartmentID;
            List<SelectListItem> _menuList = new List<SelectListItem>();

            List<long> locationIdList = _locationAppService.GetAll(true).DeserializeToLocationList().Where(x => x.ParentCode == "SP" || x.Code == "SP").Select(x => x.ID).ToList();//_locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals("SUPERVISOR", StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            var TL = LPHSPHelper.GetTeamLeadList(employeeList, null, leaderPositionList, userList, locationIdList, extraEmployeeRoles);
            var Fore = LPHSPHelper.GetForemanList(employeeList, null, userList, locationIdList, extraEmployeeRoles);
            ViewBag.Leaders = TL.Concat(Fore).Distinct().ToList();
           
		}
        public void ListSupervisorPP()
        {
            var Department = AccountDepartmentID;
            List<SelectListItem> _menuList = new List<SelectListItem>();

            List<long> locationIdList = _locationAppService.GetAll(true).DeserializeToLocationList().Where(x=> x.ParentCode == "PP" || x.Code == "PP").Select(x => x.ID).ToList();//_locationAppService.GetLocIDListByLocType(Department, "department");

            string users = _userAppService.GetAll(true);
            List<UserModel> userList = users.DeserializeToUserList();

            string employee = _employeeAppService.GetAll(true);
            List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();
            List<EmployeeRoleModel> extraEmployeeRoles = LPHSPHelper.GetExtraRole(_userRoleAppService, _userAppService, employeeList);

            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList().Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();

            List<string> leaderPositionList = jobTitleList.Where(x => x.RoleName.Equals("SUPERVISOR", StringComparison.OrdinalIgnoreCase)).Select(x => x.Title).ToList();
            var TL = LPHSPHelper.GetTeamLeadList(employeeList, null, leaderPositionList, userList, locationIdList, extraEmployeeRoles);
            var Fore = LPHSPHelper.GetForemanList(employeeList, null, userList, locationIdList, extraEmployeeRoles);
            ViewBag.LeadersPP = TL.Concat(Fore).Distinct().ToList();
           
		}
        
        public void ListNotApproveLPHSP()
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            DataSet dset = new DataSet();
            string queryAle = @"select * from EmployeeProfiles where EmployeeID in (SELECT distinct(UserEmployeeID)
                                  FROM [LPHApprovals]
                                where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID not like '00%')";
            using (SqlConnection con = new SqlConnection(strConString))
            {

                SqlCommand cmd = new SqlCommand(queryAle, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dset);

                Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
            }

            List<EmployeeModel> empl = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToEmployeeList();
            ViewBag.LeadersLPHSPNotSubmited = empl;

        }
        public void ListNotApproveLPHPP()
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            DataSet dset = new DataSet();
            string queryAle = @"select * from EmployeeProfiles where EmployeeID in (SELECT distinct(UserEmployeeID)
                                  FROM [PPLPHApprovals]
                                where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID not like '00%')";
            using (SqlConnection con = new SqlConnection(strConString))
            {

                SqlCommand cmd = new SqlCommand(queryAle, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dset);

                Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
            }

            List<EmployeeModel> empl = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToEmployeeList();
            ViewBag.LeadersLPHPPNotSubmited = empl;

        }
        [HttpPost]
        public ActionResult DelegateLPHSP(string OldSPV,string NewSPV)
        {
            try
            {
                string findUser = _userAppService.GetBy("EmployeeID", NewSPV, false);
                UserModel user = findUser.DeserializeToUser();

                string findEmployee = _employeeAppService.GetBy("EmployeeID", NewSPV, false);
                EmployeeModel emp = findEmployee.DeserializeToEmployee();

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                string queryAle = @"Update [LPHApprovals] set ApproverEmployeeID = '" + NewSPV + @"', ApproverID='" + user.ID + @"', ApproverFullName='" + emp.FullName + @"', ModifiedBy='"+AccountName+ @"', ModifiedDate=CURRENT_TIMESTAMP  
                                    where LPHSubmissionID in (
                                    SELECT LPHSubmissionID
                                      FROM [LPHApprovals]
                                    where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID not like '00%'
                                    and UserEmployeeID=" + OldSPV + @"
                                    )and ApproverEmployeeID= '" + OldSPV + @"'

                                Update [LPHApprovals] set UserEmployeeID = '" + NewSPV + @"', UserID='" + user.ID + @"', UserFullName='" + emp.FullName + @"', ModifiedBy='" + AccountName + @"', ModifiedDate=CURRENT_TIMESTAMP
                                where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID = '" + OldSPV + @"'";
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    //Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
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
        public ActionResult DelegateLPHPP(string OldSPV,string NewSPV)
        {
            try
            {
                string findUser = _userAppService.GetBy("EmployeeID", NewSPV, false);
                UserModel user = findUser.DeserializeToUser();

                string findEmployee = _employeeAppService.GetBy("EmployeeID", NewSPV, false);
                EmployeeModel emp = findEmployee.DeserializeToEmployee();

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                string queryAle = @"Update [PPLPHApprovals] set ApproverEmployeeID = '" + NewSPV + @"', ApproverID='" + user.ID + @"', ApproverFullName='" + emp.FullName + @"', ModifiedBy='" + AccountName + @"', ModifiedDate=CURRENT_TIMESTAMP  
                                    where LPHSubmissionID in (
                                    SELECT LPHSubmissionID
                                      FROM [PPLPHApprovals]
                                    where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID not like '00%'
                                    and UserEmployeeID=" + OldSPV + @"
                                    )and ApproverEmployeeID= '" + OldSPV + @"'

                                Update [PPLPHApprovals] set UserEmployeeID = '" + NewSPV + @"', UserID='" + user.ID + @"', UserFullName='" + emp.FullName + @"', ModifiedBy='" + AccountName + @"', ModifiedDate=CURRENT_TIMESTAMP
                                where  ApproverEmployeeID is null and approverid<=0 and status!='Draft' and UserEmployeeID is not null and UserEmployeeID = '" + OldSPV + @"'";
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    //Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }
                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        private string DataTableToJSONWithJavaScriptSerializer(DataTable table)
        {
            JavaScriptSerializer jsSerializer = new JavaScriptSerializer();
            List<Dictionary<string, object>> parentRow = new List<Dictionary<string, object>>();
            Dictionary<string, object> childRow;
            foreach (DataRow row in table.Rows)
            {
                childRow = new Dictionary<string, object>();
                foreach (DataColumn col in table.Columns)
                {
                    childRow.Add(col.ColumnName, row[col]);
                }
                parentRow.Add(childRow);
            }
            return jsSerializer.Serialize(parentRow);
        }
        #endregion
    }
}