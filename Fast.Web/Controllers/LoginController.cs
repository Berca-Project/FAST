#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Web.Mvc;
using System.Web.Security;
#endregion

namespace Fast.Web.Controllers
{
	public class LoginController : Controller
	{
		#region ::Init::
		private readonly IUserAppService _userAppService;
		private readonly IAccessRightAppService _accessRightAppService;
		private readonly IMenuAppService _menuAppService;
		private readonly IJobTitleAppService _jobTitleAppService;
		private readonly IChecklistApprovalAppService _checklistApprovalAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _empAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		#endregion

		#region ::Constructor::
		public LoginController(IUserAppService userAppService,
			IAccessRightAppService roleaccessRightAppService,
			IJobTitleAppService jobTitleAppService,
			IChecklistApprovalAppService checklistApprovalAppService,
			ILoggerAppService logger,
			IUserRoleAppService userRoleAppService,
			IEmployeeAppService empAppService,
			ILocationAppService locationAppService,
			IMenuAppService accessRightAppService)
		{
			_userRoleAppService = userRoleAppService;
			_jobTitleAppService = jobTitleAppService;
			_userAppService = userAppService;
			_accessRightAppService = roleaccessRightAppService;
			_checklistApprovalAppService = checklistApprovalAppService;
			_menuAppService = accessRightAppService;
			_logger = logger;
			_locationAppService = locationAppService;
			_empAppService = empAppService;
		}
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
		{
			if (Session["UserLogon"] == null)
			{
				bool usePageAuth = bool.Parse(ConfigurationManager.AppSettings["UsePageAuth"] ?? "false");
				if (usePageAuth)
				{
					long userId = 0;
					try
					{
						string userku = GetCurrentUser();

						//userku = "PMI\\f-skjlu22"; //Sampoerna-1
						if (userku.ToLower().Contains("\\f-"))
						{
							Session["IsFAccount"] = true;
							return View();
						}
						else
						{
							Session["IsFAccount"] = false;
						}

						string errorMessage = string.Empty;
						if (userku != "" && ValidateUsername(userku, out userId, out errorMessage))
						{
							return RedirectToAction("Index", "Home");
						}
						else
						{
							ViewBag.Message = string.IsNullOrEmpty(errorMessage) ? UIResources.InvalidUsername : errorMessage;
						}
					}
					catch (Exception ex)
					{
						_logger.LogError(ex.GetAllMessages(), userId);

						ViewBag.Message = ex.Message;

						Helper.LogErrorMessage("try catch ex.message: " + ex.Message.ToString(), Server.MapPath("~/Uploads/"));
					}

					Helper.LogErrorMessage("Keluar Unauthorized usePageAuth=true ketiga:", Server.MapPath("~/Uploads/"));

					return RedirectToAction("Index", "Unauthorized");
				}
				else
				{
					// usePageAuth == false, use login form to sign in
					return View();
				}
			}
			else
			{
				return RedirectToAction("Index", "Home");
			}
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult SignIn(UserModel model)
		{
			long userId = 0;
			try
			{
				Helper.LogErrorMessage("Dapat username sign in:" + model.UserName, Server.MapPath("~/Uploads/"));

				string errorMessage = string.Empty;
				if (ValidateUsername(model.UserName, out userId, out errorMessage))
				{
					return RedirectToAction("Index", "Home");
				}
				else
				{
					ViewBag.Message = string.IsNullOrEmpty(errorMessage) ? UIResources.InvalidUsername : errorMessage;
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), userId);

				Helper.LogErrorMessage("Exception try catch sign in:" + ex.GetAllMessages(), Server.MapPath("~/Uploads/"));

				ViewBag.Message = ex.Message;
			}

			return View("Index", model);
		}

		[HttpGet]
		public ActionResult SignOut()
		{
			UserModel user = (UserModel)Session["UserLogon"];
			try
			{
				HttpContext.Session.Clear();
				Session.Abandon();
				FormsAuthentication.SignOut();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message, user == null ? 0 : user.ID);
			}

			return RedirectToAction("Index", "Login");
		}
		#endregion

		#region ::Private Methods::
		private bool ValidateUsername(string username, out long userId, out string errorMessage)
		{
			List<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("UserName", username));
			filters.Add(new QueryFilter("EmployeeID", username, Operator.Equals, Operation.Or));

			string userStr = _userAppService.Get(filters);
			UserModel userModel = userStr.DeserializeToUser();
			userId = userModel.ID;

			if (string.IsNullOrEmpty(userStr))
			{
				if (username.Contains("f-"))
					errorMessage = UIResources.PleaseUseEmpId;
				else
					errorMessage = UIResources.InvalidUsername;
			}
			else if (!userModel.IsActive)
			{
				errorMessage = UIResources.UserInactive;
			}
			else
			{
				errorMessage = string.Empty;

				string jt = _jobTitleAppService.GetById(userModel.JobTitleID);
				JobTitleModel jtModel = jt.DeserializeToJobTitle();
				userModel.RoleName = jtModel.RoleName;
				userModel.JobTitle = jtModel.Title;

				bool isFAccount = Session["IsFAccount"] == null ? false : (bool)Session["IsFAccount"];
				if (isFAccount)
				{
					string[] allowedRoles = new string[] { "PRODTECH", "MECHANIC", "ELECTRICIAN", "GW", "SUPERVISOR" };
					if (!allowedRoles.ToList().Any(x => x == userModel.RoleName))
					{
						bool isAllowed = false;

						string userFRoles = _userRoleAppService.FindBy("UserID", userModel.ID, true);
						List<UserRoleModel> userFRoleList = userFRoles.DeserializeToUserRoleList();

						foreach (var item in userFRoleList)
						{
							if (allowedRoles.ToList().Any(x => x == item.RoleName))
							{
								isAllowed = true;
								break;
							}
						}

						if (!isAllowed)
						{
							errorMessage = UIResources.FAccountRolesInvalid;
							return false;
						}
					}
				}

				// set the supervisor user id if any
				if (!string.IsNullOrEmpty(userModel.SupervisorID))
				{
					string spv = _userAppService.GetBy("EmployeeID", userModel.SupervisorID);
					UserModel spvModel = spv.DeserializeToUser();
					userModel.SupervisorUserID = spvModel.ID;

					EmployeeModel emp2 = _empAppService.GetModelByEmpId(userModel.SupervisorID);
					userModel.SupervisorEmail = emp2.Email;
				}

				if (userModel.LocationID.HasValue)
				{
					string location = _locationAppService.GetById(userModel.LocationID.Value);
					LocationModel locationModel = location.DeserializeToLocation();
					if (locationModel.ParentID == 1)
					{
						userModel.ProdCenterID = locationModel.ID;
					}
					else
					{
						string loc = _locationAppService.GetById(locationModel.ParentID);
						LocationModel locModel = loc.DeserializeToLocation();
						if (locModel.ParentID == 1)
						{
							userModel.DepartmentID = locationModel.ID;
							userModel.ProdCenterID = locModel.ID;
						}
						else
						{
							string dep = _locationAppService.GetById(locModel.ID);
							LocationModel depModel = dep.DeserializeToLocation();
							userModel.DepartmentID = depModel.ID;
							userModel.ProdCenterID = depModel.ParentID;
						}
					}

					userModel.Location = _locationAppService.GetLocationFullCode(userModel.LocationID.Value);
				}

				TextInfo tinfo = new CultureInfo("en-US", false).TextInfo;
				userModel.UserName = tinfo.ToTitleCase(userModel.UserName.ToLower());
				userModel.RoleName = userModel.RoleName == null ? "" : tinfo.ToTitleCase(userModel.RoleName.ToLower());

				EmployeeModel emp = _empAppService.GetModelByEmpId(userModel.EmployeeID);
				userModel.Email = emp.Email;
				userModel.FullName = emp.FullName;

				string userRoles = _userRoleAppService.FindBy("UserID", userModel.ID, true);
				List<UserRoleModel> userRoleList = userRoles.DeserializeToUserRoleList();

				userModel.RoleNames.Add(userModel.RoleName);
				userModel.RoleNames.AddRange(userRoleList.Select(x => x.RoleName).ToList());
				userModel.RoleName = string.Join(",", userModel.RoleNames);

				List<MenuModel> Menu = GetMenuListByRoleNames(userModel.RoleNames, userModel.IsAdmin, userModel.LocationID.Value, userModel.ProdCenterID, userModel.DepartmentID);

				Helper.LogErrorMessage("Menu count:" + Menu.Count, Server.MapPath("~/Uploads/"));

				Session["UserLogon"] = userModel;
				Session["AuthMenu"] = Menu;


				return true;
			}

			return false;
		}

		private string GetCurrentUser()
		{
			string myDom = ConfigurationManager.AppSettings["MyDomain"];
			string userku = HttpContext.User.Identity.Name;
			userku = userku.Replace(myDom + "\\", "");
			userku = myDom + "\\" + userku;

			Helper.LogErrorMessage("Dapat useridentityname datanya: " + userku, Server.MapPath("~/Uploads/"));


			return userku;
		}

		private List<MenuModel> GetMenuListByRoleNames(List<string> roleNames, bool isAdmin, long locationID, long pcID, long depID)
		{
			if (isAdmin)
			{
				string accessRightList = _accessRightAppService.GetAll(true);
				IEnumerable<AccessRightDBModel> accessRightModels = accessRightList.DeserializeToAccessRightList();

				Session["AuthAccess"] = accessRightModels;

				string modules = _menuAppService.FindBy("IsActive", "true", true);
				List<MenuModel> moduleList = modules.DeserializeToMenuList().Where(x => !x.IsDeleted).OrderBy(x => x.DisplayOrder).ToList();

				return moduleList;
			}

			List<MenuModel> result = new List<MenuModel>();
			List<AccessRightDBModel> access = new List<AccessRightDBModel>();

			bool isAllUnchecked = false;

			foreach (var role in roleNames)
			{
				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("RoleName", role));
				filters.Add(new QueryFilter("LocationID", locationID.ToString()));
				filters.Add(new QueryFilter("IsDeleted", "0"));

				string accessRightList = _accessRightAppService.Find(filters);
				List<AccessRightDBModel> arList = accessRightList.DeserializeToAccessRightList();
				isAllUnchecked = arList.Count > 0 ? !arList.Any(x => (x.Read.HasValue && x.Read.Value) || (x.Write.HasValue && x.Write.Value) || (x.Print.HasValue && x.Print.Value)) : false;

				if (string.IsNullOrEmpty(accessRightList) || isAllUnchecked)
				{
					filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("RoleName", role));
					filters.Add(new QueryFilter("LocationID", depID.ToString()));
					filters.Add(new QueryFilter("IsDeleted", "0"));

					accessRightList = _accessRightAppService.Find(filters);
					arList = accessRightList.DeserializeToAccessRightList();
					isAllUnchecked = arList.Count > 0 ? !arList.Any(x => (x.Read.HasValue && x.Read.Value) || (x.Write.HasValue && x.Write.Value) || (x.Print.HasValue && x.Print.Value)) : false;

					if (string.IsNullOrEmpty(accessRightList) || isAllUnchecked)
					{
						filters = new List<QueryFilter>();
						filters.Add(new QueryFilter("RoleName", role));
						filters.Add(new QueryFilter("LocationID", pcID.ToString()));
						filters.Add(new QueryFilter("IsDeleted", "0"));

						accessRightList = _accessRightAppService.Find(filters);
						arList = accessRightList.DeserializeToAccessRightList();
						isAllUnchecked = arList.Count > 0 ? !arList.Any(x => (x.Read.HasValue && x.Read.Value) || (x.Write.HasValue && x.Write.Value) || (x.Print.HasValue && x.Print.Value)) : false;

						if (string.IsNullOrEmpty(accessRightList) || isAllUnchecked)
						{
							filters = new List<QueryFilter>();
							filters.Add(new QueryFilter("RoleName", role));
							filters.Add(new QueryFilter("LocationID", 1));
							filters.Add(new QueryFilter("IsDeleted", "0"));

							accessRightList = _accessRightAppService.Find(filters);
						}
					}
				}

				List<AccessRightDBModel> accessRightModels = accessRightList.DeserializeToAccessRightList();
				accessRightModels = accessRightModels.Where(x => x.hasRead || x.hasWrite || x.hasPrint).ToList();

				foreach (var item in accessRightModels)
				{
					var temp = access.Where(x => x.MenuID == item.MenuID).FirstOrDefault();
					if (temp != null)
					{
						if (item.Read.HasValue && item.Read.Value)
							temp.Read = item.Read.Value;
						if (item.Write.HasValue && item.Write.Value)
							temp.Write = item.Write.Value;
						if (item.Print.HasValue && item.Print.Value)
							temp.Print = item.Print.Value;
					}
					else
					{
						access.Add(item);
					}
				}
			}

			Session["AuthAccess"] = access;

			List<long> menuIDList = access.Select(x => x.MenuID).ToList();

			string menus = _menuAppService.FindBy("IsActive", "true", true);
			List<MenuModel> menuList = menus.DeserializeToMenuList();

			result = menuList.Where(y => menuIDList.Any(x => x == y.ID) && !y.IsDeleted).OrderBy(x => x.DisplayOrder).ToList();

			return result;
		}
		#endregion
	}
}