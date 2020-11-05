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
	[CustomAuthorize("user")]
	public class UserRoleController : BaseController<UserRoleModel>
	{
		private readonly ILocationAppService _locationAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IRoleAppService _roleAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IUserAppService _userAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _employeeService;
		private readonly IMenuAppService _menuService;
		private readonly IJobTitleAppService _jobTitleAppService;

		public UserRoleController(
			IUserRoleAppService userRoleAppService,
			IUserAppService userAppService,
			ILoggerAppService logger,
			ILocationAppService locationAppService,
			IJobTitleAppService jobTitleAppService,
			IRoleAppService roleAppService,
			IEmployeeAppService empService,
			IMenuAppService menuService,
			IReferenceAppService referenceAppService)
		{
			_userRoleAppService = userRoleAppService;
			_referenceAppService = referenceAppService;
			_jobTitleAppService = jobTitleAppService;
			_userAppService = userAppService;
			_menuService = menuService;
			_logger = logger;
			_roleAppService = roleAppService;
			_employeeService = empService;
			_locationAppService = locationAppService;
		}

		// GET: UserRole
		public ActionResult Index()
		{
			GetTempData();

			ViewBag.RoleList = DropDownHelper.BuildEmptyList();
			ViewBag.UserList = DropDownHelper.BuildEmptyList();

			UserRoleModel model = new UserRoleModel();
			model.Access = GetAccess(WebConstants.MenuSlug.USER_ROLE, _menuService);

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
				string userRoleList = _userRoleAppService.GetAll(true);
				List<UserRoleModel> userRoleModelList = userRoleList.DeserializeToUserRoleList();

				List<UserRoleModel> result = new List<UserRoleModel>();

				foreach (var item in userRoleModelList)
				{
					UserRoleModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
					if (exist == null)
					{
						string user = _userAppService.GetById(item.UserID, true);
						UserModel userModel = user.DeserializeToUser();
						string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID, true);
						item.Employee = emp.DeserializeToEmployee();
						item.RoleList = item.RoleName;

						result.Add(item);
					}
					else
					{
						exist.RoleList = exist.RoleList + ", " + item.RoleName;
					}
				}

				byte[] excelData = ExcelGenerator.ExportMasterUserRole(result, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-UserRole.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// GET: UserRole/Create
		public ActionResult Create()
		{
			ViewBag.RoleList = DropDownHelper.BindDropDownMultiRole(_roleAppService);
			ViewBag.UserList = GetUserList();

			UserRoleModel model = new UserRoleModel();
			model.Access = GetAccess(WebConstants.MenuSlug.USER_ROLE, _menuService);

			return PartialView(model);
		}

		// POST: UserRole/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(UserRoleModel userRoleModel)
		{
			try
			{
				ViewBag.RoleList = DropDownHelper.BuildEmptyList();
				ViewBag.UserList = DropDownHelper.BuildEmptyList();

				userRoleModel.Access = GetAccess(WebConstants.MenuSlug.USER_ROLE, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (string.IsNullOrEmpty(userRoleModel.EmployeeID))
				{
					SetFalseTempData("Please select user first");
					return RedirectToAction("Index");
				}

				if (userRoleModel.RoleNames == null || userRoleModel.RoleNames.Count() == 0)
				{
					SetFalseTempData("Please select extra role first");
					return RedirectToAction("Index");
				}

				string userT = _userAppService.GetBy("EmployeeID", userRoleModel.EmployeeID);
				UserModel userModelT = userT.DeserializeToUser();

				string jt = _jobTitleAppService.GetById(userModelT.JobTitleID);
				JobTitleModel jtModel = jt.DeserializeToJobTitle();
				string defaultRoleName = jtModel.RoleName;

				if (userRoleModel.RoleNames.Any(x => x == defaultRoleName))
				{
					SetFalseTempData(string.Format(UIResources.ExtraRoleInvalid, userModelT.UserName, defaultRoleName));
					return RedirectToAction("Index");
				}

				userRoleModel.UserID = userModelT.ID;

				foreach (var item in userRoleModel.RoleNames)
				{
					ICollection<QueryFilter> filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("RoleName", item.ToString()));
					filters.Add(new QueryFilter("UserID", userRoleModel.UserID.ToString()));
					filters.Add(new QueryFilter("IsDeleted", "0"));

					string exist = _userRoleAppService.Get(filters, true);
					if (!string.IsNullOrEmpty(exist))
					{
						string user = _userAppService.GetById(userRoleModel.UserID, true);
						UserModel userModel = user.DeserializeToUser();
						string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID);
						EmployeeModel empModel = emp.DeserializeToEmployee();

						SetFalseTempData(string.Format(UIResources.DataExist, empModel.FullName, item));
						return RedirectToAction("Index");
					}
				}

				foreach (var roleName in userRoleModel.RoleNames)
				{
					UserRoleModel newEntity = new UserRoleModel();
					newEntity.UserID = userRoleModel.UserID;
					newEntity.RoleName = roleName;
					newEntity.ModifiedBy = AccountName;
					newEntity.ModifiedDate = DateTime.Now;

					string data = JsonHelper<UserRoleModel>.Serialize(newEntity);

					_userRoleAppService.Add(data);
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

		// GET: UserRole/Edit/5
		public ActionResult Edit(int id)
		{
			UserRoleModel userRole = GetUserRole(id);
			string userRoles = _userRoleAppService.FindByNoTracking("UserID", userRole.UserID.ToString(), true);
			List<UserRoleModel> userRoleModelList = userRoles.DeserializeToUserRoleList();
			List<string> roleNameList = userRoleModelList.Select(c => c.RoleName).Distinct().ToList();

			ViewBag.RoleList = DropDownHelper.BindDropDownMultiRole(_roleAppService, roleNameList);
			ViewBag.UserList = DropDownHelper.BuildEmptyList();

			return PartialView(userRole);
		}

		// POST: UserRole/Edit/5
		[HttpPost]
		public ActionResult Edit(UserRoleModel userRoleModel)
		{
			try
			{
				ViewBag.RoleList = DropDownHelper.BindDropDownMultiRole(_roleAppService);
				ViewBag.UserList = GetUserList();

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				userRoleModel.ModifiedBy = AccountName;
				userRoleModel.ModifiedDate = DateTime.Now;

				string userRoles = _userRoleAppService.FindByNoTracking("UserID", userRoleModel.UserID.ToString(), true);
				List<UserRoleModel> userRoleList = userRoles.DeserializeToUserRoleList();

				string userT = _userAppService.GetById(userRoleModel.UserID);

				UserModel userModelT = userT.DeserializeToUser();
				string jt = _jobTitleAppService.GetById(userModelT.JobTitleID);
				JobTitleModel jtModel = jt.DeserializeToJobTitle();
				string defaultRoleName = jtModel.RoleName;

				if (userRoleModel.RoleNames.Any(x => x == defaultRoleName))
				{
					SetFalseTempData(string.Format(UIResources.ExtraRoleInvalid, userModelT.UserName, defaultRoleName));
					return RedirectToAction("Index");
				}

				foreach (var item in userRoleList)
				{
					if (!userRoleModel.RoleNames.Any(x => x == item.RoleName))
					{
						// remove if not selected						
						_userRoleAppService.Remove(item.ID);
					}
				}

				foreach (var item in userRoleModel.RoleNames)
				{
					if (!userRoleList.Any(x => x.RoleName == item))
					{
						UserRoleModel newEntity = new UserRoleModel();
						newEntity.UserID = userRoleModel.UserID;
						newEntity.RoleName = item;
						newEntity.ModifiedBy = AccountName;
						newEntity.ModifiedDate = DateTime.Now;

						string data = JsonHelper<UserRoleModel>.Serialize(newEntity);

						_userRoleAppService.Add(data);
					}
				}

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			userRoleModel.Access = GetAccess(WebConstants.MenuSlug.USER_ROLE, _menuService);

			return RedirectToAction("Index");
		}

		// GET: UserRole/Delete/5
		public ActionResult Delete(int id)
		{
			return View();
		}

		// POST: UserRole/Delete/5
		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				List<UserRoleModel> userRoles = GetUserRoleListByUserID(id);
				foreach (var item in userRoles)
				{
					_userRoleAppService.Remove(item.ID);
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
				string userRoleList = _userRoleAppService.GetAll(true);
				List<UserRoleModel> userRoleModelList = userRoleList.DeserializeToUserRoleList();

				List<UserRoleModel> result = new List<UserRoleModel>();

				foreach (var item in userRoleModelList)
				{
					UserRoleModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
					if (exist == null)
					{
						string user = _userAppService.GetById(item.UserID, true);
						UserModel userModel = user.DeserializeToUser();
						string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID, true);
						item.Employee = emp.DeserializeToEmployee();
						item.RoleList = item.RoleName;
						item.UserName = userModel.UserName;
						item.LocationID = userModel.LocationID.HasValue ? userModel.LocationID.Value : 0;

						result.Add(item);
					}
					else
					{
						exist.RoleList = exist.RoleList + ", " + item.RoleName;
					}
				}

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
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

				int recordsTotal = result.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					searchValue = searchValue.ToLower();

					result = result.Where(m => m.UserName != null && m.UserName.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Employee.EmployeeID != null && m.Employee.EmployeeID.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Employee.PositionDesc != null && m.Employee.PositionDesc.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Employee.FullName != null && m.Employee.FullName.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Location != null && m.Location.ToLower().Contains(searchValue) ||
											   m.RoleList != null && m.RoleList.ToLower().Contains(searchValue)).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "username":
								result = result.OrderBy(x => x.UserName).ToList();
								break;
							case "employeeid":
								result = result.OrderBy(x => x.Employee.EmployeeID).ToList();
								break;
							case "fullname":
								result = result.OrderBy(x => x.Employee.FullName).ToList();
								break;
							case "positiondesc":
								result = result.OrderBy(x => x.Employee.PositionDesc).ToList();
								break;
							case "location":
								result = result.OrderBy(x => x.Location).ToList();
								break;
							case "UserName":
								result = result.OrderBy(x => x.Employee.FullName).ToList();
								break;
							case "rolelist":
								result = result.OrderBy(x => x.RoleList).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "username":
								result = result.OrderByDescending(x => x.UserName).ToList();
								break;
							case "employeeid":
								result = result.OrderByDescending(x => x.Employee.EmployeeID).ToList();
								break;
							case "fullname":
								result = result.OrderByDescending(x => x.Employee.FullName).ToList();
								break;
							case "positiondesc":
								result = result.OrderByDescending(x => x.Employee.PositionDesc).ToList();
								break;
							case "location":
								result = result.OrderByDescending(x => x.Location).ToList();
								break;
							case "UserName":
								result = result.OrderByDescending(x => x.Employee.FullName).ToList();
								break;
							case "rolelist":
								result = result.OrderByDescending(x => x.RoleList).ToList();
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

				return Json(new { data = new List<UserRoleModel>() }, JsonRequestBehavior.AllowGet);
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
				string userRoleList = _userRoleAppService.GetAll(true);
				List<UserRoleModel> userRoleModelList = userRoleList.DeserializeToUserRoleList();
				List<UserRoleModel> result = new List<UserRoleModel>();

				foreach (var item in userRoleModelList)
				{
					UserRoleModel exist = result.Where(x => x.UserID == item.UserID).FirstOrDefault();
					if (exist == null)
					{
						string user = _userAppService.GetById(item.UserID, true);
						UserModel userModel = user.DeserializeToUser();
						string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID, true);
						item.Employee = emp.DeserializeToEmployee();
						item.RoleList = item.RoleName;
						item.UserName = userModel.UserName;
						item.LocationID = userModel.LocationID.HasValue ? userModel.LocationID.Value : 0;

						result.Add(item);
					}
					else
					{
						exist.RoleList = exist.RoleList + ", " + item.RoleName;
					}
				}

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
				result = result.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
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

				int recordsTotal = result.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					searchValue = searchValue.ToLower();

					result = result.Where(m => m.UserName != null && m.UserName.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Employee.EmployeeID != null && m.Employee.EmployeeID.Contains(searchValue) ||
											   m.Employee != null && m.Employee.PositionDesc != null && m.Employee.PositionDesc.Contains(searchValue) ||
											   m.Employee != null && m.Employee.FullName != null && m.Employee.FullName.ToLower().Contains(searchValue) ||
											   m.Employee != null && m.Location != null && m.Location.ToLower().Contains(searchValue) ||
											   m.RoleList != null && m.RoleList.ToLower().Contains(searchValue)).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "username":
								result = result.OrderBy(x => x.UserName).ToList();
								break;
							case "employeeid":
								result = result.OrderBy(x => x.Employee.EmployeeID).ToList();
								break;
							case "fullname":
								result = result.OrderBy(x => x.Employee.FullName).ToList();
								break;
							case "positiondesc":
								result = result.OrderBy(x => x.Employee.PositionDesc).ToList();
								break;
							case "location":
								result = result.OrderBy(x => x.Location).ToList();
								break;
							case "UserName":
								result = result.OrderBy(x => x.Employee.FullName).ToList();
								break;
							case "rolelist":
								result = result.OrderBy(x => x.RoleList).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "username":
								result = result.OrderByDescending(x => x.UserName).ToList();
								break;
							case "employeeid":
								result = result.OrderByDescending(x => x.Employee.EmployeeID).ToList();
								break;
							case "fullname":
								result = result.OrderByDescending(x => x.Employee.FullName).ToList();
								break;
							case "positiondesc":
								result = result.OrderByDescending(x => x.Employee.PositionDesc).ToList();
								break;
							case "location":
								result = result.OrderByDescending(x => x.Location).ToList();
								break;
							case "UserName":
								result = result.OrderByDescending(x => x.Employee.FullName).ToList();
								break;
							case "rolelist":
								result = result.OrderByDescending(x => x.RoleList).ToList();
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

				return Json(new { data = new List<UserRoleModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private UserRoleModel GetUserRole(long userRoleID)
		{
			string user = _userRoleAppService.GetById(userRoleID, true);
			UserRoleModel model = user.DeserializeToUserRole();
			string userData = _userAppService.GetById(model.UserID);
			UserModel userModel = userData.DeserializeToUser();
			string emp = _employeeService.GetBy("EmployeeID", userModel.EmployeeID);
			model.Employee = emp.DeserializeToEmployee();
			model.UserName = userModel.UserName;

			return model;
		}

		private List<UserRoleModel> GetUserRoleListByUserID(long userID)
		{
			string userRoles = _userRoleAppService.FindByNoTracking("UserID", userID.ToString(), true);
			List<UserRoleModel> models = userRoles.DeserializeToUserRoleList();

			return models;
		}

		private List<SelectListItem> GetUserList()
		{
			List<SelectListItem> result = DropDownHelper.BindDropDownUser(_userAppService);
			//string userRoles = _userRoleAppService.GetAll(true);
			//List<UserRoleModel> userRoleList = userRoles.DeserializeToUserRoleList();
			//result = result.Where(x => !userRoleList.Any(y => y.UserID.ToString() == x.Value)).ToList();
			return result;
		}
	}
}
