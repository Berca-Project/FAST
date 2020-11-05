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
	[CustomAuthorize("role")]
	public class RoleMasterController : BaseController<RoleModel>
	{
		private readonly IRoleAppService _roleAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public RoleMasterController(IRoleAppService roleAppService, IMenuAppService menuService, ILoggerAppService logger)
		{
			_roleAppService = roleAppService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: Role
		public ActionResult Index()
		{
			RoleModel model = new RoleModel();
			model.Access = GetAccess(WebConstants.MenuSlug.ROLE, _menuService);

			return View(model);
		}

		// GET: Role/Details/5
		public ActionResult Details(int id)
		{
			return View();
		}

		// GET: Role/Create
		public ActionResult Create()
		{
			return View();
		}

		// POST: Role/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(RoleModel roleModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return RedirectToAction("Index");
				}

				string exist = _roleAppService.GetBy("Name", roleModel.Name);
				if (!string.IsNullOrEmpty(exist))
				{
					ViewBag.Result = false;
					ViewBag.ErrorMessage = string.Format(UIResources.DataExist, "Role", roleModel.Name);

					return View(roleModel);
				}

				roleModel.ModifiedBy = AccountName;
				roleModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<RoleModel>.Serialize(roleModel);

				_roleAppService.Add(data);

				ViewBag.Result = true;
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Edit(string id)
		{
			RoleModel role = GetRole(id);

			return PartialView(role);
		}

		[HttpPost]
		public ActionResult Edit(RoleModel roleModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return RedirectToAction("Index");
				}

				roleModel.ModifiedBy = AccountName;
				roleModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<RoleModel>.Serialize(roleModel);

				_roleAppService.Update(data);

				ViewBag.Result = true;
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		[HttpPost]
		public ActionResult Delete(string name)
		{
			try
			{
				RoleModel role = GetRole(name);
				role.IsDeleted = true;

				string userData = JsonHelper<RoleModel>.Serialize(role);
				_roleAppService.Update(userData);

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
			string roleList = _roleAppService.GetAll(true);
			List<RoleModel> roles = roleList.DeserializeToRoleList();

			int recordsTotal = roles.Count();

			// Search    
			if (!string.IsNullOrEmpty(searchValue))
			{
				roles = roles.Where(m => m.Name.ToLower().Contains(searchValue.ToLower()) ||
											  m.Description.ToLower().Contains(searchValue.ToLower())).ToList();
			}

			if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
			{
				if (sortColumnDir == "asc")
				{
					switch (sortColumn.ToLower())
					{
						case "name":
							roles = roles.OrderBy(x => x.Name).ToList();
							break;
						case "description":
							roles = roles.OrderBy(x => x.Description).ToList();
							break;
						default:
							break;
					}
				}
				else
				{
					switch (sortColumn.ToLower())
					{
						case "name":
							roles = roles.OrderByDescending(x => x.Name).ToList();
							break;
						case "description":
							roles = roles.OrderByDescending(x => x.Description).ToList();
							break;
						default:
							break;
					}
				}
			}

			// total number of rows count     
			int recordsFiltered = roles.Count();

			// Paging     
			var data = roles.Skip(skip).Take(pageSize).ToList();

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		private RoleModel GetRole(string roleName)
		{
			string role = _roleAppService.GetByName(roleName, true);
			RoleModel model = role.DeserializeToRole();

			return model;
		}
	}
}
