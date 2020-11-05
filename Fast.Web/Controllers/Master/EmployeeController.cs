using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	public class EmployeeController : BaseController<EmployeeModel>
	{
		private readonly IEmployeeAppService _empService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public EmployeeController(IEmployeeAppService empService, IMenuAppService menuService, ILoggerAppService logger)
		{
			_empService = empService;
			_logger = logger;
			_menuService = menuService;
		}

		// GET: Employee
		public ActionResult Index()
		{
			EmployeeModel model = new EmployeeModel();
			model.Access = GetAccess(WebConstants.MenuSlug.EMPLOYEE, _menuService);

			return View(model);
		}

		// GET: Employee/Details/5
		public ActionResult Details(int id)
		{
			return View();
		}

		public ActionResult Edit(long id)
		{
			EmployeeModel emp = GetEmployee(id);

			return PartialView(emp);
		}

		[HttpPost]
		public ActionResult Edit(EmployeeModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return RedirectToAction("Index");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<EmployeeModel>.Serialize(model);

				_empService.Update(data);

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
		public ActionResult Delete(long id)
		{
			try
			{
				EmployeeModel emp = GetEmployee(id);
				emp.IsDeleted = true;

				string userData = JsonHelper<EmployeeModel>.Serialize(emp);
				_empService.Update(userData);

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
			string dataList = _empService.GetAll(true);
			List<EmployeeModel> empList = dataList.DeserializeToEmployeeList();

			int recordsTotal = empList.Count();

			// Search    
			if (!string.IsNullOrEmpty(searchValue))
			{
				empList = empList.Where(m => m.DepartmentDesc.ToLower().Contains(searchValue.ToLower()) ||
                                             m.UserName.ToLower().Contains(searchValue.ToLower()) || 
                                             m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
											 m.GroupName.ToLower().Contains(searchValue.ToLower()) ||
											 m.GroupType.ToLower().Contains(searchValue.ToLower()) ||
										     m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
			}

			if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
			{
				if (sortColumnDir == "asc")
				{
					switch (sortColumn.ToLower())
					{
						case "employeeid":
							empList = empList.OrderBy(x => x.EmployeeID).ToList();
							break;
						case "fullname":
							empList = empList.OrderBy(x => x.FullName).ToList();
							break;
						case "groupname":
							empList = empList.OrderBy(x => x.GroupName).ToList();
							break;
						case "grouptype":
							empList = empList.OrderBy(x => x.GroupType).ToList();
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
							empList = empList.OrderByDescending(x => x.EmployeeID).ToList();
							break;
						case "fullname":
							empList = empList.OrderByDescending(x => x.FullName).ToList();
							break;
						case "groupname":
							empList = empList.OrderByDescending(x => x.GroupName).ToList();
							break;
						case "grouptype":
							empList = empList.OrderByDescending(x => x.GroupType).ToList();
							break;
						default:
							break;
					}
				}
			}

			// total number of rows count     
			int recordsFiltered = empList.Count();

			// Paging     
			var data = empList.Skip(skip).Take(pageSize).ToList();

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		private EmployeeModel GetEmployee(long id)
		{
			string emp = _empService.GetById(id, true);
			EmployeeModel model = emp.DeserializeToEmployee();

			return model;
		}
	}
}
