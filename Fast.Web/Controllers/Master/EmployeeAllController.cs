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
	[RoutePrefix("Employees")]
	[CustomAuthorize("user")]
	public class EmployeeAllController : BaseController<EmployeeAllModel>
	{
		#region ::Init::
		private readonly IEmployeeAllAppService _empAllService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		#endregion

		#region ::Constructor::
		public EmployeeAllController(
			IEmployeeAllAppService empService,
			IMenuAppService menuService,
			IReferenceAppService referenceAppService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_empAllService = empService;
			_logger = logger;
			_menuService = menuService;
		}
		#endregion

		#region ::Public Methods::
		// GET: EmployeeAll
		public ActionResult Index()
		{
			GetTempData();

			EmployeeAllModel model = new EmployeeAllModel();
			model.Access = GetAccess(WebConstants.MenuSlug.EMPLOYEE_ALL, _menuService);

			return View(model);
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string empAllList = _empAllService.GetAll();
				List<EmployeeAllModel> empAllModelList = empAllList.DeserializeToEmployeeAllList();

				byte[] excelData = ExcelGenerator.ExportEmployeeAll(empAllModelList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Employee-All.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Edit(long id)
		{
			EmployeeAllModel emp = GetEmployeeAll(id);
			emp.GroupType = emp.GroupType == null ? "" : emp.GroupType.Trim();
			emp.GroupName = emp.GroupName == null ? "" : emp.GroupName.Trim();

			ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupTypeCode(_referenceAppService);
			ViewBag.GroupNameList = DropDownHelper.BindDropDownGroupName(_referenceAppService);

			return PartialView(emp);
		}

		[HttpPost]
		public ActionResult Edit(EmployeeAllModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<EmployeeAllModel>.Serialize(model);

				_empAllService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				EmployeeAllModel emp = GetEmployeeAll(id);
				emp.IsDeleted = true;

				string userData = JsonHelper<EmployeeAllModel>.Serialize(emp);
				_empAllService.Update(userData);

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
			string dataList = _empAllService.GetAll(true);
			List<EmployeeAllModel> empList = dataList.DeserializeToEmployeeAllList();

			int recordsTotal = empList.Count();

			// Search    //
			if (!string.IsNullOrEmpty(searchValue))
			{
				empList = empList.Where(m => m.EmployeeID != null && m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
											 m.BaseTownLocation != null && m.BaseTownLocation.ToLower().Contains(searchValue.ToLower()) ||
											 m.PositionDesc != null && m.PositionDesc.ToLower().Contains(searchValue.ToLower()) ||
											 m.CostCenter != null && m.CostCenter.ToLower().Contains(searchValue.ToLower()) ||
											 m.DepartmentDesc != null && m.DepartmentDesc.ToLower().Contains(searchValue.ToLower()) ||
											 m.FullName != null && m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
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
						case "costcenter":
							empList = empList.OrderBy(x => x.CostCenter).ToList();
							break;
						case "positiondesc":
							empList = empList.OrderBy(x => x.PositionDesc).ToList();
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
						case "costcenter":
							empList = empList.OrderByDescending(x => x.CostCenter).ToList();
							break;
						case "positiondesc":
							empList = empList.OrderByDescending(x => x.PositionDesc).ToList();
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
		#endregion

		#region ::Private Method::
		private EmployeeAllModel GetEmployeeAll(long id)
		{
			string emp = _empAllService.GetById(id, true);
			EmployeeAllModel model = emp.DeserializeToEmployeeAll();

			return model;
		}
		#endregion
	}
}
