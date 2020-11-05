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
	[CustomAuthorize("blend")]
	public class BlendController : BaseController<BlendModel>
	{
		private readonly IBlendAppService _blendAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		private readonly IReferenceAppService _referenceAppService;
		public BlendController(
			IBlendAppService blendAppService,
			ILocationAppService locationAppService,
			IReferenceAppService referenceAppService,
			IMenuAppService menuService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_blendAppService = blendAppService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: Blend
		public ActionResult Index()
		{
			GetTempData();

			BlendModel model = new BlendModel();
			model.Access = GetAccess(WebConstants.MenuSlug.BLEND, _menuService);

			return View(model);
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string blends = _blendAppService.GetAll();
				List<BlendModel> blendList = blends.DeserializeToBlendList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in blendList)
				{
					if (item.LocationID.HasValue)
					{
						if (locationMap.ContainsKey(item.LocationID.Value))
						{
							string loc;
							locationMap.TryGetValue(item.LocationID.Value, out loc);
							item.Location = loc;
						}
						else
						{
							item.Location = _locationAppService.GetLocationFullCode(item.LocationID.Value);
							locationMap.Add(item.LocationID.Value, item.Location);
						}
					}
				}

				byte[] excelData = ExcelGenerator.ExportBlends(blendList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Blends.xlsx");
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

		public ActionResult Create()
		{
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

			return PartialView();
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

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(BlendModel blendModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (blendModel.PcID == 0)
				{
					SetFalseTempData(UIResources.ProdCenterIsMissing);
					return RedirectToAction("Index");
				}

				List<long> subDepIDList = new List<long>();
				List<LocationModel> subDepList = new List<LocationModel>();

				if (blendModel.LocationID == 0 || blendModel.LocationID == null)
				{
					if (blendModel.DeptID == 0)
					{
						string departments = _locationAppService.FindBy("ParentID", blendModel.PcID.Value, true);
						List<LocationModel> depList = departments.DeserializeToLocationList();

						foreach (var item in depList)
						{
							string subDeps = _locationAppService.FindBy("ParentID", item.ID, true);
							subDepList = subDeps.DeserializeToLocationList();
						}
					}
					else
					{
						string subDeps = _locationAppService.FindBy("ParentID", blendModel.DeptID.Value, true);
						subDepList = subDeps.DeserializeToLocationList();
					}

					subDepIDList = subDepList.Select(c => c.ID).Distinct().ToList();
				}
				else
				{
					subDepIDList.Add(blendModel.LocationID.Value);
				}

				if (subDepIDList.Count == 0)
				{
					SetFalseTempData(UIResources.NoSubDepDefined);
					return RedirectToAction("Index");
				}

				foreach (var subdepid in subDepIDList)
				{
					List<QueryFilter> blendFilter = new List<QueryFilter>();
					blendFilter.Add(new QueryFilter("Code", blendModel.Code));
					blendFilter.Add(new QueryFilter("LocationID", (int)subdepid));
					string blend = _blendAppService.Get(blendFilter, true);

					BlendModel oldBlend = blend.DeserializeToBlend();
					oldBlend.ModifiedDate = DateTime.Now;
					oldBlend.ModifiedBy = AccountName;
					oldBlend.LocationID = subdepid;
					oldBlend.IsActive = true;
					oldBlend.Code = blendModel.Code;
					oldBlend.Description = blendModel.Description;
					oldBlend.OpsToKg = blendModel.OpsToKg;
					string data = JsonHelper<BlendModel>.Serialize(oldBlend);

					if (oldBlend.ID == 0)
					{
						_blendAppService.Add(data);
						break;
					}
					else
					{
						_blendAppService.Update(data);
					}
				}

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Delete(long id)
		{
			try
			{
				_blendAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Edit(long id)
		{
			string blend = _blendAppService.GetById(id, true);
			BlendModel model = blend.DeserializeToBlend();

			long countryID = 0;
			long pcID = 0;
			long depID = 0;
			long subDepID = 0;
			model.Location = DropDownHelper.ExtractLocation(_locationAppService, model.LocationID, out countryID, out pcID, out depID, out subDepID);
			model.PcID = pcID;
			model.DeptID = depID;
			model.LocationID = subDepID;

			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);

			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

			if (!model.PcID.HasValue || model.PcID == 0)
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, model.PcID.Value);

			if (!model.DeptID.HasValue || model.DeptID == 0)
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, model.DeptID.Value);

			ViewBag.StatusList = DropDownHelper.BindDropDownUserStatus();

			return PartialView(model);
		}

		[HttpPost]
		public ActionResult Edit(BlendModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (model.DeptID == 0)
				{
					SetFalseTempData("No selected department");
					return RedirectToAction("Index");
				}

                model.LocationID = model.DeptID;
				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<BlendModel>.Serialize(model);
				_blendAppService.Update(data);

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
				string blends = _blendAppService.GetAll();
				List<BlendModel> blendList = blends.DeserializeToBlendList();

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");

				blendList = blendList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

				int recordsTotal = blendList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in blendList)
					{
						if (item.LocationID.HasValue)
						{
							if (locationMap.ContainsKey(item.LocationID.Value))
							{
								string loc;
								locationMap.TryGetValue(item.LocationID.Value, out loc);
								item.Location = loc;
							}
							else
							{
								item.Location = _locationAppService.GetLocationFullCode(item.LocationID.Value);
								locationMap.Add(item.LocationID.Value, item.Location);
							}
						}
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					blendList = blendList.Where(m => m.Code != null && m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
													 m.Description != null && m.Description.ToLower().Contains(searchValue.ToLower()) ||
													 m.Department != null && m.Department.ToLower().Contains(searchValue.ToLower()) ||
													 m.ProductionCenter != null && m.ProductionCenter.ToLower().Contains(searchValue.ToLower()) ||
													 m.Location != null && m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								blendList = blendList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								blendList = blendList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								blendList = blendList.OrderBy(x => x.Location).ToList();
								break;
							case "department":
								blendList = blendList.OrderBy(x => x.Department).ToList();
								break;
							case "productioncenter":
								blendList = blendList.OrderBy(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								blendList = blendList.OrderBy(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								blendList = blendList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								blendList = blendList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								blendList = blendList.OrderByDescending(x => x.Location).ToList();
								break;
							case "department":
								blendList = blendList.OrderByDescending(x => x.Department).ToList();
								break;
							case "productioncenter":
								blendList = blendList.OrderByDescending(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								blendList = blendList.OrderByDescending(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = blendList.Count();

				// Paging     
				var data = blendList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in data)
					{
						if (item.LocationID.HasValue)
						{
							if (locationMap.ContainsKey(item.LocationID.Value))
							{
								string loc;
								locationMap.TryGetValue(item.LocationID.Value, out loc);
								item.Location = loc;
							}
							else
							{
								item.Location = _locationAppService.GetLocationFullCode(item.LocationID.Value);
								locationMap.Add(item.LocationID.Value, item.Location);
							}
						}
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<BlendModel>() }, JsonRequestBehavior.AllowGet);
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
				string blends = _blendAppService.GetAll();
				List<BlendModel> blendList = blends.DeserializeToBlendList();

				int recordsTotal = blendList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in blendList)
					{
						if (item.LocationID.HasValue)
						{
							if (locationMap.ContainsKey(item.LocationID.Value))
							{
								string loc;
								locationMap.TryGetValue(item.LocationID.Value, out loc);
								item.Location = loc;
							}
							else
							{
								item.Location = _locationAppService.GetLocationFullCode(item.LocationID.Value);
								locationMap.Add(item.LocationID.Value, item.Location);
							}
						}
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					blendList = blendList.Where(m => m.Code != null && m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
													 m.Description != null && m.Description.ToLower().Contains(searchValue.ToLower()) ||
													 m.Department != null && m.Department.ToLower().Contains(searchValue.ToLower()) ||
													 m.ProductionCenter != null && m.ProductionCenter.ToLower().Contains(searchValue.ToLower()) ||
													 m.Location != null && m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								blendList = blendList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								blendList = blendList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								blendList = blendList.OrderBy(x => x.Location).ToList();
								break;
							case "department":
								blendList = blendList.OrderBy(x => x.Department).ToList();
								break;
							case "productioncenter":
								blendList = blendList.OrderBy(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								blendList = blendList.OrderBy(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								blendList = blendList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								blendList = blendList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								blendList = blendList.OrderByDescending(x => x.Location).ToList();
								break;
							case "department":
								blendList = blendList.OrderByDescending(x => x.Department).ToList();
								break;
							case "productioncenter":
								blendList = blendList.OrderByDescending(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								blendList = blendList.OrderByDescending(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = blendList.Count();

				// Paging     
				var data = blendList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in data)
					{
						if (item.LocationID.HasValue)
						{
							if (locationMap.ContainsKey(item.LocationID.Value))
							{
								string loc;
								locationMap.TryGetValue(item.LocationID.Value, out loc);
								item.Location = loc;
							}
							else
							{
								item.Location = _locationAppService.GetLocationFullCode(item.LocationID.Value);
								locationMap.Add(item.LocationID.Value, item.Location);
							}
						}
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<BlendModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
