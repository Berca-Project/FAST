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
	[CustomAuthorize("brand")]
	public class BrandController : BaseController<BrandModel>
	{
		private readonly ILocationAppService _locationAppService;
		private readonly IBrandAppService _brandAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public BrandController(
			IBrandAppService brandService,
			ILocationAppService locationAppService,
			IReferenceAppService referenceAppService,
			IMenuAppService menuService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_brandAppService = brandService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: Brand
		public ActionResult Index()
		{
			GetTempData();

			BrandModel model = new BrandModel();
			model.Access = GetAccess(WebConstants.MenuSlug.BRAND, _menuService);

			return View(model);
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string brands = _brandAppService.GetAll();
				List<BrandModel> brandList = brands.DeserializeToBrandList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in brandList)
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

				byte[] excelData = ExcelGenerator.ExportBrands(brandList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Brands.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
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
		[ValidateAntiForgeryToken]
		public ActionResult Create(BrandModel brandModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (brandModel.PcID == 0)
				{
					SetFalseTempData(UIResources.ProdCenterIsMissing);
					return RedirectToAction("Index");
				}

				List<long> subDepIDList = new List<long>();
				List<LocationModel> subDepList = new List<LocationModel>();

				if (brandModel.DeptID == 0)
				{
					string departments = _locationAppService.FindBy("ParentID", brandModel.PcID.Value, true);
					List<LocationModel> depList = departments.DeserializeToLocationList();

					foreach (var item in depList)
					{
						string subDeps = _locationAppService.FindBy("ParentID", item.ID, true);
						subDepList = subDeps.DeserializeToLocationList();
					}
				}
				else
				{
					string subDeps = _locationAppService.FindBy("ParentID", brandModel.DeptID.Value, true);
					subDepList = subDeps.DeserializeToLocationList();
				}

				subDepIDList = subDepList.Select(c => c.ID).Distinct().ToList();

				if (subDepIDList.Count == 0)
				{
					SetFalseTempData(UIResources.NoSubDepDefined);
					return RedirectToAction("Index");
				}

				foreach (var subdepid in subDepIDList)
				{
					List<QueryFilter> brandFilter = new List<QueryFilter>();
					brandFilter.Add(new QueryFilter("Code", brandModel.Code));
					brandFilter.Add(new QueryFilter("LocationID", (int)subdepid));
					string Brand = _brandAppService.Get(brandFilter, true);

					BrandModel oldBrand = Brand.DeserializeToBrand();
					oldBrand.ModifiedDate = DateTime.Now;
					oldBrand.ModifiedBy = AccountName;
					oldBrand.LocationID = subdepid;
					oldBrand.IsActive = true;
					oldBrand.Code = brandModel.Code;
					oldBrand.BeratCigarette = brandModel.BeratCigarette;
					oldBrand.PackToStick = brandModel.PackToStick;
					oldBrand.SlofToPack = brandModel.SlofToPack;
					oldBrand.BoxToSlof = brandModel.BoxToSlof;
					oldBrand.Description = brandModel.Description;
					oldBrand.CTW = brandModel.CTW;
					oldBrand.CTF = brandModel.CTF;
					oldBrand.RSCode = brandModel.RSCode;

					string data = JsonHelper<BrandModel>.Serialize(oldBrand);

					if (oldBrand.ID == 0)
					{
						_brandAppService.Add(data);
						break;
					}
					else
					{
						_brandAppService.Update(data);
					}
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

		public ActionResult Delete(long id)
		{
			try
			{
				_brandAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
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

		public ActionResult Edit(long id)
		{
			string brand = _brandAppService.GetById(id, true);
			BrandModel model = brand.DeserializeToBrand();

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
		public ActionResult Edit(BrandModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (model.PcID == 0)
				{
					SetFalseTempData(UIResources.ProdCenterIsMissing);
					return RedirectToAction("Index");
				}

				if (model.LocationID == 0)
				{
					if (model.DeptID == 0)
					{
						model.LocationID = model.PcID;
					}
					else
					{
						model.LocationID = model.DeptID;
					}
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<BrandModel>.Serialize(model);
				_brandAppService.Update(data);

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
				string brands = _brandAppService.GetAll();
				List<BrandModel> brandList = brands.DeserializeToBrandList();

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");

				brandList = brandList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

				int recordsTotal = brandList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in brandList)
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
					brandList = brandList.Where(m => m.Code != null && m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
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
								brandList = brandList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								brandList = brandList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								brandList = brandList.OrderBy(x => x.Location).ToList();
								break;
							case "department":
								brandList = brandList.OrderBy(x => x.Department).ToList();
								break;
							case "productioncenter":
								brandList = brandList.OrderBy(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								brandList = brandList.OrderBy(x => x.IsActive).ToList();
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
								brandList = brandList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								brandList = brandList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								brandList = brandList.OrderByDescending(x => x.Location).ToList();
								break;
							case "department":
								brandList = brandList.OrderByDescending(x => x.Department).ToList();
								break;
							case "productioncenter":
								brandList = brandList.OrderByDescending(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								brandList = brandList.OrderByDescending(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = brandList.Count();

				// Paging     
				var data = brandList.Skip(skip).Take(pageSize).ToList();

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

				return Json(new { data = new List<BrandModel>() }, JsonRequestBehavior.AllowGet);
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
				string brands = _brandAppService.GetAll();
				List<BrandModel> brandList = brands.DeserializeToBrandList();

				int recordsTotal = brandList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in brandList)
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
					brandList = brandList.Where(m => m.Code != null && m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
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
								brandList = brandList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								brandList = brandList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								brandList = brandList.OrderBy(x => x.Location).ToList();
								break;
							case "department":
								brandList = brandList.OrderBy(x => x.Department).ToList();
								break;
							case "productioncenter":
								brandList = brandList.OrderBy(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								brandList = brandList.OrderBy(x => x.IsActive).ToList();
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
								brandList = brandList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								brandList = brandList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								brandList = brandList.OrderByDescending(x => x.Location).ToList();
								break;
							case "department":
								brandList = brandList.OrderByDescending(x => x.Department).ToList();
								break;
							case "productioncenter":
								brandList = brandList.OrderByDescending(x => x.ProductionCenter).ToList();
								break;
							case "isactive":
								brandList = brandList.OrderByDescending(x => x.IsActive).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = brandList.Count();

				// Paging     
				var data = brandList.Skip(skip).Take(pageSize).ToList();

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

				return Json(new { data = new List<BrandModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
