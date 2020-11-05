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
	[CustomAuthorize("machine")]
	public class MachineTypeController : BaseController<ReferenceDetailModel>
	{
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		private readonly ILocationAppService _locAppService;
		private readonly ILocationMachineTypeAppService _locMachineTypeService;

		public MachineTypeController(
			IReferenceAppService referenceAppService,
			IMenuAppService menuService,
			ILocationAppService locAppService,
			ILocationMachineTypeAppService locMachineTypeService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_menuService = menuService;
			_logger = logger;
			_locMachineTypeService = locMachineTypeService;
			_locAppService = locAppService;
		}

		// GET: MachineType
		public ActionResult Index()
		{
			GetTempData();

			ViewBag.MachineTypeList = DropDownHelper.BuildMultiEmpty();
			ViewBag.CountryList = DropDownHelper.BuildEmptyList();
			ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

			LocationMachineTypeModel model = new LocationMachineTypeModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

			return View(model);
		}

		[HttpPost]
		public ActionResult GetProductionCenterByCountryID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetProductionCenterByCountryID(_locAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetDepartmentByProdCenterID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetSubDepartmentByDepartmentID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetSubDepartmentByDepartmentID(_locAppService, _referenceAppService, id);
			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		// GET: MachineType/Create
		public ActionResult Create()
		{
			ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService);
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

			return PartialView();
		}

		// POST: MachineType/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(LocationMachineTypeModel machineTypeModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (machineTypeModel.MachineTypeIDs.Length == 0)
				{
					SetFalseTempData(UIResources.MachineNotSelected);
					return RedirectToAction("Index");
				}

				if (machineTypeModel.ProdCenterID == 0)
				{
					SetFalseTempData(UIResources.LocationInvalid);
					return RedirectToAction("Index");
				}

				List<long> subDepIDList = new List<long>();
				List<LocationModel> subDepList = new List<LocationModel>();
				string locMTList = _locMachineTypeService.GetAll(true);
				List<LocationMachineTypeModel> locMTModelList = locMTList.DeserializeToLocationMachineTypeList();

				if (machineTypeModel.SubDepartmentID == 0)
				{
					if (machineTypeModel.DepartmentID == 0)
					{
						string departments = _locAppService.FindBy("ParentID", machineTypeModel.ProdCenterID, true);
						List<LocationModel> depList = departments.DeserializeToLocationList();

						foreach (var item in depList)
						{
							string subDeps = _locAppService.FindBy("ParentID", item.ID, true);
							subDepList = subDeps.DeserializeToLocationList();
						}
					}
					else
					{
						string subDeps = _locAppService.FindBy("ParentID", machineTypeModel.DepartmentID, true);
						subDepList = subDeps.DeserializeToLocationList();
					}

					subDepIDList = subDepList.Select(c => c.ID).Distinct().ToList();
				}
				else
				{
					subDepIDList.Add(machineTypeModel.SubDepartmentID);
				}

				if (subDepIDList.Count == 0)
				{
					SetFalseTempData(UIResources.NoSubDepDefined);
					return RedirectToAction("Index");
				}

				foreach (var machineTypeID in machineTypeModel.MachineTypeIDs)
				{
					foreach (var subdepid in subDepIDList)
					{
						if (!locMTModelList.Any(x => x.LocationID == subdepid && x.MachineTypeID == machineTypeID))
						{
							LocationMachineTypeModel newData = new LocationMachineTypeModel();
							newData.LocationID = subdepid;
							newData.MachineTypeID = machineTypeID;
							newData.ModifiedDate = DateTime.Now;
							newData.ModifiedBy = AccountName;

							string data = JsonHelper<LocationMachineTypeModel>.Serialize(newData);

							_locMachineTypeService.Add(data);
						}
					}
				}

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.InvalidModelState);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult ExportExcel()
		{
			try
			{
				// Getting all data    			
				string machineTypeList = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
				List<ReferenceDetailModel> machineTypeModelList = machineTypeList.DeserializeToRefDetailList();

				// Getting all data    			
				string locMachineTypes = _locMachineTypeService.GetAll(true);
				List<LocationMachineTypeModel> locMtList = locMachineTypes.DeserializeToLocationMachineTypeList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in locMtList)
				{
					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}

					var mt = machineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
					item.MachineType = mt == null ? string.Empty : mt.Code;
				}

				byte[] excelData = ExcelGenerator.ExportMasterMachineType(locMtList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-MachineType.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// GET: MachineType/Edit/5
		public ActionResult Edit(int id)
		{
			LocationMachineTypeModel model = GetMachineType(id);

			long countryID = 0;
			long pcID = 0;
			long depID = 0;
			long subDepID = 0;
			model.Location = DropDownHelper.ExtractLocation(_locAppService, model.LocationID, out countryID, out pcID, out depID, out subDepID);
			model.CountryID = countryID;
			model.ProdCenterID = pcID;
			model.DepartmentID = depID;
			model.SubDepartmentID = subDepID;

			ViewBag.CountryList = DropDownHelper.GetCountryList(_locAppService, _referenceAppService);

			if (model.CountryID == 0)
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locAppService, _referenceAppService);
			else
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterByCountryID(_locAppService, _referenceAppService, model.CountryID);

			if (model.ProdCenterID == 0)
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locAppService, _referenceAppService, model.ProdCenterID);

			if (model.DepartmentID == 0)
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locAppService, _referenceAppService, model.DepartmentID);

			ViewBag.MachineTypeSingleList = DropDownHelper.BindDropDownMachineType(_referenceAppService);

			return PartialView(model);
		}

		// POST: MachineType/Edit/5
		[HttpPost]
		public ActionResult Edit(LocationMachineTypeModel machineTypeModel)
		{
			try
			{
				machineTypeModel.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (machineTypeModel.SubDepartmentID == 0)
				{
					SetFalseTempData(UIResources.SubDepartmentIsMissing);
					return RedirectToAction("Index");
				}

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("LocationID", machineTypeModel.SubDepartmentID.ToString()));
				filters.Add(new QueryFilter("MachineTypeID", machineTypeModel.MachineTypeID.ToString()));
				filters.Add(new QueryFilter("IsDeleted", "0"));

				string exist = _locMachineTypeService.Get(filters, true);
				if (!string.IsNullOrEmpty(exist))
				{
					SetFalseTempData(UIResources.DataExist);
					return RedirectToAction("Index");
				}

				machineTypeModel.LocationID = machineTypeModel.SubDepartmentID;
				machineTypeModel.MachineTypeID = machineTypeModel.MachineTypeID;
				machineTypeModel.ModifiedBy = AccountName;
				machineTypeModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<LocationMachineTypeModel>.Serialize(machineTypeModel);

				_locMachineTypeService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// GET: MachineType/Delete/5
		public ActionResult Delete(int id)
		{
			return View();
		}

		// POST: MachineType/Delete/5
		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				LocationMachineTypeModel machineType = GetMachineType(id);
				machineType.IsDeleted = true;

				string userData = JsonHelper<LocationMachineTypeModel>.Serialize(machineType);
				_locMachineTypeService.Update(userData);

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

				string machineTypeList = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
				List<ReferenceDetailModel> machineTypeModelList = machineTypeList.DeserializeToRefDetailList();

				// Getting all data    			
				string locMachineTypes = _locMachineTypeService.GetAll(true);
				List<LocationMachineTypeModel> locMtList = locMachineTypes.DeserializeToLocationMachineTypeList();				

				int recordsTotal = locMtList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in locMtList)
					{
						if (locationMap.ContainsKey(item.LocationID))
						{
							string loc;
							locationMap.TryGetValue(item.LocationID, out loc);
							item.Location = loc;
						}
						else
						{
							item.Location = _locAppService.GetLocationFullCode(item.LocationID);
							locationMap.Add(item.LocationID, item.Location);
						}

						var mt = machineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						item.MachineType = mt == null ? string.Empty : mt.Code;
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					locMtList = locMtList.Where(m => m.Location.ToLower().Contains(searchValue.ToLower()) ||
													 m.MachineType.ToLower().Contains(searchValue.ToLower())).ToList();

				}		

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "location":
								locMtList = locMtList.OrderBy(x => x.Location).ToList();
								break;
							case "machinetype":
								locMtList = locMtList.OrderBy(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "location":
								locMtList = locMtList.OrderByDescending(x => x.Location).ToList();
								break;
							case "machinetype":
								locMtList = locMtList.OrderByDescending(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = locMtList.Count();

				// Paging     
				var data = locMtList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in data)
					{
						if (locationMap.ContainsKey(item.LocationID))
						{
							string loc;
							locationMap.TryGetValue(item.LocationID, out loc);
							item.Location = loc;
						}
						else
						{
							item.Location = _locAppService.GetLocationFullCode(item.LocationID);
							locationMap.Add(item.LocationID, item.Location);
						}

						var mt = machineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						item.MachineType = mt == null ? string.Empty : mt.Code;
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<LocationMachineTypeModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private LocationMachineTypeModel GetMachineType(long machineTypeID)
		{
			string machineType = _locMachineTypeService.GetById(machineTypeID, true);
			LocationMachineTypeModel model = machineType.DeserializeToLocationMachineType();

			return model;
		}
	}
}
