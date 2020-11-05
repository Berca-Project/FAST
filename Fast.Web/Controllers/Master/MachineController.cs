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
	public class MachineController : BaseController<MachineModel>
	{
		private readonly IMachineAppService _machineAppService;
		private readonly ILoggerAppService _logger;
		private readonly ILocationAppService _locationAppService;
		private readonly ILocationMachineTypeAppService _locationMachineTypeAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IMenuAppService _menuService;

		public MachineController(
			IMachineAppService machineAppService,
			ILoggerAppService logger,
			IMenuAppService menuService,
			ILocationMachineTypeAppService locationMachineTypeAppService,
			ILocationAppService locationAppService,
			IReferenceAppService referenceAppService)
		{
			_machineAppService = machineAppService;
			_locationAppService = locationAppService;
			_referenceAppService = referenceAppService;
			_menuService = menuService;
			_logger = logger;
			_locationMachineTypeAppService = locationMachineTypeAppService;
		}

		// GET: Machine
		public ActionResult Index()
		{
			GetTempData();

			MachineModel model = GetIndexModel();

			return View(model);
		}

		private MachineModel GetIndexModel()
		{
			ViewBag.MachineTypeList = DropDownHelper.BuildEmptyList();
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.LegalEntityList = DropDownHelper.BindDropDownLegalEntityDesc(_referenceAppService);
			ViewBag.MachineBrandList = DropDownHelper.BindDropDownMachineBrandDesc(_referenceAppService);
			ViewBag.AreaList = DropDownHelper.BindDropDownArea(_referenceAppService);
            ViewBag.ClusterList = DropDownHelper.BindDropDownCluster(_referenceAppService);

            MachineModel model = new MachineModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

			return model;
		}

		public ActionResult ExportExcel()
		{
			try
			{
				// Getting all data    			
				string machineList = _machineAppService.GetAll(true);
				List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

				Dictionary<long, string> machineMap = new Dictionary<long, string>();
				foreach (var item in machineModelList)
				{
					if (machineMap.ContainsKey(item.MachineTypeID))
					{
						string mt;
						machineMap.TryGetValue(item.MachineTypeID, out mt);
						item.MachineType = mt;
					}
					else
					{
						item.MachineType = GetMachineTypeCode(item.MachineTypeID);
						machineMap.Add(item.MachineTypeID, item.MachineType);
					}
				}

				machineModelList = machineModelList.OrderBy(x => x.Location).ToList();

				byte[] excelData = ExcelGenerator.ExportMasterMachine(machineModelList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Machines.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		#region ::Ajax Call::
		[HttpPost]
		public ActionResult GetMachineTypeByProductionCenterID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetMachineTypeByProductionCenterID(_locationAppService, _locationMachineTypeAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetMachineTypeByDepartmentID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetMachineTypeByDepartmentID(_locationAppService, _locationMachineTypeAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetMachineTypeBySubDepartmentID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetMachineTypeBySubDepartmentID(_locationMachineTypeAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetProductionCenterByCountryID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetDepartmentByProdCenterID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, id);
			List<SelectListItem> _mtMenuList = DropDownHelper.GetMachineTypeByProductionCenterID(_locationAppService, _locationMachineTypeAppService, _referenceAppService, id);
			if (_mtMenuList.Count > 0)
				_menuList.AddRange(_mtMenuList);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetSubDepartmentByDepartmentID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, id);
			List<SelectListItem> _mtMenuList = DropDownHelper.GetMachineTypeByDepartmentID(_locationAppService, _locationMachineTypeAppService, _referenceAppService, id);
			if (_mtMenuList.Count > 0)
				_menuList.AddRange(_mtMenuList);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}
		#endregion

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(MachineModel machineModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(ModelState.GetModelStateErrors());
					return RedirectToAction("Index");
				}

				if (machineModel.SubDepartmentID == 0)
				{
					SetFalseTempData(UIResources.SubDepartmentIsMissing);
					return RedirectToAction("Index");
				}

				machineModel.LocationID = machineModel.SubDepartmentID;
				machineModel.Location = _locationAppService.GetLocationFullCode(machineModel.SubDepartmentID);
				machineModel.ModifiedBy = AccountName;
				machineModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<MachineModel>.Serialize(machineModel);

				_machineAppService.Add(data);

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// GET: Machine/Edit/5
		public ActionResult Edit(int id)
		{
			ViewBag.MachineTypeList = DropDownHelper.BindDropDownMachineType(_referenceAppService);
			ViewBag.LegalEntityList = DropDownHelper.BindDropDownLegalEntityDesc(_referenceAppService);
			ViewBag.MachineBrandList = DropDownHelper.BindDropDownMachineBrandDesc(_referenceAppService);
			ViewBag.AreaList = DropDownHelper.BindDropDownArea(_referenceAppService);
            ViewBag.ClusterList = DropDownHelper.BindDropDownCluster(_referenceAppService);

            MachineModel model = GetMachine(id);

			long countryID = 0;
			long pcID = 0;
			long depID = 0;
			long subDepID = 0;
			model.Location = DropDownHelper.ExtractLocation(_locationAppService, model.LocationID, out countryID, out pcID, out depID, out subDepID);
			model.CountryID = countryID;
			model.ProdCenterID = pcID;
			model.DepartmentID = depID;
			model.SubDepartmentID = subDepID;

			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);

			if (model.CountryID == 0)
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			else
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, model.CountryID);

			if (model.ProdCenterID == 0)
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, model.ProdCenterID);

			if (model.DepartmentID == 0)
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			else
				ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, model.DepartmentID);

			return PartialView(model);
		}

		// POST: Machine/Edit/5
		[HttpPost]
		public ActionResult Edit(MachineModel machineModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(ModelState.GetModelStateErrors());
					return RedirectToAction("Index");
				}

				if (machineModel.SubDepartmentID == 0)
				{
					SetFalseTempData(UIResources.SubDepartmentIsMissing);
					return RedirectToAction("Index");
				}

				machineModel.ModifiedBy = AccountName;
				machineModel.ModifiedDate = DateTime.Now;
				machineModel.LocationID = machineModel.SubDepartmentID;
				machineModel.Location = _locationAppService.GetLocationFullCode(machineModel.SubDepartmentID);

				string data = JsonHelper<MachineModel>.Serialize(machineModel);

				_machineAppService.Update(data);

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
				MachineModel machine = GetMachine(id);
				machine.IsDeleted = true;

				string machineData = JsonHelper<MachineModel>.Serialize(machine);
				_machineAppService.Update(machineData);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult DeleteMachineType(long id)
		{
			try
			{
				ReferenceDetailModel machine = GetMachineType(id);
				machine.IsDeleted = true;

				string machineType = JsonHelper<ReferenceDetailModel>.Serialize(machine);
				_referenceAppService.UpdateDetail(machineType);

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
				string machineList = _machineAppService.GetAll(true);
				List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

				int recordsTotal = machineModelList.Count();
				bool isLoaded = false;
				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> machineMap = new Dictionary<long, string>();
					foreach (var item in machineModelList)
					{
						if (machineMap.ContainsKey(item.MachineTypeID))
						{
							string mt;
							machineMap.TryGetValue(item.MachineTypeID, out mt);
							item.MachineType = mt;
						}
						else
						{
							item.MachineType = GetMachineTypeCode(item.MachineTypeID);
							machineMap.Add(item.MachineTypeID, item.MachineType);
						}
					}

					machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					machineModelList = machineModelList.Where(m => !string.IsNullOrEmpty(m.Location) && m.Location.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.Code) && m.Code.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineType) && m.MachineType.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineBrand) && m.MachineBrand.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.LegalEntity) && m.LegalEntity.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineSN) && m.MachineSN.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.SubProcess) && m.SubProcess.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "location":
								machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
								break;
							case "code":
								machineModelList = machineModelList.OrderBy(x => x.Code).ToList();
								break;
							case "machinebrand":
								machineModelList = machineModelList.OrderBy(x => x.MachineBrand).ToList();
								break;
							case "subprocess":
								machineModelList = machineModelList.OrderBy(x => x.SubProcess).ToList();
								break;
							case "legalentity":
								machineModelList = machineModelList.OrderBy(x => x.LegalEntity).ToList();
								break;
							case "machinesn":
								machineModelList = machineModelList.OrderBy(x => x.MachineSN).ToList();
								break;
							case "machinetype":
								machineModelList = machineModelList.OrderBy(x => x.MachineType).ToList();
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
								machineModelList = machineModelList.OrderByDescending(x => x.Location).ToList();
								break;
							case "code":
								machineModelList = machineModelList.OrderByDescending(x => x.Code).ToList();
								break;
							case "machinebrand":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineBrand).ToList();
								break;
							case "subprocess":
								machineModelList = machineModelList.OrderByDescending(x => x.SubProcess).ToList();
								break;
							case "legalentity":
								machineModelList = machineModelList.OrderByDescending(x => x.LegalEntity).ToList();
								break;
							case "machinesn":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineSN).ToList();
								break;
							case "machinetype":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = machineModelList.Count();

				// Paging     
				var data = machineModelList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> machineMap = new Dictionary<long, string>();
					foreach (var item in machineModelList)
					{
						if (machineMap.ContainsKey(item.MachineTypeID))
						{
							string mt;
							machineMap.TryGetValue(item.MachineTypeID, out mt);
							item.MachineType = mt;
						}
						else
						{
							item.MachineType = GetMachineTypeCode(item.MachineTypeID);
							machineMap.Add(item.MachineTypeID, item.MachineType);
						}
					}

					machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<MachineModel>() }, JsonRequestBehavior.AllowGet);
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
				string machineList = _machineAppService.GetAll(true);
				List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");

				machineModelList = machineModelList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

				int recordsTotal = machineModelList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> machineMap = new Dictionary<long, string>();
					foreach (var item in machineModelList)
					{
						if (machineMap.ContainsKey(item.MachineTypeID))
						{
							string mt;
							machineMap.TryGetValue(item.MachineTypeID, out mt);
							item.MachineType = mt;
						}
						else
						{
							item.MachineType = GetMachineTypeCode(item.MachineTypeID);
							machineMap.Add(item.MachineTypeID, item.MachineType);
						}
					}

					machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					machineModelList = machineModelList.Where(m => !string.IsNullOrEmpty(m.Location) && m.Location.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.Code) && m.Code.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineType) && m.MachineType.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineBrand) && m.MachineBrand.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.LegalEntity) && m.LegalEntity.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.MachineSN) && m.MachineSN.ToLower().Contains(searchValue.ToLower()) ||
																   !string.IsNullOrEmpty(m.SubProcess) && m.SubProcess.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "location":
								machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
								break;
							case "code":
								machineModelList = machineModelList.OrderBy(x => x.Code).ToList();
								break;
							case "machinebrand":
								machineModelList = machineModelList.OrderBy(x => x.MachineBrand).ToList();
								break;
							case "subprocess":
								machineModelList = machineModelList.OrderBy(x => x.SubProcess).ToList();
								break;
							case "legalentity":
								machineModelList = machineModelList.OrderBy(x => x.LegalEntity).ToList();
								break;
							case "machinesn":
								machineModelList = machineModelList.OrderBy(x => x.MachineSN).ToList();
								break;
							case "machinetype":
								machineModelList = machineModelList.OrderBy(x => x.MachineType).ToList();
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
								machineModelList = machineModelList.OrderByDescending(x => x.Location).ToList();
								break;
							case "code":
								machineModelList = machineModelList.OrderByDescending(x => x.Code).ToList();
								break;
							case "machinebrand":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineBrand).ToList();
								break;
							case "subprocess":
								machineModelList = machineModelList.OrderByDescending(x => x.SubProcess).ToList();
								break;
							case "legalentity":
								machineModelList = machineModelList.OrderByDescending(x => x.LegalEntity).ToList();
								break;
							case "machinesn":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineSN).ToList();
								break;
							case "machinetype":
								machineModelList = machineModelList.OrderByDescending(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = machineModelList.Count();

				// Paging     
				var data = machineModelList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> machineMap = new Dictionary<long, string>();
					foreach (var item in machineModelList)
					{
						if (machineMap.ContainsKey(item.MachineTypeID))
						{
							string mt;
							machineMap.TryGetValue(item.MachineTypeID, out mt);
							item.MachineType = mt;
						}
						else
						{
							item.MachineType = GetMachineTypeCode(item.MachineTypeID);
							machineMap.Add(item.MachineTypeID, item.MachineType);
						}
					}

					machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<MachineModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetAllMachineType()
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
			string machineTypeList = _referenceAppService.GetDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
			List<ReferenceDetailModel> machineTypeModelList = machineTypeList.DeserializeToRefDetailList();

			int recordsTotal = machineTypeModelList.Count();

			// Search    
			if (!string.IsNullOrEmpty(searchValue))
			{
				machineTypeModelList = machineTypeModelList.Where(m => m.Code.ToLower().Contains(searchValue.ToLower())).ToList();
			}

			if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
			{
				if (sortColumnDir == "asc")
				{
					switch (sortColumn.ToLower())
					{
						case "id":
							machineTypeModelList = machineTypeModelList.OrderBy(x => x.ID).ToList();
							break;
						case "code":
							machineTypeModelList = machineTypeModelList.OrderBy(x => x.Code).ToList();
							break;
						default:
							break;
					}
				}
				else
				{
					switch (sortColumn.ToLower())
					{
						case "id":
							machineTypeModelList = machineTypeModelList.OrderByDescending(x => x.ID).ToList();
							break;
						case "code":
							machineTypeModelList = machineTypeModelList.OrderByDescending(x => x.Code).ToList();
							break;
						default:
							break;
					}
				}
			}

			// total number of rows count     
			int recordsFiltered = machineTypeModelList.Count();

			// Paging     
			var data = machineTypeModelList.Skip(skip).Take(pageSize).ToList();

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		private MachineModel GetMachine(long machineID)
		{
			string machine = _machineAppService.GetById(machineID, true);
			MachineModel machineModel = machine.DeserializeToMachine();

			return machineModel;
		}

		private ReferenceDetailModel GetMachineType(long machineTypeID)
		{
			string machineType = _referenceAppService.GetDetailById(machineTypeID, true);
			ReferenceDetailModel machineTypeModel = machineType.DeserializeToRefDetail();

			return machineTypeModel;
		}

		private string GetMachineTypeCode(long machineTypeID)
		{
			return GetMachineType(machineTypeID).Code;
		}
	}
}
