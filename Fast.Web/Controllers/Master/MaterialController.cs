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
	[CustomAuthorize("materialcode")]
	public class MaterialCodeController : BaseController<MaterialCodeModel>
	{
		private readonly ILocationAppService _locationAppService;
		private readonly IMaterialCodeAppService _materialCodeAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public MaterialCodeController(
			IMaterialCodeAppService materialCodeService,
			ILocationAppService locationAppService,
			IReferenceAppService referenceAppService,
			IMenuAppService menuService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_materialCodeAppService = materialCodeService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: MaterialCode
		public ActionResult Index()
		{
			GetTempData();

			MaterialCodeModel model = new MaterialCodeModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MATERIAL_CODE, _menuService);

			return View(model);
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string materials = _materialCodeAppService.GetAll();
				List<MaterialCodeModel> materialList = materials.DeserializeToMaterialCodeList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in materialList)
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

				byte[] excelData = ExcelGenerator.ExportMaterialCodes(materialList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Material-Codes.xlsx");
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

		// POST: MachineType/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(MaterialCodeModel materialCodeModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (materialCodeModel.PcID == 0)
				{
					SetFalseTempData(UIResources.ProdCenterIsMissing);
					return RedirectToAction("Index");
				}

				List<long> subDepIDList = new List<long>();
				List<LocationModel> subDepList = new List<LocationModel>();

				if (materialCodeModel.LocationID == 0)
				{
					if (materialCodeModel.DeptID == 0)
					{
						string departments = _locationAppService.FindBy("ParentID", materialCodeModel.PcID.Value, true);
						List<LocationModel> depList = departments.DeserializeToLocationList();

						foreach (var item in depList)
						{
							string subDeps = _locationAppService.FindBy("ParentID", item.ID, true);
							subDepList = subDeps.DeserializeToLocationList();
						}
					}
					else
					{
						string subDeps = _locationAppService.FindBy("ParentID", materialCodeModel.DeptID.Value, true);
						subDepList = subDeps.DeserializeToLocationList();
					}

					subDepIDList = subDepList.Select(c => c.ID).Distinct().ToList();
				}
				else
				{
					subDepIDList.Add(materialCodeModel.LocationID.Value);
				}

				if (subDepIDList.Count == 0)
				{
					SetFalseTempData(UIResources.NoSubDepDefined);
					return RedirectToAction("Index");
				}

				foreach (var subdepid in subDepIDList)
				{
					List<QueryFilter> materialCodeFilter = new List<QueryFilter>();
					materialCodeFilter.Add(new QueryFilter("Code", materialCodeModel.Code));
					materialCodeFilter.Add(new QueryFilter("LocationID", (int)subdepid));
					string materialCode = _materialCodeAppService.Get(materialCodeFilter, true);

					MaterialCodeModel oldMaterialCode = materialCode.DeserializeToMaterialCode();
					oldMaterialCode.ModifiedDate = DateTime.Now;
					oldMaterialCode.ModifiedBy = AccountName;
					oldMaterialCode.LocationID = subdepid;
					oldMaterialCode.Code = materialCodeModel.Code;
					oldMaterialCode.Description = materialCodeModel.Description;
					string data = JsonHelper<MaterialCodeModel>.Serialize(oldMaterialCode);

					if (oldMaterialCode.ID == 0)
					{
						_materialCodeAppService.Add(data);
						break;
					}
					else
					{
						_materialCodeAppService.Update(data);
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
				_materialCodeAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult ExportExcel()
		{
			try
			{
				// Getting all data    			
				string materialCodes = _materialCodeAppService.GetAll();
				List<MaterialCodeModel> materialCodeList = materialCodes.DeserializeToMaterialCodeList();
				//logList = logList.OrderByDescending(x => x.Timestamp).ToList();

				//foreach (var item in logList)
				//{
				//	item.Message = item.Message.Length > 200 ? item.Message.Substring(0, 200) : item.Message;
				//}

				//byte[] excelData = ExcelGenerator.ExportMasterUserMaterialCode(logList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-UserMaterialCode.xlsx");
				//Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
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
			string materialCode = _materialCodeAppService.GetById(id, true);
			MaterialCodeModel model = materialCode.DeserializeToMaterialCode();

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

			return PartialView(model);
		}

		[HttpPost]
		public ActionResult Edit(MaterialCodeModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				if (model.LocationID == 0)
				{
					SetFalseTempData(UIResources.SubDepartmentIsMissing);
					return RedirectToAction("Index");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<MaterialCodeModel>.Serialize(model);
				_materialCodeAppService.Update(data);

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
				string materialCodes = _materialCodeAppService.GetAll();
				List<MaterialCodeModel> materialCodeList = materialCodes.DeserializeToMaterialCodeList();

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");

				materialCodeList = materialCodeList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

				int recordsTotal = materialCodeList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if ((!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn)) && !string.IsNullOrEmpty(sortColumnDir))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in materialCodeList)
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
					materialCodeList = materialCodeList.Where(m => m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
												 m.Description.ToLower().Contains(searchValue.ToLower()) ||
												 m.Location != null && m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								materialCodeList = materialCodeList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								materialCodeList = materialCodeList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								materialCodeList = materialCodeList.OrderBy(x => x.Location).ToList();
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
								materialCodeList = materialCodeList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								materialCodeList = materialCodeList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								materialCodeList = materialCodeList.OrderByDescending(x => x.Location).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = materialCodeList.Count();

				// Paging     
				var data = materialCodeList.Skip(skip).Take(pageSize).ToList();

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

				return Json(new { data = new List<MaterialCodeModel>() }, JsonRequestBehavior.AllowGet);
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
				string materialCodes = _materialCodeAppService.GetAll();
				List<MaterialCodeModel> materialCodeList = materialCodes.DeserializeToMaterialCodeList();

				int recordsTotal = materialCodeList.Count();
				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in materialCodeList)
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
					materialCodeList = materialCodeList.Where(m => m.Code.ToString().ToLower().Contains(searchValue.ToLower()) ||
												 m.Description.ToLower().Contains(searchValue.ToLower()) ||
												 m.Location != null && m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "code":
								materialCodeList = materialCodeList.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								materialCodeList = materialCodeList.OrderBy(x => x.Description).ToList();
								break;
							case "location":
								materialCodeList = materialCodeList.OrderBy(x => x.Location).ToList();
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
								materialCodeList = materialCodeList.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								materialCodeList = materialCodeList.OrderByDescending(x => x.Description).ToList();
								break;
							case "location":
								materialCodeList = materialCodeList.OrderByDescending(x => x.Location).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = materialCodeList.Count();

				// Paging     
				var data = materialCodeList.Skip(skip).Take(pageSize).ToList();

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

				return Json(new { data = new List<MaterialCodeModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
