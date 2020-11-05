using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	[CustomAuthorize("location")]
	public class LocationController : BaseController<LocationModel>
	{
		#region ::Init::
		private readonly ILocationAppService _locationAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		#endregion

		#region ::Constructor::
		public LocationController(
			ILocationAppService locationAppService,
			ILoggerAppService logger,
			IMenuAppService menuService,
			IReferenceDetailAppService referenceDetailAppService,
			IReferenceAppService referenceAppService)
		{
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_logger = logger;
			_menuService = menuService;
			_referenceDetailAppService = referenceDetailAppService;
		}
		#endregion

		#region ::Public Methods::
		[HttpPost]
		public ActionResult GetProductionCenterByCountry(string code)
		{
			List<SelectListItem> _menuList = GetMenuList(ReferenceEnum.ProdCenter, code);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetDepartmentByProductionCenter(string code)
		{
			List<SelectListItem> _menuList = GetMenuList(ReferenceEnum.Dep, code);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetSubDepartmentByDepartment(string code, string pcCode)
		{
			List<SelectListItem> _menuList = GetSubDepMenuList(ReferenceEnum.SubDep, code, pcCode);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		// GET: Location
		public ActionResult Index()
		{
			GetTempData();

			LocationTreeModel model = GetIndexModel();

			return View(model);
		}

		// GET: Location/Create
		public ActionResult Create()
		{
			return View();
		}

		// POST: Location/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(LocationTreeModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				// map new location
				if (string.IsNullOrEmpty(model.Code))
					return AddExistingLocation(model);
				else
					return CreateNewLocation(model);				
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// GET: Location/Edit/5
		public ActionResult EditSub(int id)
		{
			LocationModel locModel = GetLocation(id);

			string refList = _referenceAppService.GetDetailAll(ReferenceEnum.SubDep, true);
			List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
			ReferenceDetailModel model = refModelList.Where(x => x.Code == locModel.Code).FirstOrDefault();
			locModel.Description = model.Description;
			locModel.Type = 3;

			return PartialView("Edit", locModel);
		}

		public ActionResult EditDep(int id)
		{
			LocationModel locModel = GetLocation(id);

			string refList = _referenceAppService.GetDetailAll(ReferenceEnum.Dep, true);
			List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
			ReferenceDetailModel model = refModelList.Where(x => x.Code == locModel.Code).FirstOrDefault();
			locModel.Description = model.Description;
			locModel.Type = 2;

			return PartialView("Edit", locModel);
		}

		public ActionResult EditPC(int id)
		{
			LocationModel locModel = GetLocation(id);

			string refList = _referenceAppService.GetDetailAll(ReferenceEnum.ProdCenter, true);
			List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
			ReferenceDetailModel model = refModelList.Where(x => x.Code == locModel.Code).FirstOrDefault();
			locModel.Description = model.Description;
			locModel.Type = 1;

			return PartialView("Edit", locModel);
		}

		// POST: Location/Edit/5
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Edit(LocationModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				LocationModel oldEntity = GetLocation(model.ID);

				string referenceID = string.Empty;
				if (model.Type == 1)
					referenceID = ((int)ReferenceEnum.ProdCenter).ToString();
				else if (model.Type == 2)
					referenceID = ((int)ReferenceEnum.Dep).ToString();
				else if (model.Type == 3)
					referenceID = ((int)ReferenceEnum.SubDep).ToString();

				string refList = _referenceDetailAppService.FindByNoTracking("ReferenceID", referenceID, true);
				List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();

				ReferenceDetailModel refmodel = refModelList.Where(x => x.Code == oldEntity.Code).FirstOrDefault();
				refmodel.Code = model.Code;
				refmodel.Description = model.Description;
				refmodel.ModifiedBy = AccountName;
				refmodel.ModifiedDate = DateTime.Now;

				string refData = JsonHelper<ReferenceDetailModel>.Serialize(refmodel);

				_referenceAppService.UpdateDetail(refData);

				if (oldEntity.Code != model.Code)
				{
					string locations = _locationAppService.FindByNoTracking("Code", oldEntity.Code);
					List<LocationModel> locationList = locations.DeserializeToLocationList();

					foreach (var item in locationList)
					{
						item.Code = model.Code;
						item.ModifiedBy = AccountName;
						item.ModifiedDate = DateTime.Now;

						string data = JsonHelper<LocationModel>.Serialize(item);

						_locationAppService.Update(data);
					}
				}

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
				LocationModel location = GetLocation(id);
				location.IsDeleted = true;
				string locationData = JsonHelper<LocationModel>.Serialize(location);
				_locationAppService.Update(locationData);

				string children = _locationAppService.FindByNoTracking("ParentID", location.ID.ToString(), true);
				List<LocationModel> childrenList = children.DeserializeToLocationList();

				foreach (var child in childrenList)
				{
					child.IsDeleted = true;
					locationData = JsonHelper<LocationModel>.Serialize(child);

					_locationAppService.Update(locationData);
				}

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
				LocationTreeModel model = new LocationTreeModel();
				// Getting all data    			
				string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
				List<LocationModel> pcList = pcs.DeserializeToLocationList();

				int index = 1;
				foreach (var item in pcList)
				{
					string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
					ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
					if (!string.IsNullOrEmpty(pc))
						model.ProductionCenters.Add(pcModel);
				}

				// get department list
				foreach (var pc in model.ProductionCenters)
				{
					LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
					string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
					List<LocationModel> departmentList = departments.DeserializeToLocationList();

					foreach (var d in departmentList)
					{
						string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
						DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

						string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
						List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

						foreach (var subdeb in subdepartmentList)
						{
							string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
							if (!string.IsNullOrEmpty(sds))
								deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
						}

						pc.Departments.Add(deptModel);
					}
				}

				byte[] excelData = ExcelGenerator.ExportMasterLocation(model, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Locations.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
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

				// Get the location list
				string locationList = _locationAppService.GetAll(true);
				List<LocationModel> locationModelList = string.IsNullOrEmpty(locationList) ? null : JsonConvert.DeserializeObject<List<LocationModel>>(locationList);
				if (locationModelList == null)
					return Json(new { data = new List<LocationModel>(), recordsFiltered = 0, recordsTotal = 0, draw }, JsonRequestBehavior.AllowGet);

				// Get prod center list			
				foreach (var location in locationModelList)
				{
					// get the parent if any
					string parent = _referenceAppService.GetDetailBy("Code", location.LocationParentCode, true);
					ReferenceDetailModel parentModel = string.IsNullOrEmpty(parent) ? new ReferenceDetailModel() : JsonConvert.DeserializeObject<ReferenceDetailModel>(parent);
					location.LocationParent = parentModel.Code + " " + parentModel.Description;

					// get the location type
					string locationType = _referenceAppService.GetDetailBy("Code", location.LocationCode, true);
					ReferenceDetailModel locationTypeModel = string.IsNullOrEmpty(locationType) ? new ReferenceDetailModel() : JsonConvert.DeserializeObject<ReferenceDetailModel>(locationType);
					location.Location = locationTypeModel.Code + " " + locationTypeModel.Description;

					// get the country
					string country = _referenceAppService.GetDetailBy("Code", location.CountryCode, true);
					ReferenceDetailModel countryModel = string.IsNullOrEmpty(country) ? new ReferenceDetailModel() : JsonConvert.DeserializeObject<ReferenceDetailModel>(country);
					location.Country = countryModel.Code + " " + countryModel.Description;
				}

				int recordsTotal = locationModelList.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					locationModelList = locationModelList.Where(m => m.LocationParent.ToLower().Contains(searchValue.ToLower()) ||
												m.Location.ToLower().Contains(searchValue.ToLower()) ||
												m.Country.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "id":
								locationModelList = locationModelList.OrderBy(x => x.ID).ToList();
								break;
							case "locationparent":
								locationModelList = locationModelList.OrderBy(x => x.LocationParent).ToList();
								break;
							case "location":
								locationModelList = locationModelList.OrderBy(x => x.Location).ToList();
								break;
							case "country":
								locationModelList = locationModelList.OrderBy(x => x.Country).ToList();
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
								locationModelList = locationModelList.OrderByDescending(x => x.ID).ToList();
								break;
							case "locationparent":
								locationModelList = locationModelList.OrderByDescending(x => x.LocationParent).ToList();
								break;
							case "location":
								locationModelList = locationModelList.OrderByDescending(x => x.Location).ToList();
								break;
							case "country":
								locationModelList = locationModelList.OrderByDescending(x => x.Country).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = locationModelList.Count();

				// Paging     
				var data = locationModelList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<LocationModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
		#endregion

		#region ::Private Methods::
		private LocationTreeModel GetIndexModel()
		{
			LocationTreeModel model = new LocationTreeModel();

			try
			{
				ViewBag.ProductionCenterList = GetProductionCenter();
				ViewBag.DepartmentList = BuildEmptyList();
				ViewBag.SubDeparmentList = BuildEmptyList();

				model.Access = GetAccess(WebConstants.MenuSlug.LOCATION, _menuService);

				int index = 1;

				// get production center list
				string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
				List<LocationModel> pcList = pcs.DeserializeToLocationList();

				foreach (var item in pcList)
				{
					List<QueryFilter> refFilter = new List<QueryFilter>();
					refFilter.Add(new QueryFilter("Code", item.Code));
					refFilter.Add(new QueryFilter("ReferenceID", (int)ReferenceEnum.ProdCenter));

					string pc = _referenceAppService.GetDetail(refFilter);
					ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
					if (!string.IsNullOrEmpty(pc))
						model.ProductionCenters.Add(pcModel);
				}

				// get department list
				foreach (var pc in model.ProductionCenters)
				{
					LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
					string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
					List<LocationModel> departmentList = departments.DeserializeToLocationList();

					foreach (var d in departmentList)
					{
						List<QueryFilter> refFilter = new List<QueryFilter>();
						refFilter.Add(new QueryFilter("Code", d.Code));
						refFilter.Add(new QueryFilter("ReferenceID", (int)ReferenceEnum.Dep));
						string depts = _referenceAppService.GetDetail(refFilter);
						DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

						string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
						List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

						foreach (var subdeb in subdepartmentList)
						{
							List<QueryFilter> refDFilter = new List<QueryFilter>();
							refDFilter.Add(new QueryFilter("Code", subdeb.Code));
							refDFilter.Add(new QueryFilter("ReferenceID", (int)ReferenceEnum.SubDep));
							string sds = _referenceAppService.GetDetail(refDFilter);
							if (!string.IsNullOrEmpty(sds))
								deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
						}

						pc.Departments.Add(deptModel);
					}
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return model;
		}

		private List<SelectListItem> GetProductionCenter()
		{
			List<SelectListItem> result = new List<SelectListItem>();
			result.Add(new SelectListItem { Text = "- Create New -", Value = "0" });
			result.AddRange(GetMenuList(ReferenceEnum.ProdCenter, "ID"));
			return result;
		}

		private LocationModel GetLocation(long locationID)
		{
			string location = _locationAppService.GetById(locationID, true);
			LocationModel locationModel = location.DeserializeToLocation();

			return locationModel;
		}

		private List<SelectListItem> BuildCountryList()
		{
			List<SelectListItem> _menuList = BuildEmptyList();
			_menuList.AddRange(GetMenuList(ReferenceEnum.Country, null));

			return _menuList;
		}

		private List<SelectListItem> GetMenuList(ReferenceEnum enumValue, string code)
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();
			if (code == null || !code.Equals("0"))
			{
				long parId = 0;
				if (!string.IsNullOrEmpty(code))
				{
					string[] codeTemp = code.Split('|');
					if (codeTemp.Length > 1)
					{
						code = codeTemp[0];
						parId = long.Parse(codeTemp[1]);
					}
				}

				string refList = _referenceAppService.GetDetailAll(enumValue, true);
				List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList().OrderBy(x => x.Code).ToList();
				string locList = _locationAppService.FindBy("ParentCode", code, true);
				List<LocationModel> locModelList = locList.DeserializeToLocationList();

				long parentID = locModelList.Count > 0 ? locModelList[0].ParentID : 0;

				if (locModelList.Count == 0)
				{
					string locs = _locationAppService.FindBy("Code", code, true);
					List<LocationModel> locsList = locs.DeserializeToLocationList();
					LocationModel locModel = locsList.Where(x => x.ParentID == parId).FirstOrDefault();
					parentID = locModel == null ? parentID : locModel.ID;
				}

				foreach (var item in refModelList)
				{
					LocationModel loc = locModelList.Where(x => x.ParentID == parentID && x.Code.Equals(item.Code, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

					if (loc != null)
					{
						if (enumValue == ReferenceEnum.Country)
						{
							_menuList.Add(new SelectListItem
							{
								Text = item.Description,
								Value = item.Code + "|" + loc.ID
							});
						}
						else
						{
							_menuList.Add(new SelectListItem
							{
								Text = item.Description,
								Value = item.Code + "|" + parentID
							});
						}
					}
					else
					{
						if (enumValue != ReferenceEnum.ProdCenter)
						{
							_menuList.Add(new SelectListItem
							{
								Text = item.Description,
								Value = item.Code + "#" + parentID
							});
						}
					}
				}
			}

			return _menuList;
		}

		private List<SelectListItem> GetSubDepMenuList(ReferenceEnum enumValue, string code, string pcCode)
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();
			if (code == null || !code.Equals("0"))
			{
				if (!string.IsNullOrEmpty(code))
				{
					string[] codeTemp = code.Split('|');
					if (codeTemp.Length > 1)
					{
						code = codeTemp[0];
					}
				}
				string refList = _referenceAppService.GetDetailAll(enumValue, true);
				List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList().OrderBy(x => x.Code).ToList();

				// check if subdepart exist
				string depList = _locationAppService.FindBy("ParentCode", pcCode.Split('|')[0], true);
				List<LocationModel> depModelList = depList.DeserializeToLocationList();

				LocationModel department = depModelList.Where(x => x.Code == code).FirstOrDefault();

				string subDepList = _locationAppService.FindBy("ParentID", department.ID, true);
				List<LocationModel> subDepModelList = subDepList.DeserializeToLocationList();

				foreach (var item in refModelList)
				{
					if (subDepModelList.Any(x => x.Code.Equals(item.Code, StringComparison.OrdinalIgnoreCase)))
					{
						_menuList.Add(new SelectListItem
						{
							Text = item.Description,
							Value = item.Code + "|" + department.ID
						});
					}
					else
					{
						_menuList.Add(new SelectListItem
						{
							Text = item.Description,
							Value = item.Code + "#" + department.ID
						});
					}
				}
			}

			return _menuList;
		}

		private List<SelectListItem> BuildEmptyList()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();
			_menuList.Add(new SelectListItem
			{
				Text = "- Create New -",
				Value = "0"
			});

			return _menuList;
		}

		private ActionResult AddExistingLocation(LocationTreeModel model)
		{
			try
			{
				long parentID = 0;
				string code = string.Empty;

				if (model.SubDepartmentCode != "0")
				{
					string[] temp = model.SubDepartmentCode.Split('#');
					code = temp[0];
					parentID = long.Parse(temp[1]);
				}
				else if (model.DepartmentCode != "0")
				{
					string[] temp = model.DepartmentCode.Split('#');
					code = temp[0];
					parentID = long.Parse(temp[1]);
				}
				else if (model.ProductionCenterCode != "0")
				{
					string[] temp = model.ProductionCenterCode.Split('#');
					code = temp[0];
					parentID = long.Parse(temp[1]);
				}

				string parent = _locationAppService.GetById(parentID, true);
				string parentCode = parent.DeserializeToLocation().Code;

				LocationModel newLocation = new LocationModel();
				newLocation.ParentID = parentID;
				newLocation.ParentCode = parentCode;
				newLocation.Code = code;
				newLocation.ModifiedBy = AccountName;
				newLocation.ModifiedDate = DateTime.Now;

				string newLoc = JsonHelper<LocationModel>.Serialize(newLocation);
				_locationAppService.Add(newLoc);

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		private ActionResult CreateNewLocation(LocationTreeModel model)
		{
			try
			{
				string refDetailList = _referenceAppService.GetDetailAll(true);
				List<ReferenceDetailModel> refDetailModelList = refDetailList.DeserializeToRefDetailList();
				if (refDetailModelList.Any(x => x.Code.Equals(model.Code, StringComparison.OrdinalIgnoreCase)))
				{
					SetFalseTempData(string.Format(UIResources.DataExist, "Location", model.Code));
					return RedirectToAction("Index");
				}

				if (model.ProductionCenterCode == "0")
				{
					string refList = _referenceAppService.GetDetailAll(ReferenceEnum.ProdCenter, true);
					List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
					if (refModelList.Any(x => x.Code.Equals(model.Code, StringComparison.OrdinalIgnoreCase)))
					{
						SetFalseTempData(string.Format(UIResources.DataExist, "Production Center", model.Code));
						return RedirectToAction("Index");
					}
					else
					{
						ReferenceDetailModel newReference = new ReferenceDetailModel();
						newReference.ReferenceID = (int)ReferenceEnum.ProdCenter;
						newReference.Code = model.Code;
						newReference.Description = model.Description;
						newReference.ModifiedBy = AccountName;
						newReference.ModifiedDate = DateTime.Now;

						string newref = JsonHelper<ReferenceDetailModel>.Serialize(newReference);
						_referenceAppService.AddDetail(newref);

						LocationModel newModel = new LocationModel();
						newModel.ParentID = 1;
						newModel.ParentCode = "ID";
						newModel.Code = model.Code;
						newModel.ModifiedBy = AccountName;
						newModel.ModifiedDate = DateTime.Now;

						string newLocation = JsonHelper<LocationModel>.Serialize(newModel);
						_locationAppService.Add(newLocation);

						SetTrueTempData(UIResources.CreateSucceed);
						return RedirectToAction("Index");
					}
				}
				else if (model.DepartmentCode == "0")
				{
					string refList = _referenceAppService.GetDetailAll(ReferenceEnum.Dep, true);
					List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
					if (refModelList.Any(x => x.Code.Equals(model.Code, StringComparison.OrdinalIgnoreCase)))
					{
						SetFalseTempData(string.Format(UIResources.DataExist, "Department", model.Code));
						return RedirectToAction("Index");
					}
					else
					{
						ReferenceDetailModel newReference = new ReferenceDetailModel();
						newReference.ReferenceID = (int)ReferenceEnum.Dep;
						newReference.Code = model.Code;
						newReference.Description = model.Description;
						newReference.ModifiedBy = AccountName;
						newReference.ModifiedDate = DateTime.Now;

						string newref = JsonHelper<ReferenceDetailModel>.Serialize(newReference);
						_referenceAppService.AddDetail(newref);

						string[] temp = model.ProductionCenterCode.Split('|');
						string parentCode = temp[0];
						long parentCodeParentID = long.Parse(temp[1]);
						string locs = _locationAppService.FindBy("Code", temp[0], true);
						List<LocationModel> locsList = locs.DeserializeToLocationList();
						LocationModel locModel = locsList.Where(x => x.ParentID == parentCodeParentID).FirstOrDefault();
						long parentID = locModel == null ? 0 : locModel.ID;

						LocationModel newModel = new LocationModel();
						newModel.ParentID = parentID;
						newModel.ParentCode = parentCode;
						newModel.Code = model.Code;
						newModel.ModifiedBy = AccountName;
						newModel.ModifiedDate = DateTime.Now;

						string newLocation = JsonHelper<LocationModel>.Serialize(newModel);
						_locationAppService.Add(newLocation);

						SetTrueTempData(UIResources.CreateSucceed);
						return RedirectToAction("Index");
					}
				}
				else if (model.SubDepartmentCode == "0")
				{
					string refList = _referenceAppService.GetDetailAll(ReferenceEnum.SubDep, true);
					List<ReferenceDetailModel> refModelList = refList.DeserializeToRefDetailList();
					if (refModelList.Any(x => x.Code.Equals(model.Code, StringComparison.OrdinalIgnoreCase)))
					{
						SetFalseTempData(string.Format(UIResources.DataExist, "Sub Department", model.Code));
						return RedirectToAction("Index");
					}
					else
					{
						ReferenceDetailModel newReference = new ReferenceDetailModel();
						newReference.ReferenceID = (int)ReferenceEnum.SubDep;
						newReference.Code = model.Code;
						newReference.Description = model.Description;
						newReference.ModifiedBy = AccountName;
						newReference.ModifiedDate = DateTime.Now;

						string newref = JsonHelper<ReferenceDetailModel>.Serialize(newReference);
						_referenceAppService.AddDetail(newref);

						string[] temp = model.DepartmentCode.Split('|');
						string parentCode = temp[0];
						long parentCodeParentID = long.Parse(temp[1]);
						string locs = _locationAppService.FindBy("Code", temp[0], true);
						List<LocationModel> locsList = locs.DeserializeToLocationList();
						LocationModel locModel = locsList.Where(x => x.ParentID == parentCodeParentID).FirstOrDefault();
						long parentID = locModel == null ? 0 : locModel.ID;

						LocationModel newModel = new LocationModel();
						newModel.ParentID = parentID;
						newModel.ParentCode = parentCode;
						newModel.Code = model.Code;
						newModel.ModifiedBy = AccountName;
						newModel.ModifiedDate = DateTime.Now;

						string newLocation = JsonHelper<LocationModel>.Serialize(newModel);
						_locationAppService.Add(newLocation);

						SetTrueTempData(UIResources.CreateSucceed);
						return RedirectToAction("Index");
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
		#endregion
	}

	public class TempLocation
	{
		public List<SelectListItem> Parents { get; set; }
		public List<SelectListItem> Children { get; set; }
	}
}
