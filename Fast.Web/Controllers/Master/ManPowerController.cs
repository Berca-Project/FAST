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
	[CustomAuthorize("mpp")]
	public class ManPowerController : BaseController<ManPowerModel>
	{
		private readonly IManPowerAppService _manPowerAppService;
		private readonly IJobTitleAppService _jobTitleAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
		private readonly IMenuAppService _menuService;
		private readonly ILoggerAppService _logger;
        private readonly IRoleAppService _roleAppService;

        public ManPowerController(
			IJobTitleAppService jobTitleAppService,
			ILoggerAppService logger,
            IRoleAppService roleAppService,
            ILocationAppService locationAppService,
            IMenuAppService menuService,
			IManPowerAppService manPowerAppService,
			IReferenceAppService referenceAppService)
		{
			_manPowerAppService = manPowerAppService;
			_referenceAppService = referenceAppService;
			_jobTitleAppService = jobTitleAppService;
			_logger = logger;
			_menuService = menuService;
            _roleAppService = roleAppService;
            _locationAppService = locationAppService;
		}

		public ActionResult Index()
		{
			GetTempData();

			ManPowerModel model = new ManPowerModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MAN_POWER, _menuService);

            ViewBag.CountryList = DropDownHelper.BuildEmptyList();
            ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();

            return View(model);
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

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetSubDepartmentByDepartmentID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, id);
            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				_manPowerAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Create()
		{			
			//ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitle(_jobTitleAppService);
            ViewBag.RoleName = DropDownHelper.BindDropDownRole(_roleAppService);
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();

            return PartialView();
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(ManPowerModel mpModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

                if (mpModel.LocationID == 0)
                {
                    SetFalseTempData("No selected department");
                    return RedirectToAction("Index");
                }

                List<QueryFilter> mpFilter = new List<QueryFilter>();
				mpFilter.Add(new QueryFilter("RoleName", mpModel.RoleName.ToString()));
				mpFilter.Add(new QueryFilter("LocationID", mpModel.LocationID.ToString()));

				string exist = _manPowerAppService.Get(mpFilter);
				if (!string.IsNullOrEmpty(exist))
				{
					SetFalseTempData(string.Format(UIResources.DataExist, mpModel.RoleName, mpModel.Location));
					return RedirectToAction("Index");
				}

				mpModel.ModifiedBy = AccountName;
				mpModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<ManPowerModel>.Serialize(mpModel);

				_manPowerAppService.Add(data);

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Edit(int id)
		{
            //ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitle(_jobTitleAppService);
            ViewBag.RoleName = DropDownHelper.BindDropDownRole(_roleAppService);
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();

            string mp = _manPowerAppService.GetById(id);
			ManPowerModel model = mp.DeserializeToManPower();

            long countryID = 0;
            long pcID = 0;
            long depID = 0;
            long subDepID = 0;
            string completeLocation = DropDownHelper.ExtractLocation(_locationAppService, model.LocationID, out countryID, out pcID, out depID, out subDepID);
            model.CountryID = countryID;
            model.ProdCenterID = pcID;
            model.DepartmentID = depID;
            model.LocationID = depID;

            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);

            if (model.CountryID == 0)
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            else
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterByCountryID(_locationAppService, _referenceAppService, model.CountryID);

            if (model.ProdCenterID == 0)
                ViewBag.DepartmentList = DropDownHelper.BuildEmptyListDepartment();
            else
                ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, model.ProdCenterID);

            if (model.DepartmentID == 0)
                ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyListSubDepartment();
            else
                ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, model.DepartmentID);

            return PartialView(model);
		}

		[HttpPost]
		public ActionResult Edit(ManPowerModel model)
		{
			try
			{
				model.Access = GetAccess(WebConstants.MenuSlug.MAN_POWER, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

                if (model.LocationID == 0)
                {
                    SetFalseTempData("No selected department");
                    return RedirectToAction("Index");
                }

                string mp = _manPowerAppService.GetById(model.ID, true);
				ManPowerModel mpModel = mp.DeserializeToManPower();
				mpModel.ModifiedBy = AccountName;
				mpModel.ModifiedDate = DateTime.Now;
				mpModel.LocationID = model.LocationID;
				mpModel.Value = model.Value;
                //mpModel.JobTitle = model.JobTitle;
                mpModel.RoleName = model.RoleName;

                string data = JsonHelper<ManPowerModel>.Serialize(mpModel);

				_manPowerAppService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string mp = _manPowerAppService.GetAll();
				List<ManPowerModel> mpList = mp.DeserializeToManPowerList();
				Dictionary<long, string> jobTitleList = new Dictionary<long, string>();				

				foreach (var item in mpList)
				{
					if (jobTitleList.ContainsKey(item.JobTitleID))
					{
						string jt = "";
						jobTitleList.TryGetValue(item.JobTitleID, out jt);
						item.JobTitle = jt;
					}
					else
					{
						string jt = _jobTitleAppService.GetById(item.JobTitleID);
						JobTitleModel jtModel = jt.DeserializeToJobTitle();
						item.JobTitle = jtModel.Title;
						jobTitleList.Add(item.JobTitleID, jtModel.Title);
					}

                    item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
				}

				byte[] excelData = ExcelGenerator.ExportManPower(mpList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Man-Power.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.GenerateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID);
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

				// Getting all data    			
				string manpowers = _manPowerAppService.GetAll();
				List<ManPowerModel> manPowerModelList = manpowers.DeserializeToManPowerList();

				foreach (var item in manPowerModelList)
				{
                    item.Location = _locationAppService.GetLocationFullCode(item.LocationID);

					//string jt = _jobTitleAppService.GetById(item.JobTitleID);
					//JobTitleModel jtModel = jt.DeserializeToJobTitle();

					//item.JobTitle = jtModel.Title;
				}

                manPowerModelList = manPowerModelList.Where(x => !string.IsNullOrEmpty(x.RoleName)).ToList();
                int recordsTotal = manPowerModelList.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					manPowerModelList = manPowerModelList.Where(m => !string.IsNullOrEmpty(m.RoleName) && m.RoleName.ToLower().Contains(searchValue.ToLower()) ||
                                                                     m.Value.ToString().Contains(searchValue.ToLower()) ||
                                                                     m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "jobtitle":
								manPowerModelList = manPowerModelList.OrderBy(x => x.JobTitle).ToList();
								break;
							case "location":
								manPowerModelList = manPowerModelList.OrderBy(x => x.Location).ToList();
								break;
							case "value":
								manPowerModelList = manPowerModelList.OrderBy(x => x.Value).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "jobtitle":
								manPowerModelList = manPowerModelList.OrderByDescending(x => x.JobTitle).ToList();
								break;
							case "location":
								manPowerModelList = manPowerModelList.OrderByDescending(x => x.Location).ToList();
								break;
							case "value":
								manPowerModelList = manPowerModelList.OrderByDescending(x => x.Value).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = manPowerModelList.Count();

				// Paging     
				var data = manPowerModelList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<ManPowerModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
