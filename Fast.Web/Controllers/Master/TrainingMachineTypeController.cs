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
	[CustomAuthorize("training")]
	public class TrainingMachineTypeController : BaseController<ReferenceDetailModel>
	{
		private readonly IReferenceAppService _referenceAppService;
        private readonly ITrainingTitleMachineTypeAppService _trainingTitleMachineTypeAppService;
        private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		private readonly ITrainingTitleAppService _trainingTitleAppService;

		public TrainingMachineTypeController(
			IReferenceAppService referenceAppService,
            ITrainingTitleMachineTypeAppService trainingTitleMachineTypeAppService,
            IMenuAppService menuService,
			ITrainingTitleAppService trainingTitleAppService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
            _trainingTitleMachineTypeAppService = trainingTitleMachineTypeAppService;
            _menuService = menuService;
			_logger = logger;
			_trainingTitleAppService = trainingTitleAppService;
		}

		public ActionResult Index()
		{
			GetTempData();

			ViewBag.TrainingTitleList = DropDownHelper.BuildMultiEmpty();
			ViewBag.MachineTypeList = DropDownHelper.BuildMultiEmpty();

			TrainingTitleMachineTypeModel model = new TrainingTitleMachineTypeModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

			return View(model);
		}

		[HttpPost]
		public JsonResult AutoComplete(string prefix)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("Title", prefix, Operator.Contains));

			string trainingTitles = _trainingTitleAppService.Find(filters);
			List<TrainingTitleModel> trainingTitleModelList = trainingTitles.DeserializeToTrainingTitleList();

			trainingTitleModelList = trainingTitleModelList.OrderBy(x => x.Title).ToList();

			return Json(trainingTitleModelList, JsonRequestBehavior.AllowGet);
		}

		public ActionResult Create()
		{
			ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService);
			ViewBag.TrainingTitleList = DropDownHelper.BindDropDownTrainingTitle(_trainingTitleAppService);

			TrainingTitleMachineTypeModel model = new TrainingTitleMachineTypeModel();
			model.Access = GetAccess(WebConstants.MenuSlug.TRAINING, _menuService);

			return PartialView(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(TrainingTitleMachineTypeModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (model.TrainingTitleID == 0)
				{
					SetFalseTempData("No training title selected");
					return RedirectToAction("Index");
				}

				if (model.MachineTypeIDs == null || model.MachineTypeIDs.Count() == 0)
				{
					SetFalseTempData("Please select machine type first");
					return RedirectToAction("Index");
				}

				foreach (var item in model.MachineTypeIDs)
				{
					ICollection<QueryFilter> filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("TrainingTitleID", model.TrainingTitleID.ToString()));
					filters.Add(new QueryFilter("MachineTypeID", item.ToString()));
					filters.Add(new QueryFilter("IsDeleted", "0"));

					string exist = _trainingTitleMachineTypeAppService.Get(filters, true);
					if (!string.IsNullOrEmpty(exist))
					{
						SetFalseTempData(string.Format(UIResources.DataExist, model.TrainingTitleID, item));
						return RedirectToAction("Index");
					}
				}

				foreach (var machineId in model.MachineTypeIDs)
				{
					TrainingTitleMachineTypeModel newEntity = new TrainingTitleMachineTypeModel();
					newEntity.TrainingTitleID = model.TrainingTitleID;
					newEntity.MachineTypeID = machineId;
					newEntity.ModifiedBy = AccountName;
					newEntity.ModifiedDate = DateTime.Now;

					string data = JsonHelper<TrainingTitleMachineTypeModel>.Serialize(newEntity);

					_trainingTitleMachineTypeAppService.Add(data);
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
				string refMachineTypeList = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
				List<ReferenceDetailModel> refMachineTypeModelList = refMachineTypeList.DeserializeToRefDetailList();

				string trainingTitleList = _trainingTitleAppService.GetAll(true);
				List<TrainingTitleModel> trainingTitleModelList = trainingTitleList.DeserializeToTrainingTitleList();

				// Getting all data    			
				string trainingTitleMachineTypeList = _trainingTitleMachineTypeAppService.GetAll(true);
				List<TrainingTitleMachineTypeModel> trainingTitleMTModelList = trainingTitleMachineTypeList.DeserializeToTrainingTitleMachineTypeList();

				List<TrainingTitleMachineTypeModel> result = new List<TrainingTitleMachineTypeModel>();

				foreach (var item in trainingTitleMTModelList)
				{
					var exist = result.Where(x => x.TrainingTitleID == item.TrainingTitleID).FirstOrDefault();
					if (exist == null)
					{
						var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						item.MachineTypeList = mt == null ? string.Empty : mt.Code;

						var tt = trainingTitleModelList.Where(x => x.ID == item.TrainingTitleID).FirstOrDefault();
						item.TrainingTitle = tt == null ? string.Empty : tt.Title;

						result.Add(item);
					}
					else
					{
						var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						exist.MachineTypeList = exist.MachineTypeList + ", " + mt.Code;
					}
				}

				byte[] excelData = ExcelGenerator.ExportTrainingTitleMachineType(result, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-TrainingTitleMachineType.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Edit(int id)
		{
			TrainingTitleMachineTypeModel model = GetTrainingTitleMachineType(id);

			string machines = _trainingTitleMachineTypeAppService.FindByNoTracking("TrainingTitleID", model.TrainingTitleID.ToString(), true);
			List<TrainingTitleMachineTypeModel> machineList = machines.DeserializeToTrainingTitleMachineTypeList();
			List<long> machineIDList = machineList.Select(c => c.MachineTypeID).Distinct().ToList();

			ViewBag.MachineTypeList = DropDownHelper.BindDropDownMultiMachineType(_referenceAppService, machineIDList);
			ViewBag.TrainingTitleList = DropDownHelper.BindDropDownTrainingTitle(_trainingTitleAppService);

			return PartialView(model);
		}

		[HttpPost]
		public ActionResult Edit(TrainingTitleMachineTypeModel model)
		{
			try
			{
				model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (model.TrainingTitleID == 0)
				{
					SetFalseTempData("No training title selected");
					return RedirectToAction("Index");
				}

				if (model.MachineTypeIDs == null || model.MachineTypeIDs.Count() == 0)
				{
					SetFalseTempData("Please select machine type first");
					return RedirectToAction("Index");
				}

				string trainings = _trainingTitleMachineTypeAppService.FindByNoTracking("TrainingTitleID", model.TrainingTitleID.ToString(), true);
				List<TrainingTitleMachineTypeModel> trainingList = trainings.DeserializeToTrainingTitleMachineTypeList();

				foreach (var item in trainingList)
				{
					if (!model.MachineTypeIDs.Any(x => x == item.MachineTypeID))
					{
						// remove if not selected						
						_trainingTitleMachineTypeAppService.Remove(item.ID);
					}
				}

				foreach (var item in model.MachineTypeIDs)
				{
					if (!trainingList.Any(x => x.MachineTypeID == item))
					{
						TrainingTitleMachineTypeModel newEntity = new TrainingTitleMachineTypeModel();
						newEntity.TrainingTitleID = model.TrainingTitleID;
						newEntity.MachineTypeID = item;
						newEntity.ModifiedBy = AccountName;
						newEntity.ModifiedDate = DateTime.Now;

						string data = JsonHelper<TrainingTitleMachineTypeModel>.Serialize(newEntity);

						_trainingTitleMachineTypeAppService.Add(data);
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

		public ActionResult Delete(int id)
		{
			return View();
		}

		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				TrainingTitleMachineTypeModel model = GetTrainingTitleMachineType(id);
				model.IsDeleted = true;

				string userData = JsonHelper<TrainingTitleMachineTypeModel>.Serialize(model);
				_trainingTitleMachineTypeAppService.Update(userData);

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

				string refMachineTypeList = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.MachineType).ToString(), true);
				List<ReferenceDetailModel> refMachineTypeModelList = refMachineTypeList.DeserializeToRefDetailList();

				string trainingTitleList = _trainingTitleAppService.GetAll(true);
				List<TrainingTitleModel> trainingTitleModelList = trainingTitleList.DeserializeToTrainingTitleList();

				// Getting all data    			
				string trainingTitleMTList = _trainingTitleMachineTypeAppService.GetAll(true);
				List<TrainingTitleMachineTypeModel> trainingTitleMTModelList = trainingTitleMTList.DeserializeToTrainingTitleMachineTypeList();


				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				List<TrainingTitleMachineTypeModel> result = new List<TrainingTitleMachineTypeModel>();

				foreach (var item in trainingTitleMTModelList)
				{
					var exist = result.Where(x => x.TrainingTitleID == item.TrainingTitleID).FirstOrDefault();
					if (exist == null)
					{
						var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						item.MachineTypeList = mt == null ? string.Empty : mt.Code;

						var tt = trainingTitleModelList.Where(x => x.ID == item.TrainingTitleID).FirstOrDefault();
						item.TrainingTitle = tt == null ? string.Empty : tt.Title;

						result.Add(item);
					}
					else
					{
						var mt = refMachineTypeModelList.Where(x => x.ID == item.MachineTypeID).FirstOrDefault();
						exist.MachineTypeList = exist.MachineTypeList + ", " + mt.Code;
					}
				}

				int recordsTotal = result.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					result = result.Where(m => m.TrainingTitle.ToLower().Contains(searchValue.ToLower()) ||
											   m.MachineTypeList.ToLower().Contains(searchValue.ToLower())).ToList();

				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "trainingtitle":
								result = result.OrderBy(x => x.TrainingTitle).ToList();
								break;
							case "machinetypelist":
								result = result.OrderBy(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "trainingtitle":
								result = result.OrderByDescending(x => x.TrainingTitle).ToList();
								break;
							case "machinetypelist":
								result = result.OrderByDescending(x => x.MachineType).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = result.Count();

				// Paging     
				var data = result.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<TrainingTitleMachineTypeModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private TrainingTitleMachineTypeModel GetTrainingTitleMachineType(long trainingTitleMachineTypeID)
		{
			string trainingTitleMachineType = _trainingTitleMachineTypeAppService.GetById(trainingTitleMachineTypeID, true);
			TrainingTitleMachineTypeModel model = trainingTitleMachineType.DeserializeToTrainingTitleMachineType();

			return model;
		}
	}
}
