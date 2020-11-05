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
	[CustomAuthorize("shuttlerequest")]
	public class ShuttleRequestController : BaseController<ShuttleRequestModel>
	{
		#region ::Init::
		private readonly IShuttleRequestAppService _shuttleRequestAppService;
		private readonly IMenuAppService _menuService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IMppAppService _mppAppService;
		#endregion

		#region ::Constructor::
		public ShuttleRequestController(
			ILoggerAppService logger,
			IShuttleRequestAppService shuttleRequestAppService,
			IReferenceAppService referenceAppService,
			IEmployeeAppService employeeAppService,
			ILocationAppService locationAppService,
			IMppAppService mppAppService,
			IMenuAppService menuService)
		{
			_mppAppService = mppAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_referenceAppService = referenceAppService;
			_shuttleRequestAppService = shuttleRequestAppService;
			_menuService = menuService;
			_logger = logger;
		}
		#endregion

		#region ::Public Methods::		
		public ActionResult Index()
		{
			GetTempData();

			ShuttleRequestModel model = new ShuttleRequestModel();
			model.Access = GetAccess(WebConstants.MenuSlug.SHUTTLE_REQUEST, _menuService);

			return View(model);
		}

		public ActionResult DetailReport()
		{
			GetTempData();

			ShuttleRequestModel model = new ShuttleRequestModel();
			model.Access = GetAccess(WebConstants.MenuSlug.SHUTTLE_REQUEST, _menuService);

			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

			return View(model);
		}

		public ActionResult SummaryReport()
		{
			GetTempData();

			ShuttleRequestModel model = new ShuttleRequestModel();
			model.Access = GetAccess(WebConstants.MenuSlug.SHUTTLE_REQUEST, _menuService);

			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

			return View(model);
		}

		[HttpPost]
		public JsonResult AutoComplete(string prefix)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			if (prefix.All(Char.IsDigit))
				filters.Add(new QueryFilter("EmployeeID", prefix, Operator.StartsWith));
			else
				filters.Add(new QueryFilter("FullName", prefix, Operator.Contains));

			string emplist = _employeeAppService.Find(filters);
			List<EmployeeModel> empModelList = emplist.DeserializeToEmployeeList();

			if (prefix.All(Char.IsDigit))
			{
				empModelList = empModelList.OrderBy(x => x.EmployeeID).ToList();
			}
			else
			{
				empModelList = empModelList.OrderBy(x => x.FullName).ToList();
			}

			return Json(empModelList, JsonRequestBehavior.AllowGet);
		}

		public ActionResult Create()
		{
			ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService);
            ViewBag.HalteList = DropDownHelper.BindDropDownHalte(_referenceAppService);
            ViewBag.CostCenterList = DropDownHelper.BindDropDownCostCenter(_referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, false);

			EmployeeModel emp = _employeeAppService.GetModelByEmpId(AccountEmployeeID);

			ShuttleRequestModel model = new ShuttleRequestModel();
			model.EmployeeID = AccountEmployeeID;
			model.EmployeeFullname = emp.FullName;
			model.StartDate = DateTime.Now;
			model.EndDate = DateTime.Now;

			return PartialView(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(ShuttleRequestModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				model.EmployeeID = string.IsNullOrEmpty(model.EmployeeID) ? AccountEmployeeID : model.EmployeeID;
				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
				if (model.Department != "0")
				{
					var fd = departList.Where(x => x.Value == model.Department).FirstOrDefault();
					model.Department = fd == null ? "" : fd.Value;
				}

				List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				if (model.ProductionCenterID != "0")
				{
					var pc = pcList.Where(x => x.Value == model.ProductionCenterID).FirstOrDefault();
					model.ProductionCenter = pc == null ? "" : pc.Text;
				}

				List<ShuttleRequestModel> requestList = new List<ShuttleRequestModel>();
				for (var day = model.StartDate.Date; day.Date <= model.EndDate.Date; day = day.AddDays(1))
				{
					ShuttleRequestModel newModel = new ShuttleRequestModel
					{
						StartDate = model.StartDate,
						EndDate = model.EndDate,
						Date = day,
						CostCenter = model.CostCenter,
						Department = model.Department,
						EmployeeFullname = model.EmployeeFullname,
						EmployeeID = model.EmployeeID,
						GuestType = model.GuestType,
						ModifiedBy = model.ModifiedBy,
						ModifiedDate = model.ModifiedDate,
						Phone = model.Phone,
						ProductionCenter = model.ProductionCenter,
						ProductionCenterID = model.ProductionCenterID,
						Purpose = model.Purpose,
						RequestType = model.RequestType,
						LocationFrom = model.LocationFrom,
						LocationTo = model.LocationTo,
						Qty = model.Qty,
						Time = model.Time,
						TotalPassengers = model.TotalPassengers
					};

					requestList.Add(newModel);
				}

				string data = JsonHelper<ShuttleRequestModel>.Serialize(requestList);
				_shuttleRequestAppService.AddRange(data);

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
			string shuttleRequest = _shuttleRequestAppService.GetById(id);
			ShuttleRequestModel shuttleRequestModel = shuttleRequest.DeserializeToShuttleRequest();

			string emp = _employeeAppService.GetBy("EmployeeID", shuttleRequestModel.EmployeeID);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			shuttleRequestModel.EmployeeFullname = empModel.FullName;

			List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			if (!string.IsNullOrEmpty(shuttleRequestModel.ProductionCenter))
			{
				var pc = pcList.Where(x => x.Text == shuttleRequestModel.ProductionCenter).FirstOrDefault();
				shuttleRequestModel.ProductionCenterID = pc == null ? "" : pc.Value;
			}

			ViewBag.CanteenList = DropDownHelper.BindDropDownCanteen(_referenceAppService);
			ViewBag.CostCenterList = DropDownHelper.BindDropDownCostCenter(_referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, false);

			return PartialView(shuttleRequestModel);
		}

		[HttpPost]
		public ActionResult Edit(ShuttleRequestModel model)
		{
			try
			{
				model.Access = GetAccess(WebConstants.MenuSlug.TRAINING, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				model.EmployeeID = string.IsNullOrEmpty(model.EmployeeID) ? AccountEmployeeID : model.EmployeeID;
				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;
				List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);
				if (model.Department != "0")
				{
					var fd = departList.Where(x => x.Value == model.Department).FirstOrDefault();
					model.Department = fd == null ? "" : fd.Value;
				}
				List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				if (model.ProductionCenterID != "0")
				{
					var pc = pcList.Where(x => x.Value == model.ProductionCenterID).FirstOrDefault();
					model.ProductionCenter = pc == null ? "" : pc.Text;
				}

				string data = JsonHelper<ShuttleRequestModel>.Serialize(model);

				_shuttleRequestAppService.Update(data);

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
				string tt = _shuttleRequestAppService.GetById(id, true);
				ShuttleRequestModel ttm = tt.DeserializeToShuttleRequest();
				ttm.IsDeleted = true;

				string data = JsonHelper<ShuttleRequestModel>.Serialize(ttm);
				_shuttleRequestAppService.Update(data);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string shuttleRequests = _shuttleRequestAppService.GetAll(true);
				List<ShuttleRequestModel> shuttleModelList = shuttleRequests.DeserializeToShuttleRequestList();

				string emps = _employeeAppService.GetAll();
				List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

				foreach (var item in shuttleModelList)
				{
					var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
					item.EmployeeFullname = emp == null ? "" : emp.FullName;
				}

				byte[] excelData = ExcelGenerator.ExportShuttleRequest(shuttleModelList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Facility-Shuttle-Request.xlsx");
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

		public ActionResult GenerateExcelDetailReport(string startDate, string endDate, string location, string request)
		{
			try
			{
				DateTime startDateTemp = DateTime.Parse(startDate);
				DateTime endDateTemp = DateTime.Parse(endDate);

				List<ShuttleRequestModel> shuttleModelList = new List<ShuttleRequestModel>();

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("IsDeleted", "0"));
				filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
				filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

				// Getting all data    			
				if (request == "Form")
				{
					if (location != "0")
					{
						List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();

					foreach (var item in shuttleModelList)
					{
						var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
						item.EmployeeFullname = emp == null ? "" : emp.FullName;
					}
				}
				else if (request == "MPP")
				{
					List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}

					shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ToList();
				}
				else
				{
					List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();

					foreach (var item in shuttleModelList)
					{
						var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
						item.EmployeeFullname = emp == null ? "" : emp.FullName;
					}

					filters.Add(new QueryFilter("IsDeleted", "0"));
					filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
					filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}

					shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ToList();
				}

				byte[] excelData = ExcelGenerator.ExportShuttleRequestDetailReport(shuttleModelList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Facility-Shuttle-Request-Detail-Report.xlsx");
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

		public ActionResult GenerateExcelSummaryReport(string startDate, string endDate, string location, string request)
		{
			try
			{
				DateTime startDateTemp = DateTime.Parse(startDate);
				DateTime endDateTemp = DateTime.Parse(endDate);

				List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
				List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

				List<ShuttleRequestModel> shuttleModelList = new List<ShuttleRequestModel>();

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("IsDeleted", "0"));
				filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
				filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

				// Getting all data    			
				if (request == "Form")
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();
				}
				else if (request == "MPP")
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}
				}
				else
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					filters.Add(new QueryFilter("IsDeleted", "0"));
					filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
					filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}
				}

				shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ThenBy(x => x.LocationFrom).ThenBy(x => x.LocationTo).ToList();
				List<ShuttleRequestModel> result = new List<ShuttleRequestModel>();
				foreach (var item in shuttleModelList)
				{
					var temp = result.Where(x => x.Date == item.Date && x.Time == item.Time && x.LocationFrom == item.LocationFrom && x.LocationTo == item.LocationTo).FirstOrDefault();
					if (temp == null)
					{
						item.TotalPassengers = 1;
						result.Add(item);
					}
					else
					{
						temp.TotalPassengers++;
					}
				}

				byte[] excelData = ExcelGenerator.ExportShuttleRequestSummaryReport(result, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Facility-Shuttle-Request-Summary-Report.xlsx");
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
				string shuttles = _shuttleRequestAppService.GetAll(true);
				List<ShuttleRequestModel> shuttleModelList = shuttles.DeserializeToShuttleRequestList().OrderBy(x => x.Date).ThenBy(x => x.Time).ToList();

				string emps = _employeeAppService.GetAll();
				List<EmployeeModel> empModelList = emps.DeserializeToEmployeeList();

				List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

				foreach (var item in shuttleModelList)
				{
					var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
					item.EmployeeFullname = emp == null ? "" : emp.FullName;

					if (!string.IsNullOrEmpty(item.Department))
					{
						var fd = departList.Where(x => x.Value == item.Department).FirstOrDefault();
						item.Department = fd == null ? "" : fd.Text;
					}
				}

				int recordsTotal = shuttleModelList.Count();

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					shuttleModelList = shuttleModelList.Where(m => m.LocationFrom != null && m.LocationFrom.ToLower().Contains(searchValue.ToLower()) ||
																   m.LocationTo != null && m.LocationTo.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								shuttleModelList = shuttleModelList.OrderBy(x => x.StartDate).ToList();
								break;
							case "enddate":
								shuttleModelList = shuttleModelList.OrderBy(x => x.EndDate).ToList();
								break;
							case "time":
								shuttleModelList = shuttleModelList.OrderBy(x => x.Time).ToList();
								break;
							case "locationform":
								shuttleModelList = shuttleModelList.OrderBy(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								shuttleModelList = shuttleModelList.OrderBy(x => x.LocationTo).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.StartDate).ToList();
								break;
							case "enddate":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.EndDate).ToList();
								break;
							case "time":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.Time).ToList();
								break;
							case "locationform":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.LocationTo).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = shuttleModelList.Count();

				// Paging     
				var data = shuttleModelList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;
				ViewBag.ErrorMessage = UIResources.LoadDataFailed;

				_logger.LogError(ex.GetAllMessages(), AccountID);

				return Json(new { data = new List<ShuttleRequestModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetAllDetailReport(string startDate, string endDate, string location, string request)
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

				DateTime startDateTemp = DateTime.Parse(startDate);
				DateTime endDateTemp = DateTime.Parse(endDate);

				List<ShuttleRequestModel> shuttleModelList = new List<ShuttleRequestModel>();

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("IsDeleted", "0"));
				filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
				filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

				// Getting all data    			
				if (request == "Form")
				{
					if (location != "0")
					{
						List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();

					foreach (var item in shuttleModelList)
					{
						var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
						item.EmployeeFullname = emp == null ? "" : emp.FullName;
					}
				}
				else if (request == "MPP")
				{
					List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}

					shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ToList();
				}
				else
				{
					List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();

					foreach (var item in shuttleModelList)
					{
						var emp = empModelList.Where(x => x.EmployeeID != null && x.EmployeeID.Trim() == item.EmployeeID.Trim()).FirstOrDefault();
						item.EmployeeFullname = emp == null ? "" : emp.FullName;
					}

					filters.Add(new QueryFilter("IsDeleted", "0"));
					filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
					filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}

					shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ToList();
				}

				int recordsTotal = shuttleModelList.Count();

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					shuttleModelList = shuttleModelList.Where(m => m.LocationFrom != null && m.LocationFrom.ToLower().Contains(searchValue.ToLower()) ||
																   m.LocationTo != null && m.LocationTo.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								shuttleModelList = shuttleModelList.OrderBy(x => x.StartDate).ToList();
								break;
							case "enddate":
								shuttleModelList = shuttleModelList.OrderBy(x => x.EndDate).ToList();
								break;
							case "time":
								shuttleModelList = shuttleModelList.OrderBy(x => x.Time).ToList();
								break;
							case "locationform":
								shuttleModelList = shuttleModelList.OrderBy(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								shuttleModelList = shuttleModelList.OrderBy(x => x.LocationTo).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.StartDate).ToList();
								break;
							case "enddate":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.EndDate).ToList();
								break;
							case "time":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.Time).ToList();
								break;
							case "locationform":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								shuttleModelList = shuttleModelList.OrderByDescending(x => x.LocationTo).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = shuttleModelList.Count();

				// Paging     
				var data = shuttleModelList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;
				ViewBag.ErrorMessage = UIResources.LoadDataFailed;

				_logger.LogError(ex.GetAllMessages(), AccountID);

				return Json(new { data = new List<ShuttleRequestModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetAllSummaryReport(string startDate, string endDate, string location, string request)
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

				DateTime startDateTemp = DateTime.Parse(startDate);
				DateTime endDateTemp = DateTime.Parse(endDate);

				List<EmployeeModel> empModelList = _employeeAppService.GetAll().DeserializeToEmployeeList();
				List<SelectListItem> pcList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				List<SelectListItem> departList = DropDownHelper.BindDropDownFacilityDepartment(_referenceAppService);

				List<ShuttleRequestModel> shuttleModelList = new List<ShuttleRequestModel>();

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("IsDeleted", "0"));
				filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
				filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

				// Getting all data    			
				if (request == "Form")
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();
				}
				else if (request == "MPP")
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}
				}
				else
				{
					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();
						if (pc != null)
						{
							filters.Add(new QueryFilter("ProductionCenter", pc.Text));
						}
					}

					shuttleModelList = _shuttleRequestAppService.Find(filters).DeserializeToShuttleRequestList();

					filters.Add(new QueryFilter("IsDeleted", "0"));
					filters.Add(new QueryFilter("Date", startDateTemp.ToString(), Operator.GreaterThanOrEqual));
					filters.Add(new QueryFilter("Date", endDateTemp.ToString(), Operator.LessThanOrEqualTo));

					if (location != "0")
					{
						var pc = pcList.Where(x => x.Value == location).FirstOrDefault();

						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(location), "productioncenter");
						mppList = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

						List<MppModel> mppListShift1 = mppList.Where(x => x.Shift.Trim() == "1").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

						List<MppModel> mppListShift2 = mppList.Where(x => x.Shift.Trim() == "2").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

						List<MppModel> mppListShift3 = mppList.Where(x => x.Shift.Trim() == "3").ToList();
						GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
					}
					else
					{
						List<MppModel> mppList = _mppAppService.Find(filters).DeserializeToMppList();

						foreach (var pc in pcList)
						{
							if (pc.Value != "0")
							{
								List<long> locIDList = _locationAppService.GetLocIDListByLocType(long.Parse(pc.Value), "productioncenter");
								var mppListTemp = mppList.Where(x => locIDList.Any(y => y == x.LocationID)).ToList();

								List<MppModel> mppListShift1 = mppListTemp.Where(x => x.Shift.Trim() == "1").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift1, "1");

								List<MppModel> mppListShift2 = mppListTemp.Where(x => x.Shift.Trim() == "2").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift2, "2");

								List<MppModel> mppListShift3 = mppListTemp.Where(x => x.Shift.Trim() == "3").ToList();
								GetMppDataList(startDateTemp, endDateTemp, shuttleModelList, pc.Text, mppListShift3, "3");
							}
						}
					}
				}

				shuttleModelList = shuttleModelList.OrderBy(x => x.Date).ThenBy(x => x.Time).ThenBy(x => x.LocationFrom).ThenBy(x => x.LocationTo).ToList();
                var mppShuttles = shuttleModelList.Where(x => x.IsMPP).ToList();
                var formShuttles = shuttleModelList.Where(x => !x.IsMPP).ToList();

                List<ShuttleRequestModel> result = new List<ShuttleRequestModel>();
                foreach (var item in formShuttles)
                {
                    result.Add(item);
                }

				foreach (var item in mppShuttles)
				{
					var temp = result.Where(x => x.Date == item.Date && x.Time == item.Time && x.LocationFrom == item.LocationFrom && x.LocationTo == item.LocationTo).FirstOrDefault();
					if (temp == null)
					{
						item.TotalPassengers = 1;
						result.Add(item);
					}
					else
					{
						temp.TotalPassengers++;
					}
				}

				int recordsTotal = result.Count();

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					result = result.Where(m => m.LocationFrom != null && m.LocationFrom.ToLower().Contains(searchValue.ToLower()) ||
											   m.LocationTo != null && m.LocationTo.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								result = result.OrderBy(x => x.StartDate).ToList();
								break;
							case "enddate":
								result = result.OrderBy(x => x.EndDate).ToList();
								break;
							case "time":
								result = result.OrderBy(x => x.Time).ToList();
								break;
							case "locationform":
								result = result.OrderBy(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								result = result.OrderBy(x => x.LocationTo).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "startdate":
								result = result.OrderByDescending(x => x.StartDate).ToList();
								break;
							case "enddate":
								result = result.OrderByDescending(x => x.EndDate).ToList();
								break;
							case "time":
								result = result.OrderByDescending(x => x.Time).ToList();
								break;
							case "locationform":
								result = result.OrderByDescending(x => x.LocationFrom).ToList();
								break;
							case "locationto":
								result = result.OrderByDescending(x => x.LocationTo).ToList();
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
				ViewBag.ErrorMessage = UIResources.LoadDataFailed;

				_logger.LogError(ex.GetAllMessages(), AccountID);

				return Json(new { data = new List<ShuttleRequestModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
		#endregion

		#region ::Private Methods::
		private void GetMppDataList(DateTime startDateTemp, DateTime endDateTemp, List<ShuttleRequestModel> shuttleModelList, string prodcenter, List<MppModel> mppList, string shift)
		{
			TimeSpan startTime;
			TimeSpan endTime;

			if (shift == "1")
			{
				startTime = TimeSpan.Parse("06:00");
				endTime = TimeSpan.Parse("14:00");
			}
			else if (shift == "2")
			{
				startTime = TimeSpan.Parse("14:00");
				endTime = TimeSpan.Parse("22:00");
			}
			else
			{
				startTime = TimeSpan.Parse("22:00");
				endTime = TimeSpan.Parse("06:00");
			}

			foreach (var item in mppList)
			{
				if (shuttleModelList.Any(x => x.EmployeeID == item.EmployeeID && x.Date == item.Date))
					continue;

				string gate = item.Location.StartsWith("ID-PK") ? "Main Gate Lot 17" : item.Location.StartsWith("ID-PI") ? "Main Gate Lot 41" : "Delta / Halte Lalin";

				ShuttleRequestModel newModel = new ShuttleRequestModel
				{
					EmployeeID = item.EmployeeID,
					EmployeeFullname = item.EmployeeName,
					StartDate = startDateTemp,
					EndDate = endDateTemp,
					Date = item.Date,
					Time = startTime,
					GuestType = "Internal-MPP",
					Purpose = "Daily Shuttle",
					LocationFrom = gate,
					LocationTo = prodcenter,
					ProductionCenter = prodcenter,
					TotalPassengers = 1,
                    IsMPP = true
                };

				ShuttleRequestModel newModel2 = new ShuttleRequestModel
				{
					EmployeeID = item.EmployeeID,
					EmployeeFullname = item.EmployeeName,
					StartDate = startDateTemp,
					EndDate = endDateTemp,
					Date = shift == "3" ? item.Date.AddDays(1) : item.Date,
					Time = endTime,
					GuestType = "Internal",
					Purpose = "Daily Shuttle",
					LocationFrom = prodcenter,
					LocationTo = gate,
					ProductionCenter = prodcenter,
					TotalPassengers = 1,
                    IsMPP = true
                };

				shuttleModelList.Add(newModel);
				shuttleModelList.Add(newModel2);
			}
		}
		#endregion
	}
}
