using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	public class CalendarController : BaseController<CalendarModel>
	{
		private readonly ICalendarAppService _calendarAppService;
		private readonly ICalendarHolidayAppService _calendarHolidayAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IReferenceDetailAppService _refDetailAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public CalendarController(ICalendarAppService calendarAppService,
			ICalendarHolidayAppService calendarHolidayAppService,
			IReferenceAppService referenceAppService,
			ILocationAppService locationAppService,
			ILoggerAppService logger,
			IMenuAppService menuService,
			IReferenceDetailAppService refDetailAppService)
		{
			_menuService = menuService;
			_calendarAppService = calendarAppService;
			_calendarHolidayAppService = calendarHolidayAppService;
			_referenceAppService = referenceAppService;
			_refDetailAppService = refDetailAppService;
			_locationAppService = locationAppService;
			_logger = logger;
		}

		public ActionResult GenerateExcel(long locID, int year, long gtypeID)
		{
			try
			{
				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("GroupTypeID", gtypeID.ToString()));
				filters.Add(new QueryFilter("LocationID", locID.ToString()));

				string calendarList = _calendarAppService.Find(filters);
				List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList().Where(x => x.Date.Year == year).ToList();

				if (calendarData.Count == 0)
				{
					SetFalseTempData(UIResources.NoDataInSelectedCriteria);
					return RedirectToAction("Index");
				}

				string location = calendarData[0].Location;
				string groupType = GetGroupType(calendarData[0].GroupTypeID);

				byte[] excelData = ExcelGenerator.ExportCalendar(calendarData, location, groupType);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Calendar.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.GenerateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult GenerateExcelHoliday(long locID, int year)
		{
			try
			{
				List<long> listOfID = GetParentIDList(locID);

				string holidays = _calendarHolidayAppService.GetAll(true);
				List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList();
				List<CalendarHolidayModel> calendarData = resultHolidaysAll.Where(x => listOfID.Any(y => y == x.LocationID) && x.Date.Year == year).ToList();
				calendarData = calendarData.OrderBy(x => x.LocationID).ThenBy(x => x.Date).ToList();

				if (calendarData.Count == 0)
				{
					SetFalseTempData(UIResources.NoDataInSelectedCriteria);
					return RedirectToAction("Index");
				}

				List<CalendarHolidayModel> result = new List<CalendarHolidayModel>();
				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var itemC in calendarData)
				{
					string ht = _referenceAppService.GetDetailById(itemC.HolidayTypeID);
					ReferenceDetailModel htModel = ht.DeserializeToRefDetail();
					itemC.Color = string.IsNullOrEmpty(htModel.Code) ? itemC.Color : htModel.Code;
					result.Add(itemC);

					if (locationMap.ContainsKey(itemC.LocationID))
					{
						string loc;
						locationMap.TryGetValue(itemC.LocationID, out loc);
						itemC.Location = loc;
					}
					else
					{
						itemC.Location = _locationAppService.GetLocationFullCode(itemC.LocationID);
						locationMap.Add(itemC.LocationID, itemC.Location);
					}
				}

				byte[] excelData = ExcelGenerator.ExportCalendarHoliday(calendarData, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Holiday.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.GenerateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Index()
		{
			GetTempData();

			CalendarModel model = GetIndexModel();

			return View(model);
		}

		private CalendarModel GetIndexModel()
		{
			ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupType(_refDetailAppService);
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.YearList = DropDownHelper.BindDropDownYear();

			CalendarModel model = new CalendarModel();
			model.Access = GetAccess(WebConstants.MenuSlug.CALENDAR_UPLOAD, _menuService);
			model.PcID = AccountProdCenterID;
			model.DepID = AccountDepartmentID;
			model.SubDepID = AccountLocationID;

			return model;
		}

		[HttpPost]
		public ActionResult GetHolidayType()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();

			string reference = _referenceAppService.GetBy("Name", "HT", true);
			ReferenceModel refModel = reference.DeserializeToReference();
			if (refModel != null)
			{
				string dataList = _referenceAppService.FindDetailBy("ReferenceID", refModel.ID, true);
				List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

				foreach (var data in dataModelList)
				{
					_menuList.Add(new SelectListItem
					{
						Text = data.Description,
						Value = data.Code
					});
				}
			}

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult CheckCalendar(long locID, int year, long gtypeID)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("GroupTypeID", gtypeID.ToString()));
			filters.Add(new QueryFilter("LocationID", locID.ToString()));

			string calendarList = _calendarAppService.Find(filters);
			List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList();

			List<SelectListItem> _dataList = new List<SelectListItem>();
			if (calendarData.Any(x => x.Date.Year == year))
			{
				_dataList.Add(new SelectListItem
				{
					Text = locID.ToString(),
					Value = locID.ToString(),
				});
			}

			return Json(_dataList, JsonRequestBehavior.AllowGet);
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
		[ValidateAntiForgeryToken]
		public ActionResult Index(CalendarModel model)
		{
			IExcelDataReader reader = null;
			try
			{
				ViewBag.YearList = DropDownHelper.BindDropDownYear();
				ViewBag.GroupTypeList = DropDownHelper.BindDropDownGroupType(_refDetailAppService);
				ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

				model.Access = GetAccess(WebConstants.MenuSlug.CALENDAR_UPLOAD, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (model.PostedFilename != null && model.PostedFilename.ContentLength > 0)
				{
					Stream stream = model.PostedFilename.InputStream;

					if (model.PostedFilename.FileName.ToLower().EndsWith(".xls"))
					{
						reader = ExcelReaderFactory.CreateBinaryReader(stream);
					}
					else if (model.PostedFilename.FileName.ToLower().EndsWith(".xlsx"))
					{
						reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
					}
					else
					{
						SetFalseTempData(UIResources.FileNotSupported);
						return RedirectToAction("Index");
					}

					int fieldcount = reader.FieldCount;
					int rowcount = reader.RowCount;
					DataTable dt = new DataTable();
					DataTable dt_ = reader.AsDataSet().Tables[0];

					List<CalendarModel> result = new List<CalendarModel>();
					string jobtitle = string.Empty;
					string username = AccountName;

					for (int index = 1; index <= dt_.Rows.Count; index++)
					{
						string yearmonth = dt_.Rows[index][0].ToString();

						for (int col = 1; col < dt_.Columns.Count; col++)
						{
							string date = dt_.Rows[index][col].ToString();
							if (string.IsNullOrEmpty(date))
								break;

							date = int.Parse(date).ToString("00");
							string fullDate = yearmonth + date;

							DateTime calendarDate;
							if (!DateTime.TryParseExact(fullDate, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, out calendarDate))
							{
								SetFalseTempData(string.Format(UIResources.DateInvalid, fullDate));
								return RedirectToAction("Index");
							}

							string location = string.Empty;

							CalendarModel newCalendar = new CalendarModel();
							newCalendar.ModifiedBy = username;
							newCalendar.ModifiedDate = DateTime.Now;
							newCalendar.GroupTypeID = model.GroupTypeID;
							newCalendar.Date = calendarDate;
							newCalendar.Shift1 = dt_.Rows[index + 1][col].ToString();
							newCalendar.Shift2 = dt_.Rows[index + 2][col].ToString();
							newCalendar.Shift3 = dt_.Rows[index + 3][col].ToString();

							result.Add(newCalendar);
						}

						index += 4;
					}

					reader.Close();
					reader.Dispose();

					List<long> locationIdList = _locationAppService.GetLocIDListByLocType(model.ProdCenterID, "productioncenter");
					List<CalendarModel> totalResult = new List<CalendarModel>();

					foreach (var locID in locationIdList)
					{
						string location = _locationAppService.GetLocationFullCode(locID);
						foreach (var item in result)
						{
							CalendarModel newCalendar = new CalendarModel();
							newCalendar.ModifiedBy = item.ModifiedBy;
							newCalendar.ModifiedDate = item.ModifiedDate;
							newCalendar.GroupTypeID = item.GroupTypeID;
							newCalendar.Date = item.Date;
							newCalendar.Shift1 = item.Shift1;
							newCalendar.Shift2 = item.Shift2;
							newCalendar.Shift3 = item.Shift3;
							newCalendar.Location = location;
							newCalendar.LocationID = locID;

							totalResult.Add(newCalendar);
						}
					}

					List<CalendarModel> calendarList = new List<CalendarModel>();
					foreach (var locID in locationIdList)
					{
						ICollection<QueryFilter> filterCal = new List<QueryFilter>();
						filterCal.Add(new QueryFilter("GroupTypeID", model.GroupTypeID.ToString()));
						filterCal.Add(new QueryFilter("LocationID", locID.ToString()));

						string calendars = _calendarAppService.Find(filterCal);
						var tempList = calendars.DeserializeToCalendarList();

						if (tempList.Count > 0)
							calendarList.AddRange(tempList);
					}

					List<CalendarModel> addCalendarList = new List<CalendarModel>();
					List<CalendarModel> updateCalendarList = new List<CalendarModel>();

					foreach (var item in totalResult)
					{
						CalendarModel calModel = calendarList.Where(x => x.Date == item.Date && x.LocationID == item.LocationID).FirstOrDefault();
						if (calModel == null)
						{
							addCalendarList.Add(item);
							//string calendar = JsonHelper<CalendarModel>.Serialize(item);
							//_calendarAppService.Add(calendar);
						}
						else
						{
							calModel.Shift1 = item.Shift1;
							calModel.Shift2 = item.Shift2;
							calModel.Shift3 = item.Shift3;
							calModel.ModifiedBy = AccountName;
							calModel.ModifiedDate = DateTime.Now;

							updateCalendarList.Add(calModel);
							//string calendar = JsonHelper<CalendarModel>.Serialize(calModel);
							//_calendarAppService.Update(calendar);
						}
					}

					if (AddOrUpdateCalendar(addCalendarList, updateCalendarList))
					{
						SetTrueTempData(UIResources.UploadSucceed);
					}
					else
					{
						SetFalseTempData("Upload failed. Please try it again");
					}
				}
				else
				{
					SetFalseTempData(UIResources.FileCorrupted);
				}
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.UploadFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				reader = null;
			}

			return RedirectToAction("Index");
		}

		private bool AddOrUpdateCalendar(List<CalendarModel> addCalendarList, List<CalendarModel> updateCalendarList)
		{
			string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
			string addCalendar = "INSERT [dbo].[Calendars] ([Shift1], [Shift2], [Shift3], [LocationID], [Location], [GroupTypeID], [Date], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES " +
							"(@Shift1, @Shift2, @Shift3, @LocationID, @Location, @GroupTypeID, @Date, @IsDeleted, @ModifiedBy, @ModifiedDate)";
			string updateCalendar = "UPDATE [dbo].[Calendars] SET[Shift1] = @Shift1, [Shift2] = @Shift2, [Shift3] = @Shift3, [ModifiedBy] = @ModifiedBy, [ModifiedDate] = @ModifiedDate WHERE([ID] = @ID)";

			using (SqlConnection connection = new SqlConnection(strConString))
			{
				connection.Open();
				SqlTransaction transaction = connection.BeginTransaction();

				try
				{
					foreach (var calendar in addCalendarList)
					{
						SqlCommand command = new SqlCommand(addCalendar, connection, transaction);
						command.Parameters.Add("@Shift1", SqlDbType.VarChar).Value = calendar.Shift1;
						command.Parameters.Add("@Shift2", SqlDbType.VarChar).Value = calendar.Shift2;
						command.Parameters.Add("@Shift3", SqlDbType.VarChar).Value = calendar.Shift3;
						command.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = calendar.LocationID;
						command.Parameters.Add("@Location", SqlDbType.VarChar).Value = calendar.Location;
						command.Parameters.Add("@GroupTypeID", SqlDbType.BigInt).Value = calendar.GroupTypeID;
						command.Parameters.Add("@Date", SqlDbType.Date).Value = calendar.Date;
						command.Parameters.Add("@IsDeleted", SqlDbType.Bit).Value = calendar.IsDeleted;
						command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = calendar.ModifiedBy;
						command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = calendar.ModifiedDate;
						command.ExecuteNonQuery();
					}

					foreach (var calendar in updateCalendarList)
					{
						SqlCommand command = new SqlCommand(updateCalendar, connection, transaction);
						command.Parameters.Add("@Shift1", SqlDbType.VarChar).Value = calendar.Shift1;
						command.Parameters.Add("@Shift2", SqlDbType.VarChar).Value = calendar.Shift2;
						command.Parameters.Add("@Shift3", SqlDbType.VarChar).Value = calendar.Shift3;
						command.Parameters.Add("@ModifiedBy", SqlDbType.VarChar).Value = calendar.ModifiedBy;
						command.Parameters.Add("@ModifiedDate", SqlDbType.DateTime).Value = calendar.ModifiedDate;
						command.Parameters.Add("@ID", SqlDbType.BigInt).Value = calendar.ID;
						command.ExecuteNonQuery();
					}

					transaction.Commit();
				}
				catch
				{
					transaction.Rollback();

					return false;
				}
			}

			return true;
		}

		[HttpPost]
		public ActionResult AddHoliday(CalendarHolidayModel model)
		{
			try
			{
				ViewBag.YearList = DropDownHelper.BindDropDownYear();
				ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
				ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
				ViewBag.ColorList = DropDownHelper.BindDropDownHolidayType(_referenceAppService);

				// get holidayType based on color code
				string ht = _referenceAppService.GetDetailBy("Code", model.Color, true);
				ReferenceDetailModel htModel = ht.DeserializeToRefDetail();
				model.HolidayTypeID = htModel.ID;
				string locType = string.Empty;

				if (model.SubDepartmentID != 0)
				{
					model.LocationID = model.SubDepartmentID;
					locType = "subdepartment";
				}
				else if (model.DepartmentID != 0)
				{
					model.LocationID = model.DepartmentID;
					locType = "department";
				}
				else if (model.ProdCenterID != 0)
				{
					model.LocationID = model.ProdCenterID;
					locType = "productioncenter";
				}
				else
				{
					model.LocationID = model.CountryID;
					locType = "country";
				}

				string calendarList = _calendarAppService.GetAll(true);
				List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList();

				// check if already exist in parent location
				if (IsExistInParentLocation(locType, model.LocationID, model.Date))
				{
					SetFalseTempData("The date " + model.Date.ToString("dd-MMM-yy") + " already defined in parent location");
					return RedirectToAction("Holiday");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("Date", model.Date.ToString()));
				filters.Add(new QueryFilter("LocationID", model.LocationID.ToString()));

				string exist = _calendarHolidayAppService.Get(filters, true);
				if (string.IsNullOrEmpty(exist))
				{
					string data = JsonHelper<CalendarHolidayModel>.Serialize(model);
					_calendarHolidayAppService.Add(data);
				}
				else
				{
					CalendarHolidayModel holidaymodel = exist.DeserializeToCalendarHoliday();
					holidaymodel.HolidayTypeID = model.HolidayTypeID;
					holidaymodel.Description = model.Description;
					holidaymodel.Color = model.Color;

					string data = JsonHelper<CalendarHolidayModel>.Serialize(holidaymodel);
					_calendarHolidayAppService.Update(data);
				}

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Holiday");
		}

		private bool IsExistInParentLocation(string locationType, long locationID, DateTime date)
		{
			if (locationType == "productioncenter")
			{
				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("Date", date.ToString()));
				filters.Add(new QueryFilter("LocationID", 1));

				string exist = _calendarHolidayAppService.Get(filters, true);

				return string.IsNullOrEmpty(exist) ? false : true;
			}
			else if (locationType == "department")
			{
				LocationModel departmentModel = GetLocation(locationID);

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("Date", date.ToString()));
				filters.Add(new QueryFilter("LocationID", departmentModel.ParentID.ToString()));

				string exist = _calendarHolidayAppService.Get(filters, true);
				if (string.IsNullOrEmpty(exist))
				{
					LocationModel pcModel = GetLocation(departmentModel.ParentID);

					filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("Date", date.ToString()));
					filters.Add(new QueryFilter("LocationID", pcModel.ParentID.ToString()));

					exist = _calendarHolidayAppService.Get(filters, true);

					if (string.IsNullOrEmpty(exist))
					{
						filters = new List<QueryFilter>();
						filters.Add(new QueryFilter("Date", date.ToString()));
						filters.Add(new QueryFilter("LocationID", 1));

						exist = _calendarHolidayAppService.Get(filters, true);

						return string.IsNullOrEmpty(exist) ? false : true;
					}
				}
			}
			else if (locationType == "subdepartment")
			{
				LocationModel subdepartmentModel = GetLocation(locationID);

				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("Date", date.ToString()));
				filters.Add(new QueryFilter("LocationID", subdepartmentModel.ParentID.ToString()));

				string exist = _calendarHolidayAppService.Get(filters, true);
				if (string.IsNullOrEmpty(exist))
				{
					LocationModel departmentModel = GetLocation(subdepartmentModel.ParentID);

					filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("Date", date.ToString()));
					filters.Add(new QueryFilter("LocationID", departmentModel.ParentID.ToString()));

					exist = _calendarHolidayAppService.Get(filters, true);
					if (string.IsNullOrEmpty(exist))
					{
						LocationModel pcModel = GetLocation(departmentModel.ParentID);

						filters = new List<QueryFilter>();
						filters.Add(new QueryFilter("Date", date.ToString()));
						filters.Add(new QueryFilter("LocationID", pcModel.ParentID.ToString()));

						exist = _calendarHolidayAppService.Get(filters, true);

						if (string.IsNullOrEmpty(exist))
						{
							filters = new List<QueryFilter>();
							filters.Add(new QueryFilter("Date", date.ToString()));
							filters.Add(new QueryFilter("LocationID", 1));

							exist = _calendarHolidayAppService.Get(filters, true);

							return string.IsNullOrEmpty(exist) ? false : true;
						}
					}
				}
			}
			else
			{
				return false;
			}

			return true;
		}

		[CustomAuthorize("calendarholiday")]
		public ActionResult Holiday()
		{
			ViewBag.YearList = DropDownHelper.BindDropDownYear();
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.ColorList = DropDownHelper.BindDropDownHolidayType(_referenceAppService);

			GetTempData();

			return View();
		}

		[HttpPost]
		public ActionResult DeleteHoliday(long id)
		{
			try
			{
				//CalendarHolidayModel holiday = GetCalendarHoliday(id);
				//holiday.IsDeleted = true;

				//string userData = JsonHelper<CalendarHolidayModel>.Serialize(holiday);
				//_calendarHolidayAppService.Update(userData);
				//hard delete calendarholiday saja
				_calendarHolidayAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		private CalendarHolidayModel GetCalendarHoliday(long id)
		{
			string holiday = _calendarHolidayAppService.GetById(id, true);
			CalendarHolidayModel holidaymodel = holiday.DeserializeToCalendarHoliday();

			return holidaymodel;
		}

		[HttpPost]
		public ActionResult GetAllHoliday()
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
			string holidays = _calendarHolidayAppService.GetAll(true);

			List<CalendarHolidayModel> result = holidays.DeserializeToCalendarHolidayList().OrderBy(x => x.LocationID).ThenBy(n => n.Date).ToList();

			result = result.OrderBy(x => x.Location).ToList();

			int recordsTotal = result.Count();
			bool isLoaded = false;

			sortColumn = sortColumn == "ID" ? "" : sortColumn;

			if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
			{
				isLoaded = true;
				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}
				}
			}

			// Search    
			if (!string.IsNullOrEmpty(searchValue))
			{
				result = result.Where(m => m.Description.ToLower().Contains(searchValue.ToLower()) ||
										   m.Date.ToString("MMMM").ToLower().Contains(searchValue.ToLower()) ||
										   m.Date.ToString("dd").ToLower().Contains(searchValue.ToLower()) ||
										   m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
			}

			if (!string.IsNullOrEmpty(sortColumn))
			{
				if (sortColumnDir == "asc")
				{
					switch (sortColumn.ToLower())
					{
						case "description":
							result = result.OrderBy(x => x.Description).ToList();
							break;
						case "location":
							result = result.OrderBy(x => x.Location).ToList();
							break;
						default:
							break;
					}
				}
				else
				{
					switch (sortColumn.ToLower())
					{
						case "description":
							result = result.OrderByDescending(x => x.Description).ToList();
							break;
						case "location":
							result = result.OrderByDescending(x => x.Location).ToList();
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

			if (!isLoaded)
			{
				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}
				}
			}

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		//[CustomAuthorize("calendarreport")]
		public ActionResult Report()
		{
			ViewBag.YearList = DropDownHelper.BindDropDownYear();
			ViewBag.DepartmentList = DropDownHelper.BindDropDownDepartmentCode(_referenceAppService, _locationAppService);
			ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
			ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

			return View();
		}

		[HttpPost]
		public ActionResult GetAllCalendar()
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
				string calendarList = _calendarAppService.GetAll(true);
				List<CalendarModel> calendarModels = calendarList.DeserializeToCalendarList();

				int recordsTotal = calendarModels.Count();

				string groupType = _referenceAppService.GetDetailAll(ReferenceEnum.Group, true);
				List<ReferenceDetailModel> groupTypeList = groupType.DeserializeToRefDetailList();

				bool isLoaded = false;

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					Dictionary<long, string> gtMap = new Dictionary<long, string>();

					foreach (var item in calendarModels)
					{
						if (locationMap.ContainsKey(item.LocationID))
						{
							string loc;
							locationMap.TryGetValue(item.LocationID, out loc);
							item.Location = loc;
						}
						else
						{
							item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
							locationMap.Add(item.LocationID, item.Location);
						}

						if (gtMap.ContainsKey(item.GroupTypeID))
						{
							string gt;
							gtMap.TryGetValue(item.GroupTypeID, out gt);
							item.GroupType = gt;
						}
						else
						{
							ReferenceDetailModel gt = groupTypeList.Where(x => x.ID == item.GroupTypeID).FirstOrDefault();
							item.GroupType = gt == null ? string.Empty : gt.Code;
							gtMap.Add(item.GroupTypeID, item.GroupType);
						}
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					calendarModels = calendarModels.Where(m => (m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
															   (m.GroupType != null ? m.GroupType.ToLower().Contains(searchValue.ToLower()) : false) ||
															   (m.Shift1 != null ? m.Shift1.ToLower().Contains(searchValue.ToLower()) : false) ||
															   (m.Shift2 != null ? m.Shift2.ToLower().Contains(searchValue.ToLower()) : false) ||
															   (m.Shift3 != null ? m.Shift3.ToLower().Contains(searchValue.ToLower()) : false) ||
															   (m.Date.ToString("dd-MMM-yy") != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false)).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "date":
								calendarModels = calendarModels.OrderBy(x => x.Date).ToList();
								break;
							case "location":
								calendarModels = calendarModels.OrderBy(x => x.Location).ToList();
								break;
							case "shift1":
								calendarModels = calendarModels.OrderBy(x => x.Shift1).ToList();
								break;
							case "shift2":
								calendarModels = calendarModels.OrderBy(x => x.Shift2).ToList();
								break;
							case "shift3":
								calendarModels = calendarModels.OrderBy(x => x.Shift3).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "date":
								calendarModels = calendarModels.OrderByDescending(x => x.Date).ToList();
								break;
							case "location":
								calendarModels = calendarModels.OrderByDescending(x => x.Location).ToList();
								break;
							case "shift1":
								calendarModels = calendarModels.OrderByDescending(x => x.Shift1).ToList();
								break;
							case "shift2":
								calendarModels = calendarModels.OrderByDescending(x => x.Shift2).ToList();
								break;
							case "shift3":
								calendarModels = calendarModels.OrderByDescending(x => x.Shift3).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = calendarModels.Count();

				// Paging     
				var data = calendarModels.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					Dictionary<long, string> gtMap = new Dictionary<long, string>();

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
							item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
							locationMap.Add(item.LocationID, item.Location);
						}

						if (gtMap.ContainsKey(item.GroupTypeID))
						{
							string gt;
							gtMap.TryGetValue(item.GroupTypeID, out gt);
							item.GroupType = gt;
						}
						else
						{
							item.GroupType = GetGroupType(item.GroupTypeID);
							gtMap.Add(item.GroupTypeID, item.GroupType);
						}
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<CalendarModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private string GetGroupType(long groupTypeID)
		{
			string gtype = _refDetailAppService.GetById(groupTypeID, true);
			ReferenceDetailModel gTypeModel = gtype.DeserializeToRefDetail();

			return gTypeModel.Code;
		}

		public ActionResult EditHoliday(int id)
		{
			ViewBag.ColorList = DropDownHelper.BindDropDownHolidayType(_referenceAppService);
			ViewBag.YearList = DropDownHelper.BindDropDownYear();

			CalendarHolidayModel model = GetCalendarHoliday(id);

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

		[HttpPost]
		public ActionResult EditHoliday(CalendarHolidayModel model)
		{
			try
			{
				ViewBag.YearList = DropDownHelper.BuildEmptyList();
				ViewBag.CountryList = DropDownHelper.BuildEmptyList();
				ViewBag.ProductionCenterList = DropDownHelper.BuildEmptyList();
				ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
				ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();
				ViewBag.ColorList = DropDownHelper.BuildEmptyList();

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Holiday");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				if (model.SubDepartmentID != 0)
				{
					model.LocationID = model.SubDepartmentID;
				}
				else if (model.DepartmentID != 0)
				{
					model.LocationID = model.DepartmentID;
				}
				else if (model.ProdCenterID != 0)
				{
					model.LocationID = model.ProdCenterID;
				}
				else
				{
					model.LocationID = model.CountryID;
				}

				// get holidayType based on color code
				string ht = _referenceAppService.GetDetailBy("Code", model.Color, true);
				ReferenceDetailModel htModel = ht.DeserializeToRefDetail();
				model.HolidayTypeID = htModel.ID;

				string data = JsonHelper<CalendarHolidayModel>.Serialize(model);

				_calendarHolidayAppService.Update(data);

				SetTrueTempData(UIResources.UpdateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Holiday");
		}

		[HttpPost]
		public ActionResult GetAllCalendarWithParam(string strMonth, long locID, int year, long gtypeID)
		{
			var draw = Request.Form.GetValues("draw").FirstOrDefault();

			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("GroupTypeID", gtypeID.ToString()));
			filters.Add(new QueryFilter("LocationID", locID.ToString()));

			string calendarList = _calendarAppService.Find(filters);

			List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList();

			var data = calendarData;
			int indexMonth = 0;

			if (strMonth == "Jan")
			{
				indexMonth = 1;
				DateTime startJan = new DateTime(year, 1, 1);
				DateTime endJan = new DateTime(year, 1, 31);

				data = calendarData.Where(x => x.Date >= startJan && x.Date <= endJan).ToList();
			}

			if (strMonth == "Feb")
			{
				DateTime today = DateTime.Today;

				indexMonth = 2;
				DateTime startFeb = new DateTime(year, 2, 1);
				DateTime endFeb = new DateTime(year, 2, DateTime.DaysInMonth(year, 2));

				data = calendarData.Where(x => x.Date >= startFeb && x.Date <= endFeb).ToList();
			}

			if (strMonth == "Mar")
			{
				indexMonth = 3;
				DateTime startMar = new DateTime(year, 3, 1);
				DateTime endMar = new DateTime(year, 3, 31);

				data = calendarData.Where(x => x.Date >= startMar && x.Date <= endMar).ToList();
			}

			if (strMonth == "Apr")
			{
				indexMonth = 4;
				DateTime startApr = new DateTime(year, 4, 1);
				DateTime endApr = new DateTime(year, 4, 30);

				data = calendarData.Where(x => x.Date >= startApr && x.Date <= endApr).ToList();
			}

			if (strMonth == "May")
			{
				indexMonth = 5;
				DateTime startMay = new DateTime(year, 5, 1);
				DateTime endMay = new DateTime(year, 5, 31);

				data = calendarData.Where(x => x.Date >= startMay && x.Date <= endMay).ToList();
			}

			if (strMonth == "Jun")
			{
				indexMonth = 6;
				DateTime startJun = new DateTime(year, 6, 1);
				DateTime endJun = new DateTime(year, 6, 30);

				data = calendarData.Where(x => x.Date >= startJun && x.Date <= endJun).ToList();
			}

			if (strMonth == "Jul")
			{
				indexMonth = 7;
				DateTime startJul = new DateTime(year, 7, 1);
				DateTime endJul = new DateTime(year, 7, 31);

				data = calendarData.Where(x => x.Date >= startJul && x.Date <= endJul).ToList();
			}

			if (strMonth == "Aug")
			{
				indexMonth = 8;
				DateTime startAug = new DateTime(year, 8, 1);
				DateTime endAug = new DateTime(year, 8, 31);

				data = calendarData.Where(x => x.Date >= startAug && x.Date <= endAug).ToList();
			}

			if (strMonth == "Sept")
			{
				indexMonth = 9;
				DateTime startSept = new DateTime(year, 9, 1);
				DateTime endSept = new DateTime(year, 9, 30);

				data = calendarData.Where(x => x.Date >= startSept && x.Date <= endSept).ToList();
			}

			if (strMonth == "Oct")
			{
				indexMonth = 10;
				DateTime startOct = new DateTime(year, 10, 1);
				DateTime endOct = new DateTime(year, 10, 31);

				data = calendarData.Where(x => x.Date >= startOct && x.Date <= endOct).ToList();
			}

			if (strMonth == "Nov")
			{
				indexMonth = 11;
				DateTime startNov = new DateTime(year, 11, 1);
				DateTime endNov = new DateTime(year, 11, 30);

				data = calendarData.Where(x => x.Date >= startNov && x.Date <= endNov).ToList();
			}

			if (strMonth == "Dec")
			{
				indexMonth = 12;
				DateTime startDec = new DateTime(year, 12, 1);
				DateTime endDec = new DateTime(year, 12, 31);

				data = calendarData.Where(x => x.Date >= startDec && x.Date <= endDec).ToList();
			}

			List<long> listOfID = GetParentIDList(locID);

			string holidays = _calendarHolidayAppService.GetAll(true);
			List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList().Where(x => listOfID.Any(y => y == x.LocationID) && x.Date.Month == indexMonth).OrderBy(x => x.Date).ToList();
			foreach (var item in data)
			{
				var holiday = resultHolidaysAll.Where(x => x.Date == item.Date).FirstOrDefault();
				if (holiday != null)
				{
					item.Shift1 = holiday.Color;
					item.Shift2 = holiday.Color;
					item.Shift3 = holiday.Color;
				}
			}

			Dictionary<long, string> locationMap = new Dictionary<long, string>();
			Dictionary<long, string> gtMap = new Dictionary<long, string>();

			foreach (var item in data)
			{
				item.Week = getCurrentWeekNumber(item.Date).ToString();

				if (locationMap.ContainsKey(item.LocationID))
				{
					string loc;
					locationMap.TryGetValue(item.LocationID, out loc);
					item.Location = loc;
				}
				else
				{
					item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
					locationMap.Add(item.LocationID, item.Location);
				}

				if (gtMap.ContainsKey(item.GroupTypeID))
				{
					string gt;
					gtMap.TryGetValue(item.GroupTypeID, out gt);
					item.GroupType = gt;
				}
				else
				{
					item.GroupType = GetGroupType(item.GroupTypeID);
					gtMap.Add(item.GroupTypeID, item.GroupType);
				}
			}

			// Returning Json Data    
			return Json(new { data, draw }, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetAllHolidayWithParam(long locID, string locType, int year)
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

			List<long> listOfID = GetParentIDList(locID);

			// Getting all data    			
			//string calendarList = _calendarAppService.GetAll(true);
			//List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList();
			//calendarData = calendarData.Where(x => listOfID.Any(y => y == x.LocationID) && x.Date.Year == year).ToList();

			string holidays = _calendarHolidayAppService.GetAll(true);
			List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList();
			resultHolidaysAll = resultHolidaysAll.Where(x => listOfID.Any(y => y == x.LocationID) && x.Date.Year == year).OrderBy(x => x.Date).ToList();

			List<CalendarHolidayModel> result = new List<CalendarHolidayModel>();

			foreach (var itemC in resultHolidaysAll)
			{
				string ht = _referenceAppService.GetDetailById(itemC.HolidayTypeID);
				ReferenceDetailModel htModel = ht.DeserializeToRefDetail();
				itemC.Color = string.IsNullOrEmpty(htModel.Code) ? itemC.Color : htModel.Code;
				result.Add(itemC);
			}

			int recordsTotal = result.Count();

			bool isLoaded = false;

			sortColumn = sortColumn == "ID" ? "" : sortColumn;

			if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
			{
				isLoaded = true;
				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}
				}
			}

			// Search    
			if (!string.IsNullOrEmpty(searchValue))
			{
				result = result.Where(m => m.Description.ToLower().Contains(searchValue.ToLower()) ||
										   m.Location.ToLower().Contains(searchValue.ToLower())).ToList();
			}

			if (!string.IsNullOrEmpty(sortColumn))
			{
				if (sortColumnDir == "asc")
				{
					switch (sortColumn.ToLower())
					{
						case "locationid":
							result = result.OrderBy(x => x.LocationID).ToList();
							break;
						case "description":
							result = result.OrderBy(x => x.Description).ToList();
							break;
						case "location":
							result = result.OrderBy(x => x.Location).ToList();
							break;
						default:
							break;
					}
				}
				else
				{
					switch (sortColumn.ToLower())
					{
						case "locationid":
							result = result.OrderByDescending(x => x.LocationID).ToList();
							break;
						case "description":
							result = result.OrderByDescending(x => x.Description).ToList();
							break;
						case "location":
							result = result.OrderByDescending(x => x.Location).ToList();
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

			if (!isLoaded)
			{
				Dictionary<long, string> locationMap = new Dictionary<long, string>();
				foreach (var item in result)
				{
					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}
				}
			}

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetHolidayList()
		{
			List<CalendarHolidayModel> holidayList = _calendarHolidayAppService.GetAll(true).DeserializeToCalendarHolidayList();
			holidayList = holidayList.Where(x => x.Date.Year == DateTime.Now.Year && x.LocationID == 1).OrderByDescending(x => x.Date).ToList();

			List<ReferenceDetailModel> holidayTypeList = _referenceAppService.GetDetailAll(ReferenceEnum.HolidayType).DeserializeToRefDetailList();

			foreach (var item in holidayList)
			{
				var ht = holidayTypeList.Where(x => x.ID == item.HolidayTypeID).FirstOrDefault();
				if (ht != null)
				{
					item.HolidayType = ht.Description;
				}
				item.DateStr = item.Date.ToString("dd MMMM yyyy");
			}

			return Json(holidayList, JsonRequestBehavior.AllowGet);
		}

		private List<long> GetParentIDList(long locID)
		{
			List<long> result = new List<long>();

			LocationModel model = GetLocation(locID);
			if (model.ParentID != 0)
			{
				result.Add(model.ID);
				model = GetLocation(model.ParentID);
				if (model.ParentID != 0)
				{
					result.Add(model.ID);
					model = GetLocation(model.ParentID);
					if (model.ParentID != 0)
					{
						result.Add(model.ID);
						model = GetLocation(model.ParentID);
						result.Add(model.ID);
					}
					else
					{
						result.Add(model.ID);
					}
				}
				else
				{
					result.Add(model.ID);
				}
			}
			else
			{
				result.Add(model.ID);
			}

			return result;
		}

		[HttpPost]
		public ActionResult GetWorkingSummaryWithParam(long locID, int year, long gtypeID)
		{
			var draw = Request.Form.GetValues("draw").FirstOrDefault();
			var start = Request.Form.GetValues("start").FirstOrDefault();
			var length = Request.Form.GetValues("length").FirstOrDefault();

			// Paging Size (10,20,50,100)    
			int pageSize = length != null ? Convert.ToInt32(length) : 0;
			int skip = start != null ? Convert.ToInt32(start) : 0;

			// Getting all data    			
			string calendarList = _calendarAppService.FindBy("LocationID", locID, true);
			List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList();
			calendarData = calendarData.Where(x => x.Date.Year == year && x.GroupTypeID == gtypeID).ToList();

			List<long> listOfID = GetParentIDList(locID);
			string holidays = _calendarHolidayAppService.GetAll(true);
			List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList().Where(x => listOfID.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
			resultHolidaysAll = resultHolidaysAll.Where(x => x.Date.Year == year).ToList();

			string holidayType = _referenceAppService.GetDetailAll(ReferenceEnum.HolidayType, true);
			List<ReferenceDetailModel> htModelList = holidayType.DeserializeToRefDetailList();

			string holidayCat = _referenceAppService.GetDetailAll(ReferenceEnum.CategoryHoliday, true);
			List<ReferenceDetailModel> hcModelList = holidayCat.DeserializeToRefDetailList();

			List<ReferenceDetailModel> anyholiday = hcModelList.Where(x => x.Code.ToLower().Equals("holiday")).ToList();
			List<ReferenceDetailModel> anyleave = hcModelList.Where(x => x.Code.ToLower().Equals("leave")).ToList();
			List<ReferenceDetailModel> anyshif = hcModelList.Where(x => x.Code.ToLower().Equals("shiftoff")).ToList();
			List<ReferenceDetailModel> anyprod = hcModelList.Where(x => x.Code.ToLower().Equals("prodoff")).ToList();

			ReferenceDetailModel htholiday = htModelList.Where(x => x.Description.ToLower().Contains("libur nasional") || anyholiday.Any(y => y.Description.ToLower() == x.Description.ToLower())).FirstOrDefault();
			ReferenceDetailModel htleaveHMS = htModelList.Where(x => x.Description.ToLower().Contains("cuti bersama perusahaan") || x.Description.ToLower().Contains("cuti bersama hms") || anyleave.Any(y => y.Description.ToLower() == x.Description.ToLower())).FirstOrDefault();
			ReferenceDetailModel htleavePemerintah = htModelList.Where(x => x.Description.ToLower().Contains("cuti bersama pemerintah") || anyleave.Any(y => y.Description.ToLower() == x.Description.ToLower())).LastOrDefault();
			ReferenceDetailModel htshift = htModelList.Where(x => x.Description.ToLower().Contains("shift off") || anyshif.Any(y => y.Description.ToLower() == x.Description.ToLower())).FirstOrDefault();
			ReferenceDetailModel htprod = htModelList.Where(x => x.Description.ToLower().Contains("acara internal") || anyprod.Any(y => y.Description.ToLower() == x.Description.ToLower())).FirstOrDefault();

			List<CalendarWorkingSummaryModel> result = new List<CalendarWorkingSummaryModel>();

			long idHoliday = htholiday == null ? 0 : htholiday.ID;
			long idLeaveHMS = htleaveHMS == null ? 0 : htleaveHMS.ID;
			long idLeavePemerintah = htleavePemerintah == null ? 0 : htleavePemerintah.ID;
			long idProdOff = htprod == null ? 0 : htprod.ID;
			long idShiftOff = htshift == null ? 0 : htshift.ID;

			string[] columnNames = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

			// Get January Data
			int totalDay = 0;
			int totalHoliday = 0;
			int totalLeave = 0;
			int totalProdOff = 0;
			int totalShiftOff = 0;
			for (int i = 0; i < 13; i++)
			{
				CalendarWorkingSummaryModel newModel = new CalendarWorkingSummaryModel();
				newModel.ColumnName = columnNames[i];
				if (i < 12)
				{
					var monthlyCalendarHoliday = resultHolidaysAll.Where(x => x.Date.Month == (i + 1)).ToList();
					newModel.Holiday = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idHoliday).Count();
					newModel.Leaves = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idLeaveHMS || x.HolidayTypeID == idLeavePemerintah).Count();
					newModel.ProdOff = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idProdOff).Count();
					newModel.ShiftOff = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idShiftOff).Count();
					newModel.Days = calendarData.Where(x => x.Date.Month == (i + 1)).Count();
					newModel.WorkDays = newModel.Days - (newModel.Holiday + newModel.Leaves + newModel.ProdOff + newModel.ShiftOff);

					totalDay += newModel.Days;
					totalHoliday += newModel.Holiday;
					totalLeave += newModel.Leaves;
					totalProdOff += newModel.ProdOff;
					totalShiftOff += newModel.ShiftOff;

				}
				else
				{
					newModel.Days = totalDay;
					newModel.Holiday = totalHoliday;
					newModel.Leaves = totalLeave;
					newModel.ProdOff = totalProdOff;
					newModel.ShiftOff = totalShiftOff;
					newModel.WorkDays = totalDay - (totalHoliday + totalLeave + totalProdOff + totalShiftOff);
				}

				result.Add(newModel);
			}

			int recordsTotal = result.Count();

			// total number of rows count     
			int recordsFiltered = result.Count();

			// Paging     
			var data = result.Skip(skip).Take(pageSize).ToList();

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult GetWorkingGroupSummaryWithParam(long locID, int year, long gtypeID)
		{
			var draw = Request.Form.GetValues("draw").FirstOrDefault();
			var start = Request.Form.GetValues("start").FirstOrDefault();
			var length = Request.Form.GetValues("length").FirstOrDefault();

			// Paging Size (10,20,50,100)    
			int pageSize = length != null ? Convert.ToInt32(length) : 0;
			int skip = start != null ? Convert.ToInt32(start) : 0;

			// Getting all data    			
			string calendars = _calendarAppService.FindBy("LocationID", locID, true);
			List<CalendarModel> calendarList = calendars.DeserializeToCalendarList();
			calendarList = calendarList.Where(x => x.Date.Year == year && x.GroupTypeID == gtypeID).ToList();

			List<long> listOfID = GetParentIDList(locID);
			string holidays = _calendarHolidayAppService.GetAll(true);
			List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList().Where(x => listOfID.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
			resultHolidaysAll = resultHolidaysAll.Where(x => x.Date.Year == year).ToList();
			List<DateTime> holidayList = resultHolidaysAll.Select(x => x.Date).Distinct().ToList();

			List<CalendarWorkingSummaryGroupModel> result = new List<CalendarWorkingSummaryGroupModel>();

			string[] columnNames = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

			int totalA = 0, totalB = 0, totalC = 0, totalD = 0;

			for (int i = 0; i < 13; i++)
			{
				CalendarWorkingSummaryGroupModel newModel = new CalendarWorkingSummaryGroupModel();
				newModel.ColumnName = columnNames[i];

				if (i < 12)
				{
					int month = i + 1;

					// get A amount
					var calList = calendarList.Where(x => x.Date.Month == month && (x.Shift1 == "A" || x.Shift2 == "A" || x.Shift3 == "A")).ToList();
					newModel.A = GetWorkingDays(holidayList, calList);
					totalA += newModel.A;
					// get B amount
					calList = calendarList.Where(x => x.Date.Month == month && (x.Shift1 == "B" || x.Shift2 == "B" || x.Shift3 == "B")).ToList();
					newModel.B = GetWorkingDays(holidayList, calList);
					totalB += newModel.B;
					// get C amount
					calList = calendarList.Where(x => x.Date.Month == month && (x.Shift1 == "C" || x.Shift2 == "C" || x.Shift3 == "C")).ToList();
					newModel.C = GetWorkingDays(holidayList, calList);
					totalC += newModel.C;
					// get D amount
					calList = calendarList.Where(x => x.Date.Month == month && (x.Shift1 == "D" || x.Shift2 == "D" || x.Shift3 == "D")).ToList();
					newModel.D = GetWorkingDays(holidayList, calList);
					totalD += newModel.D;
				}
				else
				{
					newModel.A = totalA;
					newModel.B = totalB;
					newModel.C = totalC;
					newModel.D = totalD;
				}

				result.Add(newModel);
			}

			int recordsTotal = result.Count();

			// total number of rows count     
			int recordsFiltered = result.Count();

			// Paging     
			var data = result.Skip(skip).Take(pageSize).ToList();

			// Returning Json Data    
			return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
		}

		private int GetWorkingDays(List<DateTime> holidayList, List<CalendarModel> calendarList)
		{
			int result = 0;
			foreach (var item in calendarList)
			{
				if (!holidayList.Any(x => x == item.Date))
					result++;
			}

			return result;
		}

		private LocationModel GetLocation(long locationID)
		{
			string location = _locationAppService.GetById(locationID, true);
			LocationModel locationModel = location.DeserializeToLocation();

			return locationModel;
		}

		public ActionResult DownloadCalendarTemplate()
		{
			try
			{
				string filepath = Server.MapPath("..") + "\\Templates\\TemplateCalendar.xlsx";

				if (System.IO.File.Exists(filepath))
				{
					byte[] fileBytes = GetFile(filepath);

					return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateCalendar.xlsx");
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult DownloadHolidayTemplate()
		{
			try
			{
				string filepath = Server.MapPath("..") + "\\Templates\\TemplateHoliday.xlsx";

				if (System.IO.File.Exists(filepath))
				{
					byte[] fileBytes = GetFile(filepath);

					return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateHoliday.xlsx");
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Holiday");
		}

		private byte[] GetFile(string filepath)
		{
			FileStream fs = System.IO.File.OpenRead(filepath);
			byte[] data = new byte[fs.Length];
			int br = fs.Read(data, 0, data.Length);
			if (br != fs.Length)
				throw new System.IO.IOException(filepath);
			return data;
		}

		[HttpPost]
		public ActionResult GeneratePDF(CalendarModel model)
		{
			if (model.SubDepartmentID != 0)
			{
				model.LocationID = model.SubDepartmentID;
			}
			else if (model.DepartmentID != 0)
			{
				model.LocationID = model.DepartmentID;
			}
			else
			{
				model.LocationID = model.ProdCenterID;
			}

			model.Location = _locationAppService.GetLocationFullCode(model.LocationID);

			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("GroupTypeID", model.GroupTypeID.ToString()));
			filters.Add(new QueryFilter("LocationID", model.LocationID.ToString()));

			string calendarList = _calendarAppService.Find(filters);
			List<CalendarModel> calendarData = calendarList.DeserializeToCalendarList().Where(x => x.Date.Year == model.Year).ToList();
			if (calendarData.Count == 0)
			{
				SetFalseTempData(UIResources.NoDataInSelectedCriteria);
				return RedirectToAction("Index");
			}

			#region Get Data
			string holidays = _calendarHolidayAppService.GetAll(true);
			List<long> listOfID = GetParentIDList(model.LocationID);
			List<CalendarHolidayModel> resultHolidaysAll = holidays.DeserializeToCalendarHolidayList().Where(x => listOfID.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();
			List<CalendarHolidayModel> resultHolidays = new List<CalendarHolidayModel>();

			foreach (var itemC in resultHolidaysAll)
			{
				var checkC = calendarData.Where(x => x.Date.Year.ToString().ToLower().Contains(itemC.Date.Year.ToString()));
				if (checkC.Count() > 0)
				{
					string ht = _referenceAppService.GetDetailById(itemC.HolidayTypeID);
					ReferenceDetailModel htModel = ht.DeserializeToRefDetail();
					itemC.Color = string.IsNullOrEmpty(htModel.Code) ? itemC.Color : htModel.Code;
					resultHolidays.Add(itemC);
				}
			}

			var data = calendarData;

			DateTime startJan = new DateTime(model.Year, 1, 1);
			DateTime endJan = new DateTime(model.Year, 1, 31);
			var dataJan = calendarData.Where(x => x.Date >= startJan && x.Date <= endJan).ToList();

			DateTime startFeb = new DateTime(model.Year, 2, 1);
			DateTime endFeb = new DateTime(model.Year, 2, DateTime.DaysInMonth(model.Year, 2));
			var dataFeb = calendarData.Where(x => x.Date >= startFeb && x.Date <= endFeb).ToList();

			DateTime startMar = new DateTime(model.Year, 3, 1);
			DateTime endMar = new DateTime(model.Year, 3, 31);
			var dataMar = calendarData.Where(x => x.Date >= startMar && x.Date <= endMar).ToList();

			DateTime startApr = new DateTime(model.Year, 4, 1);
			DateTime endApr = new DateTime(model.Year, 4, 30);
			var dataApr = calendarData.Where(x => x.Date >= startApr && x.Date <= endApr).ToList();

			DateTime startMay = new DateTime(model.Year, 5, 1);
			DateTime endMay = new DateTime(model.Year, 5, 31);
			var dataMay = calendarData.Where(x => x.Date >= startMay && x.Date <= endMay).ToList();

			DateTime startJun = new DateTime(model.Year, 6, 1);
			DateTime endJun = new DateTime(model.Year, 6, 30);
			var dataJun = calendarData.Where(x => x.Date >= startJun && x.Date <= endJun).ToList();

			DateTime startJul = new DateTime(model.Year, 7, 1);
			DateTime endJul = new DateTime(model.Year, 7, 31);
			var dataJul = calendarData.Where(x => x.Date >= startJul && x.Date <= endJul).ToList();

			DateTime startAug = new DateTime(model.Year, 8, 1);
			DateTime endAug = new DateTime(model.Year, 8, 31);
			var dataAug = calendarData.Where(x => x.Date >= startAug && x.Date <= endAug).ToList();

			DateTime startSept = new DateTime(model.Year, 9, 1);
			DateTime endSept = new DateTime(model.Year, 9, 30);
			var dataSept = calendarData.Where(x => x.Date >= startSept && x.Date <= endSept).ToList();

			DateTime startOct = new DateTime(model.Year, 10, 1);
			DateTime endOct = new DateTime(model.Year, 10, 31);
			var dataOct = calendarData.Where(x => x.Date >= startOct && x.Date <= endOct).ToList();

			DateTime startNov = new DateTime(model.Year, 11, 1);
			DateTime endNov = new DateTime(model.Year, 11, 30);
			var dataNov = calendarData.Where(x => x.Date >= startNov && x.Date <= endNov).ToList();

			DateTime startDec = new DateTime(model.Year, 12, 1);
			DateTime endDec = new DateTime(model.Year, 12, 31);
			var dataDec = calendarData.Where(x => x.Date >= startDec && x.Date <= endDec).ToList();

			#endregion

			#region Define Document, Column Table , Font
			BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.EMBEDDED);
			Font font = new Font(bf, 9);
			Font fontWk = new Font(bf, 7);

			Document pdfDoc = new Document(PageSize.A4.Rotate(), 20, 20, 20, 20);
			PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
			pdfDoc.Open();

			Image image = Image.GetInstance(Server.MapPath("~/Content/theme/images/fast-blue.jpg"));
			pdfDoc.Add(image);

			pdfDoc.Add(new Paragraph(new Chunk("Title          : " + model.Location + " Calendar " + model.Year, font)));
			pdfDoc.Add(new Paragraph(new Chunk("Generated By   : " + AccountName, font)));
			pdfDoc.Add(new Paragraph(new Chunk("Generated Date : " + DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), font)));

			int[] tblWidth31 = { 5, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2 };
			int[] tblWidth30 = { 5, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2 };
			int[] tblWidth29 = { 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
			int[] tblWidth28 = { 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };

			#endregion

			#region Januari

			string monthName = getMonth(new DateTime(model.Year, 1, 1));
			int daysJan = DateTime.DaysInMonth(model.Year, 1);

			PdfPTable tableJan = new PdfPTable(32);
			tableJan.WidthPercentage = 100;
			tableJan.HorizontalAlignment = 0;
			tableJan.SpacingBefore = 20f;
			tableJan.SpacingAfter = 10f;

			if (daysJan == 31)
			{
				tableJan.SetWidths(tblWidth31);
			}
			// month - week
			PdfPCell mCell = new PdfPCell(new Phrase(monthName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableJan.AddCell(mCell);
			string weekTemp = "0";
			foreach (var item in dataJan)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableJan.AddCell(wCell);
				}
				else
				{
					if (dataJan.IndexOf(item) == dataJan.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJan.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableJan.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJan.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableJan.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}
			// day name
			PdfPCell dCell = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJan.AddCell(dCell);
			foreach (var item in dataJan)
			{
				var checkHJan = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJan.Count() > 0)
				{
					foreach (var itemHJan in resultHolidays)
					{
						if (itemHJan.Date == item.Date)
						{
							var colorHJan = HexToColor(itemHJan.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHJan);
							tableJan.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJan.AddCell(dyCell);
				}

			}
			// date 
			PdfPCell dtCell = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJan.AddCell(dtCell);
			foreach (var item in dataJan)
			{
				var checkHJan = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJan.Count() > 0)
				{
					foreach (var itemHJan in resultHolidays)
					{
						if (itemHJan.Date == item.Date)
						{
							var colorHJan = HexToColor(itemHJan.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHJan);
							tableJan.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJan.AddCell(dteCell);
				}

			}

			// shift 1
			PdfPCell sCell1 = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJan.AddCell(sCell1);
			foreach (var item in dataJan)
			{
				var checkHJan = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJan.Count() > 0)
				{
					foreach (var itemHJan in resultHolidays)
					{
						if (itemHJan.Date == item.Date)
						{
							var colorHJan = HexToColor(itemHJan.Color);
							string shift1Val = "";
							if (itemHJan.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJan);
							tableJan.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJan.AddCell(cellColJ);
				}


			}
			// shift 2
			PdfPCell sCell2 = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJan.AddCell(sCell2);
			foreach (var item in dataJan)
			{
				var checkHJan = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJan.Count() > 0)
				{
					foreach (var itemHJan in resultHolidays)
					{
						if (itemHJan.Date == item.Date)
						{
							var colorHJan = HexToColor(itemHJan.Color);
							string shift2Val = "";
							if (itemHJan.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJan);
							tableJan.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJan.AddCell(cellColJ);
				}

			}
			// shift 3
			PdfPCell sCell3 = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJan.AddCell(sCell3);
			foreach (var item in dataJan)
			{
				var checkHJan = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJan.Count() > 0)
				{
					foreach (var itemHJan in resultHolidays)
					{
						if (itemHJan.Date == item.Date)
						{
							var colorHJan = HexToColor(itemHJan.Color);
							string shift3Val = "";
							if (itemHJan.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJan);
							tableJan.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJan.AddCell(cellColJ);
				}


			}


			pdfDoc.Add(tableJan);
			#endregion

			#region Februari

			string monthNameFeb = getMonth(new DateTime(model.Year, 2, 1));
			int daysFeb = DateTime.DaysInMonth(model.Year, 2);
			int colFeb = 0;

			if (daysFeb == 29)
			{
				colFeb = 30;
			}
			if (daysFeb == 28)
			{
				colFeb = 29;
			}

			PdfPTable tableFeb = new PdfPTable(colFeb);
			tableFeb.WidthPercentage = 100;
			tableFeb.HorizontalAlignment = 0;
			tableFeb.SpacingBefore = 15f;
			tableFeb.SpacingAfter = 10f;

			if (daysFeb == 29)
			{
				tableFeb.SetWidths(tblWidth29);
			}
			if (daysFeb == 28)
			{
				tableFeb.SetWidths(tblWidth28);
			}

			PdfPCell mCellF = new PdfPCell(new Phrase(monthNameFeb, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellF.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableFeb.AddCell(mCellF);
			foreach (var item in dataFeb)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableFeb.AddCell(wCell);
				}
				else
				{
					if (dataFeb.IndexOf(item) == dataFeb.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableFeb.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableFeb.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableFeb.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableFeb.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellF = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableFeb.AddCell(dCellF);
			foreach (var item in dataFeb)
			{
				var checkHFeb = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHFeb.Count() > 0)
				{
					foreach (var itemHFeb in resultHolidays)
					{
						if (itemHFeb.Date == item.Date)
						{
							var colorHFeb = HexToColor(itemHFeb.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHFeb);
							tableFeb.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableFeb.AddCell(dyCell);
				}


			}

			PdfPCell dtCellF = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableFeb.AddCell(dtCellF);
			foreach (var item in dataFeb)
			{
				var checkHFeb = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHFeb.Count() > 0)
				{
					foreach (var itemHFeb in resultHolidays)
					{
						if (itemHFeb.Date == item.Date)
						{
							var colorHFeb = HexToColor(itemHFeb.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHFeb);
							tableFeb.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableFeb.AddCell(dteCell);
				}

			}

			PdfPCell sCell1F = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableFeb.AddCell(sCell1F);
			foreach (var item in dataFeb)
			{
				var checkHFeb = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHFeb.Count() > 0)
				{
					foreach (var itemHFeb in resultHolidays)
					{
						if (itemHFeb.Date == item.Date)
						{
							var colorHFeb = HexToColor(itemHFeb.Color);
							string shift1Val = "";
							if (itemHFeb.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHFeb);
							tableFeb.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableFeb.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2F = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableFeb.AddCell(sCell2F);
			foreach (var item in dataFeb)
			{
				var checkHFeb = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHFeb.Count() > 0)
				{
					foreach (var itemHFeb in resultHolidays)
					{
						if (itemHFeb.Date == item.Date)
						{
							var colorHFeb = HexToColor(itemHFeb.Color);
							string shift2Val = "";
							if (itemHFeb.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHFeb);
							tableFeb.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableFeb.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3F = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableFeb.AddCell(sCell3F);
			foreach (var item in dataFeb)
			{
				var checkHFeb = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHFeb.Count() > 0)
				{
					foreach (var itemHFeb in resultHolidays)
					{
						if (itemHFeb.Date == item.Date)
						{
							var colorHFeb = HexToColor(itemHFeb.Color);
							string shift3Val = "";
							if (itemHFeb.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHFeb);
							tableFeb.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableFeb.AddCell(cellColJ);
				}

			}


			pdfDoc.Add(tableFeb);
			#endregion

			#region Maret

			string monthNameMar = getMonth(new DateTime(model.Year, 3, 1));
			int daysMar = DateTime.DaysInMonth(model.Year, 3);
			int colMar = 0;

			if (daysMar == 30)
			{
				colMar = 31;
			}
			if (daysMar == 31)
			{
				colMar = 32;
			}

			PdfPTable tableMar = new PdfPTable(colMar);
			tableMar.WidthPercentage = 100;
			tableMar.HorizontalAlignment = 0;
			tableMar.SpacingBefore = 15f;
			tableMar.SpacingAfter = 10f;

			if (daysMar == 31)
			{
				tableMar.SetWidths(tblWidth31);
			}
			if (daysFeb == 30)
			{
				tableMar.SetWidths(tblWidth30);
			}

			PdfPCell mCellM = new PdfPCell(new Phrase(monthNameMar, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellM.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableMar.AddCell(mCellM);
			foreach (var item in dataMar)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableMar.AddCell(wCell);
				}
				else
				{
					if (dataMar.IndexOf(item) == dataMar.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableMar.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableMar.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableMar.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableMar.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellM = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMar.AddCell(dCellM);
			foreach (var item in dataMar)
			{
				var checkHMar = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMar.Count() > 0)
				{
					foreach (var itemHMar in resultHolidays)
					{
						if (itemHMar.Date == item.Date)
						{
							var colorHMar = HexToColor(itemHMar.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHMar);
							tableMar.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableMar.AddCell(dyCell);
				}
			}

			PdfPCell dtCellM = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMar.AddCell(dtCellM);
			foreach (var item in dataMar)
			{
				var checkHMar = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMar.Count() > 0)
				{
					foreach (var itemHMar in resultHolidays)
					{
						if (itemHMar.Date == item.Date)
						{
							var colorHMar = HexToColor(itemHMar.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHMar);
							tableMar.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableMar.AddCell(dteCell);
				}

			}
			PdfPCell sCell1M = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMar.AddCell(sCell1M);
			foreach (var item in dataMar)
			{
				var checkHMar = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMar.Count() > 0)
				{
					foreach (var itemHMar in resultHolidays)
					{
						if (itemHMar.Date == item.Date)
						{
							var colorHMar = HexToColor(itemHMar.Color);
							string shift1Val = "";
							if (itemHMar.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMar);
							tableMar.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMar.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2M = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMar.AddCell(sCell2M);
			foreach (var item in dataMar)
			{
				var checkHMar = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMar.Count() > 0)
				{
					foreach (var itemHMar in resultHolidays)
					{
						if (itemHMar.Date == item.Date)
						{
							var colorHMar = HexToColor(itemHMar.Color);
							string shift2Val = "";
							if (itemHMar.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMar);
							tableMar.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMar.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3M = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMar.AddCell(sCell3M);
			foreach (var item in dataMar)
			{
				var checkHMar = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMar.Count() > 0)
				{
					foreach (var itemHMar in resultHolidays)
					{
						if (itemHMar.Date == item.Date)
						{
							var colorHMar = HexToColor(itemHMar.Color);
							string shift3Val = "";
							if (itemHMar.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMar);
							tableMar.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMar.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableMar);

			#endregion

			#region April

			string monthNameApr = getMonth(new DateTime(model.Year, 4, 1));
			int daysApr = DateTime.DaysInMonth(model.Year, 4);
			int colApr = 0;

			if (daysApr == 30)
			{
				colApr = 31;
			}
			if (daysApr == 31)
			{
				colApr = 32;
			}

			PdfPTable tableApr = new PdfPTable(colApr);
			tableApr.WidthPercentage = 100;
			tableApr.HorizontalAlignment = 0;
			tableApr.SpacingBefore = 15f;
			tableApr.SpacingAfter = 10f;

			if (daysApr == 31)
			{
				tableApr.SetWidths(tblWidth31);
			}
			if (daysApr == 30)
			{
				tableApr.SetWidths(tblWidth30);
			}

			PdfPCell mCellA = new PdfPCell(new Phrase(monthNameApr, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellA.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableApr.AddCell(mCellA);
			foreach (var item in dataApr)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableApr.AddCell(wCell);
				}
				else
				{
					if (dataApr.IndexOf(item) == dataApr.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableApr.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableApr.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableApr.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableApr.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellA = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableApr.AddCell(dCellA);
			foreach (var item in dataApr)
			{
				var checkHApr = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHApr.Count() > 0)
				{
					foreach (var itemHApr in resultHolidays)
					{
						if (itemHApr.Date == item.Date)
						{
							var colorHApr = HexToColor(itemHApr.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHApr);
							tableApr.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableApr.AddCell(dyCell);
				}

			}

			PdfPCell dtCellA = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableApr.AddCell(dtCellA);
			foreach (var item in dataApr)
			{
				var checkHApr = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHApr.Count() > 0)
				{
					foreach (var itemHApr in resultHolidays)
					{
						if (itemHApr.Date == item.Date)
						{
							var colorHApr = HexToColor(itemHApr.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHApr);
							tableApr.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableApr.AddCell(dteCell);
				}

			}

			PdfPCell sCell1A = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableApr.AddCell(sCell1A);
			foreach (var item in dataApr)
			{
				var checkHApr = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHApr.Count() > 0)
				{
					foreach (var itemHApr in resultHolidays)
					{
						if (itemHApr.Date == item.Date)
						{
							var colorHApr = HexToColor(itemHApr.Color);
							string shift1Val = "";
							if (itemHApr.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHApr);
							tableApr.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableApr.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2A = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableApr.AddCell(sCell2A);
			foreach (var item in dataApr)
			{
				var checkHApr = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHApr.Count() > 0)
				{
					foreach (var itemHApr in resultHolidays)
					{
						if (itemHApr.Date == item.Date)
						{
							var colorHApr = HexToColor(itemHApr.Color);
							string shift2Val = "";
							if (itemHApr.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHApr);
							tableApr.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableApr.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3A = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableApr.AddCell(sCell3A);
			foreach (var item in dataApr)
			{
				var checkHApr = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHApr.Count() > 0)
				{
					foreach (var itemHApr in resultHolidays)
					{
						if (itemHApr.Date == item.Date)
						{
							var colorHApr = HexToColor(itemHApr.Color);
							string shift3Val = "";
							if (itemHApr.Color == "#3366cc")
							{
								shift3Val = item.Shift2;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHApr);
							tableApr.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableApr.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableApr);

			#endregion

			pdfDoc.NewPage();

			#region Mei

			string monthNameMei = getMonth(new DateTime(model.Year, 5, 1));
			int daysMei = DateTime.DaysInMonth(model.Year, 5);
			int colMei = 0;

			if (daysMei == 30)
			{
				colMei = 31;
			}
			if (daysMei == 31)
			{
				colMei = 32;
			}

			PdfPTable tableMei = new PdfPTable(colMei);
			tableMei.WidthPercentage = 100;
			tableMei.HorizontalAlignment = 0;
			tableMei.SpacingBefore = 20f;
			tableMei.SpacingAfter = 10f;

			if (daysMei == 31)
			{
				tableMei.SetWidths(tblWidth31);
			}
			if (daysMei == 30)
			{
				tableMei.SetWidths(tblWidth30);
			}

			PdfPCell mCellMe = new PdfPCell(new Phrase(monthNameMei, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellMe.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableMei.AddCell(mCellMe);
			foreach (var item in dataMay)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableMei.AddCell(wCell);
				}
				else
				{
					if (dataMay.IndexOf(item) == dataMay.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						//if (dayName == "Sun")
						//{

						//    wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						//    wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						//    tableMei.AddCell(wCell);
						//}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableMei.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableMei.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableMei.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellMe = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMei.AddCell(dCellMe);
			foreach (var item in dataMay)
			{
				var checkHMay = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMay.Count() > 0)
				{
					foreach (var itemHMay in resultHolidays)
					{
						if (itemHMay.Date == item.Date)
						{
							var colorHMay = HexToColor(itemHMay.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHMay);
							tableMei.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableMei.AddCell(dyCell);
				}
			}

			PdfPCell dtCellMe = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMei.AddCell(dtCellMe);
			foreach (var item in dataMay)
			{
				var checkHMay = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMay.Count() > 0)
				{
					foreach (var itemHMay in resultHolidays)
					{
						if (itemHMay.Date == item.Date)
						{
							var colorHMay = HexToColor(itemHMay.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHMay);
							tableMei.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableMei.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Me = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMei.AddCell(sCell1Me);
			foreach (var item in dataMay)
			{
				var checkHMay = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMay.Count() > 0)
				{
					foreach (var itemHMay in resultHolidays)
					{
						if (itemHMay.Date == item.Date)
						{
							var colorHMay = HexToColor(itemHMay.Color);
							string shift1Val = "";
							if (itemHMay.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMay);
							tableMei.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMei.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Me = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMei.AddCell(sCell2Me);
			foreach (var item in dataMay)
			{
				var checkHMay = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMay.Count() > 0)
				{
					foreach (var itemHMay in resultHolidays)
					{
						if (itemHMay.Date == item.Date)
						{
							var colorHMay = HexToColor(itemHMay.Color);
							string shift2Val = "";
							if (itemHMay.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMay);
							tableMei.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMei.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Me = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableMei.AddCell(sCell3Me);
			foreach (var item in dataMay)
			{
				var checkHMay = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHMay.Count() > 0)
				{
					foreach (var itemHMay in resultHolidays)
					{
						if (itemHMay.Date == item.Date)
						{
							var colorHMay = HexToColor(itemHMay.Color);
							string shift3Val = "";
							if (itemHMay.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHMay);
							tableMei.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableMei.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableMei);

			#endregion

			#region Juni

			string monthNameJuni = getMonth(new DateTime(model.Year, 6, 1));
			int daysJun = DateTime.DaysInMonth(model.Year, 6);
			int colJun = 0;

			if (daysJun == 30)
			{
				colJun = 31;
			}
			if (daysJun == 31)
			{
				colJun = 32;
			}

			PdfPTable tableJun = new PdfPTable(colJun);
			tableJun.WidthPercentage = 100;
			tableJun.HorizontalAlignment = 0;
			tableJun.SpacingBefore = 15f;
			tableJun.SpacingAfter = 10f;

			if (daysJun == 31)
			{
				tableJun.SetWidths(tblWidth31);
			}
			if (daysJun == 30)
			{
				tableJun.SetWidths(tblWidth30);
			}

			PdfPCell mCellJu = new PdfPCell(new Phrase(monthNameJuni, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellJu.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableJun.AddCell(mCellJu);
			foreach (var item in dataJun)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableJun.AddCell(wCell);
				}
				else
				{
					if (dataJun.IndexOf(item) == dataJun.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJun.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableJun.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJun.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableJun.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellJu = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJun.AddCell(dCellJu);
			foreach (var item in dataJun)
			{
				var checkHJun = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJun.Count() > 0)
				{
					foreach (var itemHJun in resultHolidays)
					{
						if (itemHJun.Date == item.Date)
						{
							var colorHJun = HexToColor(itemHJun.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHJun);
							tableJun.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJun.AddCell(dyCell);
				}

			}

			PdfPCell dtCellJu = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJun.AddCell(dtCellJu);
			foreach (var item in dataJun)
			{
				var checkHJun = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJun.Count() > 0)
				{
					foreach (var itemHJun in resultHolidays)
					{
						if (itemHJun.Date == item.Date)
						{
							var colorHJun = HexToColor(itemHJun.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHJun);
							tableJun.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJun.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Ju = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJun.AddCell(sCell1Ju);
			foreach (var item in dataJun)
			{
				var checkHJun = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJun.Count() > 0)
				{
					foreach (var itemHJun in resultHolidays)
					{
						if (itemHJun.Date == item.Date)
						{
							var colorHJun = HexToColor(itemHJun.Color);
							string shift1Val = "";
							if (itemHJun.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJun);
							tableJun.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJun.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Ju = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJun.AddCell(sCell2Ju);
			foreach (var item in dataJun)
			{
				var checkHJun = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJun.Count() > 0)
				{
					foreach (var itemHJun in resultHolidays)
					{
						if (itemHJun.Date == item.Date)
						{
							var colorHJun = HexToColor(itemHJun.Color);
							string shift2Val = "";
							if (itemHJun.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJun);
							tableJun.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJun.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Ju = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJun.AddCell(sCell3Ju);
			foreach (var item in dataJun)
			{
				var checkHJun = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJun.Count() > 0)
				{
					foreach (var itemHJun in resultHolidays)
					{
						if (itemHJun.Date == item.Date)
						{
							var colorHJun = HexToColor(itemHJun.Color);
							string shift3Val = "";
							if (itemHJun.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJun);
							tableJun.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJun.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableJun);

			#endregion

			#region Juli

			string monthNameJuli = getMonth(new DateTime(model.Year, 7, 1));
			int daysJul = DateTime.DaysInMonth(model.Year, 7);
			int colJul = 0;

			if (daysJul == 30)
			{
				colJul = 31;
			}
			if (daysJul == 31)
			{
				colJul = 32;
			}

			PdfPTable tableJul = new PdfPTable(colJul);
			tableJul.WidthPercentage = 100;
			tableJul.HorizontalAlignment = 0;
			tableJul.SpacingBefore = 15f;
			tableJul.SpacingAfter = 10f;

			if (daysJul == 31)
			{
				tableJul.SetWidths(tblWidth31);
			}
			if (daysJul == 30)
			{
				tableJul.SetWidths(tblWidth30);
			}

			PdfPCell mCellJul = new PdfPCell(new Phrase(monthNameJuli, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellJul.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableJul.AddCell(mCellJul);
			foreach (var item in dataJul)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableJul.AddCell(wCell);
				}
				else
				{
					if (dataJul.IndexOf(item) == dataJul.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJul.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableJul.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableJul.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableJul.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellJul = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJul.AddCell(dCellJul);
			foreach (var item in dataJul)
			{
				var checkHJul = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJul.Count() > 0)
				{
					foreach (var itemHJul in resultHolidays)
					{
						if (itemHJul.Date == item.Date)
						{
							var colorHJul = HexToColor(itemHJul.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHJul);
							tableJul.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJul.AddCell(dyCell);
				}

			}

			PdfPCell dtCellJul = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJul.AddCell(dtCellJul);
			foreach (var item in dataJul)
			{
				var checkHJul = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJul.Count() > 0)
				{
					foreach (var itemHJul in resultHolidays)
					{
						if (itemHJul.Date == item.Date)
						{
							var colorHJul = HexToColor(itemHJul.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHJul);
							tableJul.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableJul.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Jul = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJul.AddCell(sCell1Jul);
			foreach (var item in dataJul)
			{
				var checkHJul = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJul.Count() > 0)
				{
					foreach (var itemHJul in resultHolidays)
					{
						if (itemHJul.Date == item.Date)
						{
							var colorHJul = HexToColor(itemHJul.Color);
							string shift1Val = "";
							if (itemHJul.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJul);
							tableJul.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJul.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Jul = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJul.AddCell(sCell2Jul);
			foreach (var item in dataJul)
			{
				var checkHJul = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJul.Count() > 0)
				{
					foreach (var itemHJul in resultHolidays)
					{
						if (itemHJul.Date == item.Date)
						{
							var colorHJul = HexToColor(itemHJul.Color);
							string shift2Val = "";
							if (itemHJul.Color == "#3366cc")
							{
								shift2Val = item.Shift1;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJul);
							tableJul.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJul.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Jul = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableJul.AddCell(sCell3Jul);
			foreach (var item in dataJul)
			{
				var checkHJul = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHJul.Count() > 0)
				{
					foreach (var itemHJul in resultHolidays)
					{
						if (itemHJul.Date == item.Date)
						{
							var colorHJul = HexToColor(itemHJul.Color);
							string shift3Val = "";
							if (itemHJul.Color == "#3366cc")
							{
								shift3Val = item.Shift1;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHJul);
							tableJul.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableJul.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableJul);

			#endregion

			#region Agustus

			string monthNameAug = getMonth(new DateTime(model.Year, 8, 1));
			int daysAug = DateTime.DaysInMonth(model.Year, 7);
			int colAug = 0;

			if (daysAug == 30)
			{
				colAug = 31;
			}
			if (daysAug == 31)
			{
				colAug = 32;
			}

			PdfPTable tableAug = new PdfPTable(colAug);
			tableAug.WidthPercentage = 100;
			tableAug.HorizontalAlignment = 0;
			tableAug.SpacingBefore = 15f;
			tableAug.SpacingAfter = 10f;

			if (daysAug == 31)
			{
				tableAug.SetWidths(tblWidth31);
			}
			if (daysAug == 30)
			{
				tableAug.SetWidths(tblWidth30);
			}

			PdfPCell mCellAug = new PdfPCell(new Phrase(monthNameAug, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellAug.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableAug.AddCell(mCellAug);
			foreach (var item in dataAug)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableAug.AddCell(wCell);
				}
				else
				{
					if (dataAug.IndexOf(item) == dataAug.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableAug.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableAug.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableAug.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableAug.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellAug = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableAug.AddCell(dCellAug);
			foreach (var item in dataAug)
			{
				var checkHAug = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHAug.Count() > 0)
				{
					foreach (var itemHAug in resultHolidays)
					{
						if (itemHAug.Date == item.Date)
						{
							var colorHAug = HexToColor(itemHAug.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHAug);
							tableAug.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableAug.AddCell(dyCell);
				}

			}

			PdfPCell dtCellAug = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableAug.AddCell(dtCellAug);
			foreach (var item in dataAug)
			{
				var checkHAug = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHAug.Count() > 0)
				{
					foreach (var itemHAug in resultHolidays)
					{
						if (itemHAug.Date == item.Date)
						{
							var colorHAug = HexToColor(itemHAug.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHAug);
							tableAug.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableAug.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Aug = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableAug.AddCell(sCell1Aug);
			foreach (var item in dataAug)
			{
				var checkHAug = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHAug.Count() > 0)
				{
					foreach (var itemHAug in resultHolidays)
					{
						if (itemHAug.Date == item.Date)
						{
							var colorHAug = HexToColor(itemHAug.Color);
							string shift1Val = "";
							if (itemHAug.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHAug);
							tableAug.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableAug.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Aug = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableAug.AddCell(sCell2Aug);
			foreach (var item in dataAug)
			{
				var checkHAug = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHAug.Count() > 0)
				{
					foreach (var itemHAug in resultHolidays)
					{
						if (itemHAug.Date == item.Date)
						{
							var colorHAug = HexToColor(itemHAug.Color);
							string shift2Val = "";
							if (itemHAug.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHAug);
							tableAug.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableAug.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Aug = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableAug.AddCell(sCell3Aug);
			foreach (var item in dataAug)
			{
				var checkHAug = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHAug.Count() > 0)
				{
					foreach (var itemHAug in resultHolidays)
					{
						if (itemHAug.Date == item.Date)
						{
							var colorHAug = HexToColor(itemHAug.Color);
							string shift3Val = "";
							if (itemHAug.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHAug);
							tableAug.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableAug.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableAug);

			#endregion

			pdfDoc.NewPage();

			#region September

			string monthNameSept = getMonth(new DateTime(model.Year, 9, 1));
			int daysSept = DateTime.DaysInMonth(model.Year, 9);
			int colSept = 0;

			if (daysSept == 30)
			{
				colSept = 31;
			}
			if (daysSept == 31)
			{
				colSept = 32;
			}

			PdfPTable tableSept = new PdfPTable(colSept);
			tableSept.WidthPercentage = 100;
			tableSept.HorizontalAlignment = 0;
			tableSept.SpacingBefore = 20f;
			tableSept.SpacingAfter = 10f;

			if (daysSept == 31)
			{
				tableSept.SetWidths(tblWidth31);
			}
			if (daysSept == 30)
			{
				tableSept.SetWidths(tblWidth30);
			}

			PdfPCell mCellSept = new PdfPCell(new Phrase(monthNameSept, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellSept.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableSept.AddCell(mCellSept);
			foreach (var item in dataSept)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableSept.AddCell(wCell);
				}
				else
				{
					if (dataSept.IndexOf(item) == dataSept.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableSept.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableSept.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableSept.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableSept.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellSept = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableSept.AddCell(dCellSept);
			foreach (var item in dataSept)
			{
				var checkHSept = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHSept.Count() > 0)
				{
					foreach (var itemHSept in resultHolidays)
					{
						if (itemHSept.Date == item.Date)
						{
							var colorHSept = HexToColor(itemHSept.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHSept);
							tableSept.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableSept.AddCell(dyCell);
				}

			}

			PdfPCell dtCellSept = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableSept.AddCell(dtCellSept);
			foreach (var item in dataSept)
			{
				var checkHSept = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHSept.Count() > 0)
				{
					foreach (var itemHSept in resultHolidays)
					{
						if (itemHSept.Date == item.Date)
						{
							var colorHSept = HexToColor(itemHSept.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHSept);
							tableSept.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableSept.AddCell(dteCell);
				}

			}


			PdfPCell sCell1Sept = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableSept.AddCell(sCell1Sept);
			foreach (var item in dataSept)
			{
				var checkHSept = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHSept.Count() > 0)
				{
					foreach (var itemHSept in resultHolidays)
					{
						if (itemHSept.Date == item.Date)
						{
							var colorHSept = HexToColor(itemHSept.Color);
							string shift1Val = "";
							if (itemHSept.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHSept);
							tableSept.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableSept.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Sept = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableSept.AddCell(sCell2Sept);
			foreach (var item in dataSept)
			{
				var checkHSept = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHSept.Count() > 0)
				{
					foreach (var itemHSept in resultHolidays)
					{
						if (itemHSept.Date == item.Date)
						{
							var colorHSept = HexToColor(itemHSept.Color);
							string shift2Val = "";
							if (itemHSept.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHSept);
							tableSept.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableSept.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Sept = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableSept.AddCell(sCell3Sept);
			foreach (var item in dataSept)
			{
				var checkHSept = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHSept.Count() > 0)
				{
					foreach (var itemHSept in resultHolidays)
					{
						if (itemHSept.Date == item.Date)
						{
							var colorHSept = HexToColor(itemHSept.Color);
							string shift3Val = "";
							if (itemHSept.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHSept);
							tableSept.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableSept.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableSept);

			#endregion

			#region Oktober

			string monthNameOct = getMonth(new DateTime(model.Year, 10, 1));
			int daysOct = DateTime.DaysInMonth(model.Year, 10);
			int colOct = 0;

			if (daysOct == 30)
			{
				colOct = 31;
			}
			if (daysOct == 31)
			{
				colOct = 32;
			}

			PdfPTable tableOct = new PdfPTable(colOct);
			tableOct.WidthPercentage = 100;
			tableOct.HorizontalAlignment = 0;
			tableOct.SpacingBefore = 15f;
			tableOct.SpacingAfter = 10f;

			if (daysOct == 31)
			{
				tableOct.SetWidths(tblWidth31);
			}
			if (daysOct == 30)
			{
				tableOct.SetWidths(tblWidth30);
			}

			PdfPCell mCellOct = new PdfPCell(new Phrase(monthNameOct, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellOct.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableOct.AddCell(mCellOct);
			foreach (var item in dataOct)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableOct.AddCell(wCell);
				}
				else
				{
					if (dataOct.IndexOf(item) == dataOct.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableOct.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableOct.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableOct.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableOct.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellOct = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableOct.AddCell(dCellOct);
			foreach (var item in dataOct)
			{
				var checkHOct = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHOct.Count() > 0)
				{
					foreach (var itemHOct in resultHolidays)
					{
						if (itemHOct.Date == item.Date)
						{
							var colorHOct = HexToColor(itemHOct.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHOct);
							tableOct.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableOct.AddCell(dyCell);
				}
			}

			PdfPCell dtCellOct = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableOct.AddCell(dtCellOct);
			foreach (var item in dataOct)
			{
				var checkHOct = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHOct.Count() > 0)
				{
					foreach (var itemHOct in resultHolidays)
					{
						if (itemHOct.Date == item.Date)
						{
							var colorHOct = HexToColor(itemHOct.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHOct);
							tableOct.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableOct.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Oct = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableOct.AddCell(sCell1Oct);
			foreach (var item in dataOct)
			{
				var checkHOct = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHOct.Count() > 0)
				{
					foreach (var itemHOct in resultHolidays)
					{
						if (itemHOct.Date == item.Date)
						{
							var colorHOct = HexToColor(itemHOct.Color);
							string shift1Val = "";
							if (itemHOct.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHOct);
							tableOct.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableOct.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Oct = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableOct.AddCell(sCell2Oct);
			foreach (var item in dataOct)
			{
				var checkHOct = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHOct.Count() > 0)
				{
					foreach (var itemHOct in resultHolidays)
					{
						if (itemHOct.Date == item.Date)
						{
							var colorHOct = HexToColor(itemHOct.Color);
							string shift2Val = "";
							if (itemHOct.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHOct);
							tableOct.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableOct.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Oct = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableOct.AddCell(sCell3Oct);
			foreach (var item in dataOct)
			{
				var checkHOct = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHOct.Count() > 0)
				{
					foreach (var itemHOct in resultHolidays)
					{
						if (itemHOct.Date == item.Date)
						{
							var colorHOct = HexToColor(itemHOct.Color);
							string shift3Val = "";
							if (itemHOct.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHOct);
							tableOct.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableOct.AddCell(cellColJ);
				}


			}

			pdfDoc.Add(tableOct);

			#endregion

			#region November

			string monthNameNov = getMonth(new DateTime(model.Year, 11, 1));
			int daysNov = DateTime.DaysInMonth(model.Year, 11);
			int colNov = 0;

			if (daysNov == 30)
			{
				colNov = 31;
			}
			if (daysNov == 31)
			{
				colNov = 32;
			}

			PdfPTable tableNov = new PdfPTable(colNov);
			tableNov.WidthPercentage = 100;
			tableNov.HorizontalAlignment = 0;
			tableNov.SpacingBefore = 15f;
			tableNov.SpacingAfter = 10f;

			if (daysNov == 31)
			{
				tableNov.SetWidths(tblWidth31);
			}
			if (daysNov == 30)
			{
				tableNov.SetWidths(tblWidth30);
			}

			PdfPCell mCellNov = new PdfPCell(new Phrase(monthNameNov, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellNov.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableNov.AddCell(mCellNov);
			foreach (var item in dataNov)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableNov.AddCell(wCell);
				}
				else
				{
					if (dataNov.IndexOf(item) == dataNov.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableNov.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableNov.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableNov.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableNov.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellNov = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableNov.AddCell(dCellNov);
			foreach (var item in dataNov)
			{
				var checkHNov = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHNov.Count() > 0)
				{
					foreach (var itemHNov in resultHolidays)
					{
						if (itemHNov.Date == item.Date)
						{
							var colorHNov = HexToColor(itemHNov.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHNov);
							tableNov.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableNov.AddCell(dyCell);
				}

			}

			PdfPCell dtCellNov = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableNov.AddCell(dtCellNov);
			foreach (var item in dataNov)
			{
				var checkHNov = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHNov.Count() > 0)
				{
					foreach (var itemHNov in resultHolidays)
					{
						if (itemHNov.Date == item.Date)
						{
							var colorHNov = HexToColor(itemHNov.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHNov);
							tableNov.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableNov.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Nov = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableNov.AddCell(sCell1Nov);
			foreach (var item in dataNov)
			{
				var checkHNov = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHNov.Count() > 0)
				{
					foreach (var itemHNov in resultHolidays)
					{
						if (itemHNov.Date == item.Date)
						{
							var colorHNov = HexToColor(itemHNov.Color);
							string shift1Val = "";
							if (itemHNov.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHNov);
							tableNov.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableNov.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Nov = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableNov.AddCell(sCell2Jul);
			foreach (var item in dataNov)
			{
				var checkHNov = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHNov.Count() > 0)
				{
					foreach (var itemHNov in resultHolidays)
					{
						if (itemHNov.Date == item.Date)
						{
							var colorHNov = HexToColor(itemHNov.Color);
							string shift2Val = "";
							if (itemHNov.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHNov);
							tableNov.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableNov.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Nov = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableNov.AddCell(sCell3Nov);
			foreach (var item in dataNov)
			{
				var checkHNov = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHNov.Count() > 0)
				{
					foreach (var itemHNov in resultHolidays)
					{
						if (itemHNov.Date == item.Date)
						{
							var colorHNov = HexToColor(itemHNov.Color);
							string shift3Val = "";
							if (itemHNov.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHNov);
							tableNov.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableNov.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableNov);

			#endregion

			#region Desember

			string monthNameDes = getMonth(new DateTime(model.Year, 12, 1));
			int daysDes = DateTime.DaysInMonth(model.Year, 12);
			int colDes = 0;

			if (daysDes == 30)
			{
				colDes = 31;
			}
			if (daysDes == 31)
			{
				colDes = 32;
			}

			PdfPTable tableDes = new PdfPTable(colDes);
			tableDes.WidthPercentage = 100;
			tableDes.HorizontalAlignment = 0;
			tableDes.SpacingBefore = 15f;
			tableDes.SpacingAfter = 10f;

			if (daysDes == 31)
			{
				tableDes.SetWidths(tblWidth31);
			}
			if (daysDes == 30)
			{
				tableDes.SetWidths(tblWidth30);
			}

			PdfPCell mCellDes = new PdfPCell(new Phrase(monthNameDes, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			mCellDes.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
			tableDes.AddCell(mCellDes);
			foreach (var item in dataDec)
			{
				var dayCheck = getDay(item.Date);
				var week = getCurrentWeekNumber(item.Date);
				var dayName = getDay(item.Date);
				if (dayName == "Thu")
				{
					PdfPCell wCell = new PdfPCell(new Phrase("W " + week, fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
					wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;
					tableDes.AddCell(wCell);
				}
				else
				{
					if (dataDec.IndexOf(item) == dataDec.Count - 1)
					{
						PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						if (dayName == "Sun")
						{

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableDes.AddCell(wCell);
						}

						wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
						wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

						tableDes.AddCell(wCell);
					}
					else
					{
						if (dayName == "Sun")
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER;

							tableDes.AddCell(wCell);
						}
						else
						{
							PdfPCell wCell = new PdfPCell(new Phrase("", fontWk)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

							wCell.BackgroundColor = new iTextSharp.text.Color(120, 238, 22);
							wCell.Border = Rectangle.NO_BORDER | Rectangle.TOP_BORDER;

							tableDes.AddCell(wCell);
						}

					}
				}


				weekTemp = week;

			}

			PdfPCell dCellDes = new PdfPCell(new Phrase("Day", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableDes.AddCell(dCellDes);
			foreach (var item in dataDec)
			{
				var checkHDes = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHDes.Count() > 0)
				{
					foreach (var itemHDes in resultHolidays)
					{
						if (itemHDes.Date == item.Date)
						{
							var colorHDes = HexToColor(itemHDes.Color);
							var dayName = getDay(item.Date);
							PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dyCell.BackgroundColor = new iTextSharp.text.Color(colorHDes);
							tableDes.AddCell(dyCell);

							break;
						}
					}
				}
				else
				{
					var dayName = getDay(item.Date);
					PdfPCell dyCell = new PdfPCell(new Phrase(dayName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableDes.AddCell(dyCell);
				}

			}

			PdfPCell dtCellDes = new PdfPCell(new Phrase("Date", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableDes.AddCell(dtCellDes);
			foreach (var item in dataDec)
			{
				var checkHDes = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHDes.Count() > 0)
				{
					foreach (var itemHDes in resultHolidays)
					{
						if (itemHDes.Date == item.Date)
						{
							var colorHDes = HexToColor(itemHDes.Color);

							PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							dteCell.BackgroundColor = new iTextSharp.text.Color(colorHDes);
							tableDes.AddCell(dteCell);

							break;
						}
					}
				}
				else
				{
					PdfPCell dteCell = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					tableDes.AddCell(dteCell);
				}

			}

			PdfPCell sCell1Des = new PdfPCell(new Phrase("Shift 1", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableDes.AddCell(sCell1Des);
			foreach (var item in dataDec)
			{
				var checkHDes = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHDes.Count() > 0)
				{
					foreach (var itemHDes in resultHolidays)
					{
						if (itemHDes.Date == item.Date)
						{
							var colorHDes = HexToColor(itemHDes.Color);
							string shift1Val = "";
							if (itemHDes.Color == "#3366cc")
							{
								shift1Val = item.Shift1;
							}
							else
							{
								shift1Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift1Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHDes);
							tableDes.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift1, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift1 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift1 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift1 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift1 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableDes.AddCell(cellColJ);
				}

			}

			PdfPCell sCell2Des = new PdfPCell(new Phrase("Shift 2", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableDes.AddCell(sCell2Des);
			foreach (var item in dataDec)
			{
				var checkHDes = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHDes.Count() > 0)
				{
					foreach (var itemHDes in resultHolidays)
					{
						if (itemHDes.Date == item.Date)
						{
							var colorHDes = HexToColor(itemHDes.Color);
							string shift2Val = "";
							if (itemHDes.Color == "#3366cc")
							{
								shift2Val = item.Shift2;
							}
							else
							{
								shift2Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift2Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHDes);
							tableDes.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift2, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift2 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift2 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift2 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift2 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableDes.AddCell(cellColJ);
				}

			}

			PdfPCell sCell3Des = new PdfPCell(new Phrase("Shift 3", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			tableDes.AddCell(sCell3Des);
			foreach (var item in dataDec)
			{
				var checkHDes = resultHolidays.Where(x => x.Date.ToString("dd-MMM-yy").ToLower().Contains(item.Date.ToString("dd-MMM-yy").ToLower()));
				if (checkHDes.Count() > 0)
				{
					foreach (var itemHDes in resultHolidays)
					{
						if (itemHDes.Date == item.Date)
						{
							var colorHDes = HexToColor(itemHDes.Color);
							string shift3Val = "";
							if (itemHDes.Color == "#3366cc")
							{
								shift3Val = item.Shift3;
							}
							else
							{
								shift3Val = "";
							}
							PdfPCell cellColJ = new PdfPCell(new Phrase(shift3Val, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
							cellColJ.BackgroundColor = new iTextSharp.text.Color(colorHDes);
							tableDes.AddCell(cellColJ);
							break;
						}
					}
				}
				else
				{
					PdfPCell cellColJ = new PdfPCell(new Phrase(item.Shift3, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
					if (item.Shift3 == "A")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
					}
					if (item.Shift3 == "B")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
					}
					if (item.Shift3 == "C")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
					}
					if (item.Shift3 == "D")
					{
						cellColJ.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
					}
					tableDes.AddCell(cellColJ);
				}

			}

			pdfDoc.Add(tableDes);

			#endregion

			#region Color
			PdfPTable tableColorBlock = new PdfPTable(15);
			tableColorBlock.DefaultCell.Border = Rectangle.NO_BORDER;
			tableColorBlock.WidthPercentage = 80;
			tableColorBlock.HorizontalAlignment = 0;
			tableColorBlock.SpacingBefore = 10f;
			tableColorBlock.SpacingAfter = 10f;

			int[] tblWidthColorBlock = { 1, 5, 1, 1, 5, 1, 1, 5, 1, 1, 5, 1, 1, 10, 1 };
			tableColorBlock.SetWidths(tblWidthColorBlock);

			List<SelectListItem> _menuList = new List<SelectListItem>();

			string reference = _referenceAppService.GetBy("Name", "HT", true);
			ReferenceModel refModel = reference.DeserializeToReference();
			if (refModel != null)
			{
				string dataList = _referenceAppService.FindDetailBy("ReferenceID", refModel.ID, true);
				List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();
				dataModelList = dataModelList.OrderBy(x => x.Description.Length).ToList();

				foreach (var item in dataModelList)
				{
					PdfPCell cellCol1 = new PdfPCell(new Phrase(" "));
					System.Drawing.Color colorcell = System.Drawing.ColorTranslator.FromHtml(item.Code);
					cellCol1.BackgroundColor = new iTextSharp.text.Color(colorcell);
					tableColorBlock.AddCell(cellCol1);
					PdfPCell dh1 = new PdfPCell(new Phrase(" " + item.Description, font));
					dh1.Border = Rectangle.NO_BORDER;
					tableColorBlock.AddCell(dh1);
					tableColorBlock.AddCell(" ");
				}
			}

			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");

			PdfPCell cellCol3 = new PdfPCell(new Phrase(" "));
			cellCol3.BackgroundColor = new iTextSharp.text.Color(255, 153, 204);
			tableColorBlock.AddCell(cellCol3);
			PdfPCell dh3 = new PdfPCell(new Phrase(" Group A", font));
			dh3.Border = Rectangle.NO_BORDER;
			tableColorBlock.AddCell(dh3);
			tableColorBlock.AddCell(" ");

			PdfPCell cellCol8 = new PdfPCell(new Phrase(" "));
			cellCol8.BackgroundColor = new iTextSharp.text.Color(255, 204, 153);
			tableColorBlock.AddCell(cellCol8);
			PdfPCell dh8 = new PdfPCell(new Phrase(" Group B", font));
			dh8.Border = Rectangle.NO_BORDER;
			tableColorBlock.AddCell(dh8);
			tableColorBlock.AddCell(" ");

			PdfPCell cellCol4 = new PdfPCell(new Phrase(" "));
			cellCol4.BackgroundColor = new iTextSharp.text.Color(255, 255, 153);
			tableColorBlock.AddCell(cellCol4);
			PdfPCell dh4 = new PdfPCell(new Phrase(" Group C", font));
			dh4.Border = Rectangle.NO_BORDER;
			tableColorBlock.AddCell(dh4);
			tableColorBlock.AddCell(" ");

			PdfPCell cellCol9 = new PdfPCell(new Phrase(" "));
			cellCol9.BackgroundColor = new iTextSharp.text.Color(226, 239, 219);
			tableColorBlock.AddCell(cellCol9);
			PdfPCell dh9 = new PdfPCell(new Phrase(" Group D", font));
			dh9.Border = Rectangle.NO_BORDER;
			tableColorBlock.AddCell(dh9);
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");
			tableColorBlock.AddCell(" ");

			pdfDoc.Add(tableColorBlock);
			#endregion

			pdfDoc.NewPage();

			#region Holiday 

			if (resultHolidays.Any(x => x.Date.Month <= 6))
			{
				PdfPTable table = new PdfPTable(3);

				int[] tblWidthH1 = { 2, 1, 7 };
				table.SetWidths(tblWidthH1);
				table.TotalWidth = 400f;
				table.LockedWidth = true;

				PdfPCell cellColH1 = new PdfPCell(new Phrase("Bulan", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH1.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH1);

				PdfPCell cellColH2 = new PdfPCell(new Phrase("Tgl", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH2.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH2);

				PdfPCell cellColH3 = new PdfPCell(new Phrase("Informasi", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH3.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH3);

				foreach (var item in resultHolidays)
				{
					if (item.Date.Month <= 6)
					{
						var monthH = getMonth(item.Date);
						var informationH = item.Description;
						var colorH = HexToColor(item.Color);

						PdfPCell cellColHI = new PdfPCell(new Phrase(monthH, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						PdfPCell cellColHD = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						PdfPCell cellColHIn = new PdfPCell(new Phrase(informationH, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

						cellColHI.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
						table.AddCell(cellColHI);
						cellColHD.BackgroundColor = new iTextSharp.text.Color(colorH);
						table.AddCell(cellColHD);
						cellColHIn.BackgroundColor = new iTextSharp.text.Color(colorH);
						table.AddCell(cellColHIn);
					}
				}
				table.WriteSelectedRows(0, -1, pdfDoc.Left, pdfDoc.Top, pdfWriter.DirectContent);
			}

			if (resultHolidays.Any(x => x.Date.Month > 6))
			{
				PdfPTable table = new PdfPTable(3);
				int[] tblWidthH2 = { 2, 1, 7 };
				table.SetWidths(tblWidthH2);
				table.TotalWidth = 350f;
				table.LockedWidth = true;

				PdfPCell cellColH4 = new PdfPCell(new Phrase("Bulan", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH4.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH4);

				PdfPCell cellColH5 = new PdfPCell(new Phrase("Tgl", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH5.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH5);

				PdfPCell cellColH6 = new PdfPCell(new Phrase("Informasi", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				cellColH6.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
				table.AddCell(cellColH6);

				foreach (var item in resultHolidays)
				{
					if (item.Date.Month > 6)
					{
						var monthH = getMonth(item.Date);
						var informationH = item.Description;
						var colorH = HexToColor(item.Color);

						PdfPCell cellColHI = new PdfPCell(new Phrase(monthH, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						PdfPCell cellColHD = new PdfPCell(new Phrase(item.Date.Day.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
						PdfPCell cellColHIn = new PdfPCell(new Phrase(informationH, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

						cellColHI.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
						table.AddCell(cellColHI);
						cellColHD.BackgroundColor = new iTextSharp.text.Color(colorH);
						table.AddCell(cellColHD);
						cellColHIn.BackgroundColor = new iTextSharp.text.Color(colorH);
						table.AddCell(cellColHIn);
					}
				}

				table.WriteSelectedRows(0, -1, pdfDoc.Left + 450, pdfDoc.Top, pdfWriter.DirectContent);
			}
			#endregion

			#region Working Summary 

			PdfPTable tableWS = new PdfPTable(7);
			//int[] tableWidthWS = { 2, 1, 7 };
			//tableWS.SetWidths(tableWidthWS);
			tableWS.TotalWidth = 400f;
			tableWS.LockedWidth = true;

			PdfPCell cellColWSH1 = new PdfPCell(new Phrase("Month", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH1.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH1);

			PdfPCell cellColWSH2 = new PdfPCell(new Phrase("Days", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH2.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH2);

			PdfPCell cellColWSH3 = new PdfPCell(new Phrase("Holiday", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH3.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH3);

			PdfPCell cellColWSH4 = new PdfPCell(new Phrase("Leave", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH4.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH4);

			PdfPCell cellColWSH5 = new PdfPCell(new Phrase("Prod Off", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH5.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH5);

			PdfPCell cellColWSH6 = new PdfPCell(new Phrase("Shift Off", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH6.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH6);

			PdfPCell cellColWSH7 = new PdfPCell(new Phrase("Work Days", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColWSH7.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableWS.AddCell(cellColWSH7);

			string holidayType = _referenceAppService.GetDetailAll(ReferenceEnum.HolidayType, true);
			List<ReferenceDetailModel> htModelList = holidayType.DeserializeToRefDetailList();

			ReferenceDetailModel htholiday = htModelList.Where(x => x.Description.ToLower().Contains("libur nasional")).FirstOrDefault();
			ReferenceDetailModel htleaveHMS = htModelList.Where(x => x.Description.ToLower().Contains("cuti bersama hms")).FirstOrDefault();
			ReferenceDetailModel htleavePemerintah = htModelList.Where(x => x.Description.ToLower().Contains("cuti bersama pemerintah")).FirstOrDefault();
			ReferenceDetailModel htshift = htModelList.Where(x => x.Description.ToLower().Contains("shift off")).FirstOrDefault();
			ReferenceDetailModel htprod = htModelList.Where(x => x.Description.ToLower().Contains("acara internal")).FirstOrDefault();
			List<CalendarWorkingSummaryModel> result1 = new List<CalendarWorkingSummaryModel>();

			long idHoliday = htholiday == null ? 0 : htholiday.ID;
			long idLeaveHMS = htleaveHMS == null ? 0 : htleaveHMS.ID;
			long idLeavePemerintah = htleavePemerintah == null ? 0 : htleavePemerintah.ID;
			long idProdOff = htprod == null ? 0 : htprod.ID;
			long idShiftOff = htshift == null ? 0 : htshift.ID;

			string[] columnNames = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

			// Get January Data
			int totalDay = 0;
			int totalHoliday = 0;
			int totalLeave = 0;
			int totalProdOff = 0;
			int totalShiftOff = 0;
			for (int i = 0; i < 13; i++)
			{
				CalendarWorkingSummaryModel newModel = new CalendarWorkingSummaryModel();
				newModel.ColumnName = columnNames[i];
				if (i < 12)
				{
					var monthlyCalendarHoliday = resultHolidaysAll.Where(x => x.Date.Month == (i + 1)).ToList();
					newModel.Holiday = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idHoliday).Count();
					newModel.Leaves = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idLeaveHMS || x.HolidayTypeID == idLeavePemerintah).Count();
					newModel.ProdOff = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idProdOff).Count();
					newModel.ShiftOff = monthlyCalendarHoliday.Where(x => x.HolidayTypeID == idShiftOff).Count();
					newModel.Days = calendarData.Where(x => x.Date.Month == (i + 1)).Count();
					newModel.WorkDays = newModel.Days - (newModel.Holiday + newModel.Leaves + newModel.ProdOff + newModel.ShiftOff);

					totalDay += newModel.Days;
					totalHoliday += newModel.Holiday;
					totalLeave += newModel.Leaves;
					totalProdOff += newModel.ProdOff;
					totalShiftOff += newModel.ShiftOff;
				}
				else
				{
					newModel.Days = totalDay;
					newModel.Holiday = totalHoliday;
					newModel.Leaves = totalLeave;
					newModel.ProdOff = totalProdOff;
					newModel.ShiftOff = totalShiftOff;
					newModel.WorkDays = totalDay - (totalHoliday + totalLeave + totalProdOff + totalShiftOff);
				}

				result1.Add(newModel);
			}

			foreach (var item in result1)
			{
				PdfPCell cellColWS1 = new PdfPCell(new Phrase(item.ColumnName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS11 = new PdfPCell(new Phrase(item.Days.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS2 = new PdfPCell(new Phrase(item.Holiday.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS3 = new PdfPCell(new Phrase(item.Leaves.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS4 = new PdfPCell(new Phrase(item.ProdOff.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS5 = new PdfPCell(new Phrase(item.ShiftOff.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColWS6 = new PdfPCell(new Phrase(item.WorkDays.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

				tableWS.AddCell(cellColWS1);
				tableWS.AddCell(cellColWS11);
				tableWS.AddCell(cellColWS2);
				tableWS.AddCell(cellColWS3);
				tableWS.AddCell(cellColWS4);
				tableWS.AddCell(cellColWS5);
				tableWS.AddCell(cellColWS6);
			}

			tableWS.WriteSelectedRows(0, -1, pdfDoc.Left, pdfDoc.Top - 300, pdfWriter.DirectContent);

			PdfPTable tableGWS = new PdfPTable(5);
			//int[] tblWidthGWS = { 2, 1, 7 };
			//tableGWS.SetWidths(tblWidthGWS);
			tableGWS.TotalWidth = 350f;
			tableGWS.LockedWidth = true;

			PdfPCell cellColGWSH1 = new PdfPCell(new Phrase("Month", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColGWSH1.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableGWS.AddCell(cellColGWSH1);

			PdfPCell cellColGWSH2 = new PdfPCell(new Phrase("A", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColGWSH2.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableGWS.AddCell(cellColGWSH2);

			PdfPCell cellColGWSH3 = new PdfPCell(new Phrase("B", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColGWSH3.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableGWS.AddCell(cellColGWSH3);

			PdfPCell cellColGWSH4 = new PdfPCell(new Phrase("C", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColGWSH4.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableGWS.AddCell(cellColGWSH4);

			PdfPCell cellColGWSH5 = new PdfPCell(new Phrase("D", font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
			cellColGWSH5.BackgroundColor = new iTextSharp.text.Color(255, 87, 51);
			tableGWS.AddCell(cellColGWSH5);

			List<CalendarWorkingSummaryGroupModel> result = new List<CalendarWorkingSummaryGroupModel>();

			int totalA = 0, totalB = 0, totalC = 0, totalD = 0;

			// Get January Data
			for (int i = 0; i < 13; i++)
			{
				CalendarWorkingSummaryGroupModel newModel = new CalendarWorkingSummaryGroupModel();
				newModel.ColumnName = columnNames[i];

				// get A amount
				if (i < 12)
				{
					int month = i + 1;
					newModel.A = calendarData.Where(x => x.Date.Month == month && (x.Shift1 == "A" || x.Shift2 == "A" || x.Shift3 == "A")).Count();
					totalA += newModel.A;
					// get B amount
					newModel.B = calendarData.Where(x => x.Date.Month == month && (x.Shift1 == "B" || x.Shift2 == "B" || x.Shift3 == "B")).Count();
					totalB += newModel.B;
					// get C amount
					newModel.C = calendarData.Where(x => x.Date.Month == month && (x.Shift1 == "C" || x.Shift2 == "C" || x.Shift3 == "C")).Count();
					totalC += newModel.C;
					// get D amount
					newModel.D = calendarData.Where(x => x.Date.Month == month && (x.Shift1 == "D" || x.Shift2 == "D" || x.Shift3 == "D")).Count();
					totalD += newModel.D;
				}
				else
				{
					newModel.A = totalA;
					newModel.B = totalB;
					newModel.C = totalC;
					newModel.D = totalD;
				}

				result.Add(newModel);
			}

			foreach (var item in result)
			{
				PdfPCell cellColGWS1 = new PdfPCell(new Phrase(item.ColumnName, font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColGWS2 = new PdfPCell(new Phrase(item.A.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColGWS3 = new PdfPCell(new Phrase(item.B.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColGWS4 = new PdfPCell(new Phrase(item.C.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };
				PdfPCell cellColGWS5 = new PdfPCell(new Phrase(item.D.ToString(), font)) { HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER };

				tableGWS.AddCell(cellColGWS1);
				tableGWS.AddCell(cellColGWS2);
				tableGWS.AddCell(cellColGWS3);
				tableGWS.AddCell(cellColGWS4);
				tableGWS.AddCell(cellColGWS5);
			}

			tableGWS.WriteSelectedRows(0, -1, pdfDoc.Left + 450, pdfDoc.Top - 300, pdfWriter.DirectContent);

			#endregion

			pdfWriter.CloseStream = false;
			pdfDoc.Close();
			Response.Buffer = true;
			Response.ContentType = "application/pdf";
			Response.AddHeader("content-disposition", "attachment;filename=Report.pdf");
			Response.Cache.SetCacheability(HttpCacheability.NoCache);
			Response.Write(pdfDoc);
			Response.End();

			return View();
		}

		#region Helper
		public string getMonth(DateTime myDate)
		{
			DateTime mtime = myDate;
			string resMonth = mtime.ToString("MMMM");
			return resMonth;
		}
		public string getDate(DateTime myDates)
		{
			DateTime mtime = myDates;
			string resDate = mtime.ToString("dd");
			return resDate;
		}

		public string getDay(DateTime myDate)
		{
			DateTime mTime = myDate;
			string resDay = mTime.ToString("ddd");
			return resDay;
		}

		public string getCurrentWeekNumber(DateTime dateTime)
		{
			var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dateTime, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
			return weeknum.ToString();/*dateTime.ToString($"{weeknum}")*/
		}

		public static System.Drawing.Color HexToColor(string hexString)
		{
			//replace # occurences
			if (hexString.IndexOf('#') != -1)
				hexString = hexString.Replace("#", "");

			int r, g, b = 0;

			r = int.Parse(hexString.Substring(0, 2), NumberStyles.AllowHexSpecifier);
			g = int.Parse(hexString.Substring(2, 2), NumberStyles.AllowHexSpecifier);
			b = int.Parse(hexString.Substring(4, 2), NumberStyles.AllowHexSpecifier);

			return System.Drawing.Color.FromArgb(r, g, b);
		}
		#endregion

	}
}

