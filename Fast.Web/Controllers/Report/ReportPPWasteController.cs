using Fast.Application.Interfaces;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Fast.Web.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Web.Mvc;


namespace Fast.Web.Controllers.Report
{
	[CustomAuthorize("ReportPPWaste")]
	public class ReportPPWasteController : BaseController<PPLPHModel>
	{
		private readonly IPPLPHAppService _ppLphAppService;
		private readonly IPPLPHApprovalsAppService _ppLphApprovalAppService;
		private readonly IPPLPHComponentsAppService _ppLphComponentsAppService;
		private readonly IPPLPHLocationsAppService _ppLphLocationsAppService;
		private readonly IPPLPHValuesAppService _ppLphValuesAppService;
		private readonly IPPLPHValueHistoriesAppService _ppLphValueHistoriesAppService;
		private readonly IPPLPHExtrasAppService _ppLphExtrasAppService;
		private readonly IPPLPHSubmissionsAppService _ppLphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		public ReportPPWasteController(
		  IPPLPHAppService ppLPHAppService,
		  IPPLPHComponentsAppService ppLPHComponentsAppService,
		  IPPLPHLocationsAppService ppLPHLocationsAppService,
		  IPPLPHValuesAppService ppLPHValuesAppService,
		  IPPLPHApprovalsAppService ppLPHApprovalsAppService,
		  IPPLPHValueHistoriesAppService ppLPHValueHistoriesAppService,
		  IPPLPHExtrasAppService ppLPHExtrasAppService,
		  IPPLPHSubmissionsAppService ppLPHSubmissionsAppService,
		  ILoggerAppService logger,
		  IReferenceAppService referenceAppService,
		  ILocationAppService locationAppService,
		  IEmployeeAppService employeeAppService,
		  IWeeksAppService weeksAppService,
		  IUserAppService userAppService,
		  IUserRoleAppService userRoleAppService,
		  IMachineAppService machineAppService,
		  IReferenceDetailAppService referenceDetailAppService)
		{
			_ppLphAppService = ppLPHAppService;
			_ppLphComponentsAppService = ppLPHComponentsAppService;
			_ppLphLocationsAppService = ppLPHLocationsAppService;
			_ppLphValuesAppService = ppLPHValuesAppService;
			_ppLphApprovalAppService = ppLPHApprovalsAppService;
			_ppLphValueHistoriesAppService = ppLPHValueHistoriesAppService;
			_ppLphExtrasAppService = ppLPHExtrasAppService;
			_ppLphSubmissionsAppService = ppLPHSubmissionsAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_weeksAppService = weeksAppService;
			_userAppService = userAppService;
			_userRoleAppService = userRoleAppService;
			_machineAppService = machineAppService;
			_referenceDetailAppService = referenceDetailAppService;
		}

		public ActionResult Index()
		{
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			return View();
        }

        [HttpPost]
        public ActionResult SetDashboard(List<ReportWasteModel> waste, string location, string type)
        {
            try
            {
                type = type.Replace("LPHPrimary", "").Replace("Controller", "");
                string modifiedDate = DateTime.Now.ToString("yyyy-MM-dd");
                string query = "";
                foreach(int year in waste.Select(x => x.Year).Distinct())
                {
                    List<ReportWasteModel> wasteThisYear = waste.Where(x => x.Year == year).ToList();
                    foreach (int week in wasteThisYear.Select(x => x.Week).Distinct())
                    {
                        List<ReportWasteModel> wasteThisWeek = wasteThisYear.Where(x => x.Week == week).ToList();
                        double dry = wasteThisWeek.Where(x => x.WasteType.ToLower().Contains("dry")).Sum(x => x.WasteDetail.Sum(y => y.Value));
                        double wet = wasteThisWeek.Where(x => x.WasteType.ToLower().Contains("wet")).Sum(x => x.WasteDetail.Sum(y => y.Value));
                        double dust = wasteThisWeek.Where(x => x.WasteType.ToLower().Contains("dust")).Sum(x => x.WasteDetail.Sum(y => y.Value));
                        double hot = wasteThisWeek.Where(x => x.WasteType.ToLower().Contains("hot")).Sum(x => x.WasteDetail.Sum(y => y.Value));
                        query += @"
UPDATE[dbo].[DashboardWastes] SET [Dry] = '" + dry + "', [Wet] = '" + wet + "', [Dust] = '" + dust + "', [Hot] = '" + hot + "', [ModifiedDate] = '" + modifiedDate + @"' WHERE [Year] = '" + year + @"' AND [Week] = '" + week + @"' AND [Location] = '" + location + @"' AND [Type] = '" + type + @"'
IF @@ROWCOUNT=0 INSERT INTO [dbo].[DashboardWastes] ([Year],[Week],[Location],[Type],[Dry],[Wet],[Dust],[Hot],[ModifiedDate]) VALUES ('" + year + "','" + week + "','" + location + "','" + type + "','" + dry + "','" + wet + "','" + dust + "','" + hot + "','" + modifiedDate + "')";
                    }
                }
                ExecuteQuery(query);
                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }

        }

        [HttpPost]
		public ActionResult ExportTable(List<ReportWasteModel> waste)
		{
			try
			{

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Waste");
				int cellPos = 1;
				int rowPos = 1;
				List<String> Days = new List<string>() { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
				Dictionary<string, List<List<string>>> sheets = new Dictionary<string, List<List<string>>>();
				List<List<string>> sheetVal = new List<List<string>>();
				List<String> HeaderShift = new List<String>(){
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3",
					"Shift 1",
					"Shift 2",
					"Shift 3" };
				//sheetVal.Add(Header);

				Sheet.Cells[1, 1].Value = "Week";
				Sheet.Cells[1, 2].Value = "Year";
				Sheet.Cells[1, 3].Value = "Area";
				Sheet.Cells[1, 4].Value = "Waste";
				Sheet.Cells[1, 5].Value = "Waste Type";
				Sheet.Cells[1, 1, 2, 1].Merge = true;
				Sheet.Cells[1, 2, 2, 2].Merge = true;
				Sheet.Cells[1, 3, 2, 3].Merge = true;
				Sheet.Cells[1, 4, 2, 4].Merge = true;
				Sheet.Cells[1, 5, 2, 5].Merge = true;

				Sheet.Cells[1, 1, 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				Sheet.Cells[1, 2, 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				Sheet.Cells[1, 3, 2, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				Sheet.Cells[1, 4, 2, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				Sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);

				cellPos = 6;
				foreach (String hValue in Days)
				{
					Sheet.Cells[1, cellPos].Value = hValue;
					Sheet.Cells[1, cellPos, 1, cellPos + 2].Merge = true;
					Sheet.Cells[1, cellPos, 1, cellPos + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					cellPos = cellPos + 3;
				}
				cellPos = 6;
				foreach (String hValue in HeaderShift)
				{
					Sheet.Cells[2, cellPos].Value = hValue;
					Sheet.Cells[2, cellPos].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					cellPos++;
				}
				Sheet.Cells[1, cellPos].Value = "Total";
				Sheet.Cells[1, cellPos, 2, cellPos].Merge = true;
				Sheet.Cells[1, cellPos, 2, cellPos].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				Sheet.Cells[1, 1, 2, cellPos].AutoFitColumns();
				using (var range = Sheet.Cells[1, 1, 2, cellPos])
				{
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
					range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				}
				rowPos = 2;
				foreach (ReportWasteModel wm in waste)
				{
					cellPos = 1;
					rowPos++;
					List<String> Content = new List<String>();

					Sheet.Cells[rowPos, cellPos++].Value = wm.Week.ToString();
					Sheet.Cells[rowPos, cellPos++].Value = wm.Year.ToString();
					Sheet.Cells[rowPos, cellPos++].Value = wm.Area;
					Sheet.Cells[rowPos, cellPos++].Value = wm.Waste;
					Sheet.Cells[rowPos, cellPos++].Value = wm.WasteType;

					double totalWaste = 0;
					if (wm.WasteDetail != null)
					{
						foreach (String day in Days)
						{
							for (int shift = 1; shift <= 3; shift++)
							{
								WasteDetailModel wdm = wm.WasteDetail.Where(x => x.Date == day && x.Shift == shift).FirstOrDefault();
								if (wdm != null)
								{
									Sheet.Cells[rowPos, cellPos++].Value = wdm.Value.ToString();
									totalWaste += wdm.Value;
								}
								else
								{
									Sheet.Cells[rowPos, cellPos++].Value = "-";
								}
							}
						}
					}
					Sheet.Cells[rowPos, cellPos].Value = totalWaste.ToString();
				}

				Sheet.Cells[1, 1, rowPos, cellPos].AutoFitColumns();
				Sheet.Cells[1, 1, rowPos, cellPos].Style.Border.BorderAround(ExcelBorderStyle.Thin);

				Session["DownloadExcel_FileManager"] = Ep.GetAsByteArray();
				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}

		}
		public ActionResult Download()
		{

			if (Session["DownloadExcel_FileManager"] != null)
			{
				byte[] data = Session["DownloadExcel_FileManager"] as byte[];
				return File(data, "application/octet-stream", "FileManager.xlsx");
			}
			else
			{
				return new EmptyResult();
			}
		}
		[HttpPost]
		public ActionResult GetReportWithParam(string StartDate, string EndDate, string LPH, string ProdCenterID, string Waste, string WasteType) //param nya menyusul
		{
			try
			{
				DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
				DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

				if (dtStartDate > dtEndDate)
				{
					SetFalseTempData("Start Date must be less than End Date");
					return RedirectToAction("Index");
				}

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(long.Parse(ProdCenterID), "productioncenter");
				List<LPHModel> lphList = _ppLphAppService.GetAll(true).DeserializeToLPHList()
					.Where(x =>
						locationIdList.Any(y => y == x.LocationID) &&
						x.Header == LPH
					).ToList();
				List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.GetAll(true).DeserializeToPPLPHSubmissionsList()
					.Where(x => lphList.Any(y => y.ID == x.LPHID) &&
						x.Date >= dtStartDate &&
						x.Date <= dtEndDate).ToList();
				List<PPLPHApprovalsModel> approval = _ppLphApprovalAppService.GetAll(true).DeserializeToPPLPHApprovalList();
				lphList = lphList.Where(l =>
				{
					PPLPHSubmissionsModel s = submissions.Where(x => x.LPHID == l.ID).ToList().FirstOrDefault();
					if (s != null) return approval.Where(x => x.LPHSubmissionID == s.ID && x.Status.Trim().ToLower() == "approved").Count() > 0;
					return false;
				}).ToList();

				List<LPHExtrasModel> wasteList = _ppLphExtrasAppService.GetAll(true).DeserializeToLPHExtraList()
					.Where(x => lphList.Any(y => y.ID == x.LPHID) && x.HeaderName == "waste").ToList();
				if (Waste != "")
				{
					List<LPHExtrasModel> wasteFiltered = wasteList.Where(x => x.FieldName == "Waste" && x.Value.Contains(Waste)).ToList();
					wasteList = wasteList.Where(x => wasteFiltered.Any(y => y.LPHID == x.LPHID && y.RowNumber == x.RowNumber)).ToList();
				}
				if (WasteType != "")
				{
					List<LPHExtrasModel> wasteFiltered = wasteList.Where(x => x.FieldName == "WasteType" && x.Value.Contains(WasteType)).ToList();
					wasteList = wasteList.Where(x => wasteFiltered.Any(y => y.LPHID == x.LPHID && y.RowNumber == x.RowNumber)).ToList();
				}

				List<WasteModel> wasteMatrik = new List<WasteModel>();
				//wasteList.ForEach(extraItem=>
				//{

				//});
				lphList.ForEach(lphItem =>
				{
					List<LPHExtrasModel> wasteFiltered = wasteList.Where(x => x.LPHID == lphItem.ID).ToList();
					wasteFiltered.GroupBy(x => x.RowNumber).Select(g => g.First()).ToList().ForEach(rowNum =>
					{
						List<LPHExtrasModel> wasteTmp = wasteFiltered.Where(x => x.RowNumber == rowNum.RowNumber).ToList();
						WasteModel wmChild = new WasteModel();
						wmChild.LPHID = lphItem.ID;
						wasteTmp.ForEach(extItem =>
						{
							switch (extItem.FieldName)
							{
								case "Area":
									wmChild.Area = extItem.Value;
									break;
								case "Waste":
									wmChild.Waste = extItem.Value;
									break;
								case "WasteType":
									wmChild.WasteType = extItem.Value;
									break;
								case "Frequency":
									wmChild.Frequency = extItem.Value;
									break;
								case "Weight":
									wmChild.Weight = double.TryParse(extItem.Value, out double result) ? result : 0;
									break;
								case "Remarks":
									wmChild.Remarks = extItem.Value;
									break;
							}
						});
						wasteMatrik.Add(wmChild);
					});
				});

				List<ReportWasteModel> dataFinal = new List<ReportWasteModel>();
				int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
				wasteMatrik.Select(x => x.Area).Distinct().ToList().ForEach(wasteByArea =>
				{
					List<WasteModel> wasteThisArea = wasteMatrik.Where(x => x.Area == wasteByArea).ToList();
					wasteThisArea.Select(x => x.Waste).Distinct().ToList().ForEach(wasteByWaste =>
					{
						WasteModel wasteThisWaste = wasteThisArea.Where(x => x.Waste == wasteByWaste).ToList().FirstOrDefault();

						DateTime endDate = dtEndDate/*.AddDays(DayOfWeek.Saturday - dtEndDate.DayOfWeek + 1)*/;
						ReportWasteModel rwm = new ReportWasteModel
						{
							Week = getWeek(dtStartDate),
							Year = dtStartDate.Year,
							Area = wasteByArea,
							Waste = wasteByWaste,
							WasteType = wasteThisWaste.WasteType == null ? "" : wasteThisWaste.WasteType,
							WasteDetail = new List<WasteDetailModel>(),
							Total = 0,
						};
						for (DateTime processedDate = dtStartDate/*.AddDays(-((7 + (dtStartDate.DayOfWeek - DayOfWeek.Monday)) % 7))*/; processedDate <= endDate; processedDate = processedDate.AddDays(1))
						{
							if (getWeek(processedDate) != rwm.Week)
							{
								dataFinal.Add(rwm);
								rwm = new ReportWasteModel
								{
									Week = getWeek(processedDate),
									Year = processedDate.Year,
									Area = wasteByArea,
									Waste = wasteByWaste,
									WasteType = wasteThisWaste.WasteType == null ? "" : wasteThisWaste.WasteType,
									WasteDetail = new List<WasteDetailModel>(),
									Total = 0,
								};
							}
							for (int i = 1; i <= 3; i++)
							{
								List<PPLPHSubmissionsModel> subs = submissions.Where(x => x.Date == processedDate && x.Shift.Trim() == i.ToString()).ToList();
								double weightSum = 0;
								weightSum = wasteMatrik.Where(x => subs.Any(y => y.LPHID == x.LPHID) && x.Area == wasteByArea && x.Waste == wasteByWaste).Sum(x =>
								{
									return x.Weight;
								});
								WasteDetailModel wdm = new WasteDetailModel
								{
									Date = processedDate.DayOfWeek.ToString(),
									Shift = i,
									Value = weightSum,
								};
								rwm.WasteDetail.Add(wdm);

								rwm.Total += weightSum;

							}
						}
						dataFinal.Add(rwm);
					});
				});
				dataFinal = dataFinal.OrderBy(x => x.Week).OrderBy(x => x.Year).ToList();

				return Json(new { Status = "True", Waste = dataFinal }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { data = new List<PPLPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
