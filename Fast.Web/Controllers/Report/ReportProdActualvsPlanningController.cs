using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.Report;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportProdActualvsPlanningController : BaseController<WppStpModel>
    {
        private readonly IWppStpAppService _wppStpAppService;
        private readonly IWppAppService _wppAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IUserAppService _userAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly ILoggerAppService _logger;
        private readonly IMachineAppService _machineAppService;

        public ReportProdActualvsPlanningController(
            IWppAppService wppAppService,
            IWppStpAppService wppStpAppService,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IUserAppService userAppService,
            IEmployeeAppService employeeAppService,
            IMachineAppService machineAppService,
            ILoggerAppService logger
             )
        {
            _machineAppService = machineAppService;
            _wppStpAppService = wppStpAppService;
            _locationAppService = locationAppService;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
            _employeeAppService = employeeAppService;
            _wppAppService = wppAppService;
            _logger = logger;
        }

        #region ::Public Methods::
        public ActionResult Index()
        {
            ViewBag.GranularityList = DropDownHelper.BindDropDownGranularity();
            ViewBag.ProductList = DropDownHelper.BindDropDownMultiProduct();
            ViewBag.ProductionCenterList = DropDownHelper.GetMultiProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            return View();
        }

        public ActionResult ExportExcelByBrand(string startDate, string endDate, string startWeek, string endWeek, string locationIDList, string productType, string granularity)
        {
            try
            {
                #region Validate Parameters
                if (granularity == "Shiftly" || granularity == "Daily")
                {
                    if (string.IsNullOrEmpty(startDate))
                    {
                        SetFalseTempData("start date is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endDate))
                    {
                        SetFalseTempData("end date is not selected");
                        return RedirectToAction("Index");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(startWeek))
                    {
                        SetFalseTempData("start week is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endWeek))
                    {
                        SetFalseTempData("end week is not selected");
                        return RedirectToAction("Index");
                    }
                }

                if (locationIDList.Count() == 0)
                {
                    SetFalseTempData("Production center is not selected");
                    return RedirectToAction("Index");
                }
                #endregion

                #region Extract Parameters
                DateTime startDateFL, endDateFL;
                string weekNumber, endWeekNumber;
                List<long> locIDList;
                List<string> locationList;

                ExtractParameterExcel(startDate, endDate, startWeek, endWeek, locationIDList, granularity, out startDateFL, out endDateFL, out weekNumber, out endWeekNumber, out locIDList, out locationList);
                #endregion

                List<string> weekList = new List<string>();
                if (granularity == "Weekly")
                {
                    for (var week = int.Parse(weekNumber); week <= int.Parse(endWeekNumber); week++)
                    {
                        weekList.Add(week.ToString());
                    }
                }

                ReportProdActualPlanningModel model = GetWppByBrand(productType, startDateFL, endDateFL, weekList, locIDList, granularity, locationList);
                string location = string.Join(", ", locationList);

                if (locationIDList.Count() > 1)
                {
                    model.ShiftReports = GetShiftReport(startDateFL, endDateFL, model);
                }

                byte[] excelData = ExcelGenerator.ExportReportActualPlanningByBrand(model.ShiftReports, startDateFL, endDateFL, AccountName, location, model.Granularity, weekList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Actual-Report-By-Brand.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult ExportExcelByLU(string startDate, string endDate, string startWeek, string endWeek, string locationIDList, string productType, string granularity)
        {
            try
            {
                #region Validate Parameters
                if (granularity == "Shiftly" || granularity == "Daily")
                {
                    if (string.IsNullOrEmpty(startDate))
                    {
                        SetFalseTempData("start date is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endDate))
                    {
                        SetFalseTempData("end date is not selected");
                        return RedirectToAction("Index");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(startWeek))
                    {
                        SetFalseTempData("start week is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endWeek))
                    {
                        SetFalseTempData("end week is not selected");
                        return RedirectToAction("Index");
                    }
                }

                if (locationIDList.Count() == 0)
                {
                    SetFalseTempData("Production center is not selected");
                    return RedirectToAction("Index");
                }
                #endregion

                #region Extract Parameters
                DateTime startDateFL, endDateFL;
                string weekNumber, endWeekNumber;
                List<long> locIDList;
                List<string> locationList;

                ExtractParameterExcel(startDate, endDate, startWeek, endWeek, locationIDList, granularity, out startDateFL, out endDateFL, out weekNumber, out endWeekNumber, out locIDList, out locationList);
                #endregion

                List<string> weekList = new List<string>();
                if (granularity == "Weekly")
                {
                    for (var week = int.Parse(weekNumber); week <= int.Parse(endWeekNumber); week++)
                    {
                        weekList.Add(week.ToString());
                    }
                }

                ReportProdActualPlanningModel model = GetWppByLU(productType, startDateFL, endDateFL, weekList, locIDList, granularity, locationList);
                string location = string.Join(", ", locationList);

                if (locationIDList.Count() > 1)
                {
                    model.ShiftReports = GetShiftReport(startDateFL, endDateFL, model);
                }

                byte[] excelData = ExcelGenerator.ExportReportActualPlanningByLU(model.ShiftReports, startDateFL, endDateFL, AccountName, location, locIDList, model.Granularity, weekList);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Actual-Report-By-LU.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult ExportExcelRawData(string startDate, string endDate, string startWeek, string endWeek, string locationIDList, string productType, string granularity)
        {
            try
            {
                #region Validate Parameters
                if (granularity == "Shiftly" || granularity == "Daily")
                {
                    if (string.IsNullOrEmpty(startDate))
                    {
                        SetFalseTempData("start date is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endDate))
                    {
                        SetFalseTempData("end date is not selected");
                        return RedirectToAction("Index");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(startWeek))
                    {
                        SetFalseTempData("start week is not selected");
                        return RedirectToAction("Index");
                    }

                    if (string.IsNullOrEmpty(endWeek))
                    {
                        SetFalseTempData("end week is not selected");
                        return RedirectToAction("Index");
                    }
                }

                if (locationIDList.Count() == 0)
                {
                    SetFalseTempData("Production center is not selected");
                    return RedirectToAction("Index");
                }
                #endregion

                #region Extract Parameters
                DateTime startDateFL, endDateFL;
                string weekNumber, endWeekNumber;
                List<long> locIDList;
                List<string> locationList;

                ExtractParameterExcel(startDate, endDate, startWeek, endWeek, locationIDList, granularity, out startDateFL, out endDateFL, out weekNumber, out endWeekNumber, out locIDList, out locationList);
                #endregion

                List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList();
                machineList = machineList.Where(x => x.LinkUp != null).ToList();

                List<WppStpModel> wpps = GetWPPList(startDateFL, endDateFL, productType, locIDList);

                List<string> wppLocationList = wpps.Select(x => x.Location).Distinct().ToList();

                List<ReportActualPlanningShiftModel> result = new List<ReportActualPlanningShiftModel>();
                foreach (var location in wppLocationList)
                {
                    List<string> itemCodeList = wpps.Where(x => x.Location == location).Select(x => x.Brand).Distinct().ToList();
                    foreach (var itemCode in itemCodeList)
                    {
                        for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                        {
                            var wpp = wpps.Where(x => x.Brand == itemCode && x.Date.Date == day.Date && x.Location == location).FirstOrDefault();
                            if (wpp == null) continue;

                            var machineLU = machineList.Where(x => wpp.Packer != null && x.Code == wpp.Packer.Trim()).FirstOrDefault();
                            string packerLU = machineLU == null ? "-" : machineLU.LinkUp;

                            if (wpp.Shift1 != 0)
                            {
                                ReportActualPlanningShiftModel planShift1Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Plan",
                                    Shift = "1",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Shift1
                                };

                                ReportActualPlanningShiftModel actualShift1Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Actual",
                                    Shift = "1",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Actual1
                                };

                                result.Add(planShift1Report);
                                result.Add(actualShift1Report);
                            }
                            if (wpp.Shift2 != 0)
                            {
                                ReportActualPlanningShiftModel planShift2Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Plan",
                                    Shift = "2",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Shift2
                                };
                                ReportActualPlanningShiftModel actualShift2Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Actual",
                                    Shift = "2",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Actual2
                                };

                                result.Add(planShift2Report);
                                result.Add(actualShift2Report);
                            }
                            if (wpp.Shift3 != 0)
                            {
                                ReportActualPlanningShiftModel planShift3Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Plan",
                                    Shift = "3",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Shift3
                                };
                                ReportActualPlanningShiftModel actualShift3Report = new ReportActualPlanningShiftModel
                                {
                                    Date = day,
                                    Location = location,
                                    ItemCode = wpp.Brand,
                                    Description = wpp.Description,
                                    Type = "Actual",
                                    Shift = "3",
                                    UOM = WebConstants.DEFAULT_STICK,
                                    LU = packerLU,
                                    Market = wpp.Others,
                                    Total = wpp.Actual3
                                };

                                result.Add(planShift3Report);
                                result.Add(actualShift3Report);
                            }
                        }
                    }
                }

                result = result.OrderBy(x => x.Location).ThenBy(x => x.Date).ThenBy(x => x.Shift).ThenBy(x => x.ItemCode).ToList();

                byte[] excelData = ExcelGenerator.ExportReportActualPlanningRawData(result, startDateFL, endDateFL, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Planning-Actual-Raw-Data-Report.xlsx");
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
        public ActionResult GetReportDataByBrand(string startDate, string endDate, string startWeek, string endWeek, string[] locationIDList, string productType, string granularity)
        {
            #region Validate Parameters
            if (granularity == "Shiftly" || granularity == "Daily")
            {
                if (string.IsNullOrEmpty(startDate))
                {
                    SetFalseTempData("start date is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endDate))
                {
                    SetFalseTempData("end date is not selected");
                    return RedirectToAction("Index");
                }
            }
            else
            {
                if (string.IsNullOrEmpty(startWeek))
                {
                    SetFalseTempData("start week is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endWeek))
                {
                    SetFalseTempData("end week is not selected");
                    return RedirectToAction("Index");
                }
            }

            if (locationIDList.Count() == 0)
            {
                SetFalseTempData("Production center is not selected");
                return RedirectToAction("Index");
            }
            #endregion

            #region Extract Parameters
            DateTime startDateFL, endDateFL;
            string weekNumber, endWeekNumber;
            List<long> locIDList;
            List<string> locationList;

            ExtractParameter(startDate, endDate, startWeek, endWeek, locationIDList, granularity, out startDateFL, out endDateFL, out weekNumber, out endWeekNumber, out locIDList, out locationList);
            #endregion

            List<string> dateList = new List<string>();
            if (granularity == "Shiftly" || granularity == "Daily")
            {
                for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                {
                    dateList.Add(day.ToString("dd-MMM-yy"));
                }
            }

            List<string> weekList = new List<string>();
            if (granularity == "Weekly")
            {
                for (var week = int.Parse(weekNumber); week <= int.Parse(endWeekNumber); week++)
                {
                    weekList.Add(week.ToString());
                }
            }

            ReportProdActualPlanningModel model = GetWppByBrand(productType, startDateFL, endDateFL, weekList, locIDList, granularity);

            if (locationIDList.Count() > 1)
            {
                model.ShiftReports = GetShiftReport(startDateFL, endDateFL, model);
            }

            // Returning Json Data    
            return Json(new { dataList = model.ShiftReports, dateList, weekList, model.Granularity, weekNumber, endWeekNumber }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetReportDataByLU(string startDate, string endDate, string startWeek, string endWeek, string[] locationIDList, string productType, string granularity)
        {
            #region Validate Parameters
            if (granularity == "Shiftly" || granularity == "Daily")
            {
                if (string.IsNullOrEmpty(startDate))
                {
                    SetFalseTempData("start date is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endDate))
                {
                    SetFalseTempData("end date is not selected");
                    return RedirectToAction("Index");
                }
            }
            else
            {
                if (string.IsNullOrEmpty(startWeek))
                {
                    SetFalseTempData("start week is not selected");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(endWeek))
                {
                    SetFalseTempData("end week is not selected");
                    return RedirectToAction("Index");
                }
            }

            if (locationIDList.Count() == 0)
            {
                SetFalseTempData("Production center is not selected");
                return RedirectToAction("Index");
            }
            #endregion

            #region Extract Parameters
            DateTime startDateFL, endDateFL;
            string weekNumber, endWeekNumber;
            List<long> locIDList;
            List<string> locationList;

            ExtractParameter(startDate, endDate, startWeek, endWeek, locationIDList, granularity, out startDateFL, out endDateFL, out weekNumber, out endWeekNumber, out locIDList, out locationList);
            #endregion

            List<string> dateList = new List<string>();
            if (granularity == "Shiftly" || granularity == "Daily")
            {
                for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                {
                    dateList.Add(day.ToString("dd-MMM-yy"));
                }
            }

            List<string> weekList = new List<string>();
            if (granularity == "Weekly")
            {
                for (var week = int.Parse(weekNumber); week <= int.Parse(endWeekNumber); week++)
                {
                    weekList.Add(week.ToString());
                }
            }

            ReportProdActualPlanningModel model = GetWppByLU(productType, startDateFL, endDateFL, weekList, locIDList, granularity);

            if (locationIDList.Count() > 1)
            {
                model.ShiftReports = GetShiftReport(startDateFL, endDateFL, model);
            }

            // Returning Json Data    
            return Json(new { dataList = model.ShiftReports, dateList, weekList, model.Granularity, weekNumber, endWeekNumber }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region ::Private Methods::
        private static List<ReportActualPlanningShiftModel> GetShiftReport(DateTime startDateFL, DateTime endDateFL, ReportProdActualPlanningModel model)
        {
            List<string> itemCodeList = model.ShiftReports.Select(x => x.ItemCode).Distinct().ToList();
            List<ReportActualPlanningShiftModel> shiftReports = new List<ReportActualPlanningShiftModel>();

            foreach (var itemCode in itemCodeList)
            {
                var tempList = model.ShiftReports.Where(x => x.ItemCode == itemCode).ToList();
                if (tempList.Count == 3)
                {
                    shiftReports.AddRange(tempList);
                }
                else
                {
                    var tempPlanReports = tempList.Where(x => x.Type == "Plan").ToList();
                    var tempActualReports = tempList.Where(x => x.Type == "Actual").ToList();
                    var tempTotalReports = tempList.Where(x => x.Type != "Plan" && x.Type != "Actual").ToList();

                    ReportActualPlanningShiftModel planReport = new ReportActualPlanningShiftModel();
                    planReport.ItemCode = tempPlanReports[0].ItemCode;
                    planReport.Description = tempPlanReports[0].Description;
                    planReport.Type = tempPlanReports[0].Type;
                    planReport.UOM = tempPlanReports[0].UOM;
                    planReport.Remark = tempPlanReports[0].Remark;
                    planReport.LU = tempPlanReports[0].LU;
                    planReport.Market = tempPlanReports[0].Market;
                    planReport.Total = tempPlanReports.Sum(x => x.Total);
                    planReport.ShiftList = new List<ShiftModel>();

                    List<ShiftModel> tempShiftModelList = new List<ShiftModel>();

                    foreach (var planRpt in tempPlanReports)
                    {
                        tempShiftModelList.AddRange(planRpt.ShiftList);
                    }

                    for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                    {
                        ShiftModel shift = new ShiftModel();
                        shift.Date = day;
                        shift.Shift1 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift1);
                        shift.Shift2 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift2);
                        shift.Shift3 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift3);
                        shift.AllShift = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.AllShift);
                        planReport.ShiftList.Add(shift);
                    }

                    ReportActualPlanningShiftModel actualReport = new ReportActualPlanningShiftModel();
                    actualReport.ItemCode = tempActualReports[0].ItemCode;
                    actualReport.Description = tempActualReports[0].Description;
                    actualReport.Type = tempActualReports[0].Type;
                    actualReport.UOM = tempActualReports[0].UOM;
                    actualReport.Remark = tempActualReports[0].Remark;
                    actualReport.LU = tempActualReports[0].LU;
                    actualReport.Market = tempActualReports[0].Market;
                    actualReport.Total = tempActualReports.Sum(x => x.Total);
                    actualReport.ShiftList = new List<ShiftModel>();

                    tempShiftModelList = new List<ShiftModel>();

                    foreach (var actualRpt in tempActualReports)
                    {
                        tempShiftModelList.AddRange(actualRpt.ShiftList);
                    }

                    for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                    {
                        ShiftModel shift = new ShiftModel();
                        shift.Date = day;
                        shift.Shift1 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift1);
                        shift.Shift2 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift2);
                        shift.Shift3 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift3);
                        shift.AllShift = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.AllShift);
                        actualReport.ShiftList.Add(shift);
                    }

                    ReportActualPlanningShiftModel totalReport = new ReportActualPlanningShiftModel();
                    totalReport.ItemCode = tempTotalReports[0].ItemCode;
                    totalReport.Description = tempTotalReports[0].Description;
                    totalReport.Type = tempTotalReports[0].Type;
                    totalReport.UOM = tempTotalReports[0].UOM;
                    totalReport.Remark = tempTotalReports[0].Remark;
                    totalReport.LU = tempTotalReports[0].LU;
                    totalReport.Market = tempTotalReports[0].Market;
                    totalReport.Total = tempTotalReports.Sum(x => x.Total) / tempTotalReports.Count;
                    totalReport.TotalStr = totalReport.Total == 0 ? "" : totalReport.Total.ToString("F") + " %";
                    totalReport.ShiftList = new List<ShiftModel>();

                    tempShiftModelList = new List<ShiftModel>();

                    foreach (var totalRpt in tempTotalReports)
                    {
                        tempShiftModelList.AddRange(totalRpt.ShiftList);
                    }

                    for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                    {
                        ShiftModel shift = new ShiftModel();
                        shift.Date = day;
                        shift.Shift1 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift1);
                        shift.Shift2 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift2);
                        shift.Shift3 = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.Shift3);
                        shift.AllShift = tempShiftModelList.Where(x => x.Date == day).Sum(x => x.AllShift);

                        shift.Shift1Str = shift.Shift1 == 0 ? "" : shift.Shift1.ToString("F") + " %";
                        shift.Shift2Str = shift.Shift2 == 0 ? "" : shift.Shift2.ToString("F") + " %";
                        shift.Shift3Str = shift.Shift3 == 0 ? "" : shift.Shift3.ToString("F") + " %";
                        shift.AllShiftStr = shift.AllShift == 0 ? "" : shift.AllShift.ToString("F") + " %";

                        totalReport.ShiftList.Add(shift);
                    }

                    shiftReports.Add(planReport);
                    shiftReports.Add(actualReport);
                    shiftReports.Add(totalReport);
                }
            }

            return shiftReports;
        }

        private void ExtractParameter(string startDate, string endDate, string startWeek, string endWeek, string[] locationIDList, string granularity, out DateTime startDateFL, out DateTime endDateFL, out string weekNumber, out string endWeekNumber, out List<long> locIDList, out List<string> locationList)
        {
            weekNumber = "";
            endWeekNumber = "";

            if (granularity == "Shiftly" || granularity == "Daily")
            {
                startDateFL = DateTime.Parse(startDate);
                endDateFL = DateTime.Parse(endDate);
            }
            else
            {
                string[] startWeekFilter = startWeek.Split('-');

                weekNumber = startWeekFilter[1].Replace("W", "").Replace("0", "");

                startDateFL = Helper.FirstDateOfWeek(startWeekFilter[0], weekNumber);

                string[] endWeekFilter = endWeek.Split('-');

                endWeekNumber = endWeekFilter[1].Replace("W", "").Replace("0", "");

                endDateFL = Helper.LastDateOfWeek(endWeekFilter[0], endWeekNumber);
            }

            locIDList = new List<long>();
            locationList = new List<string>();

            foreach (var item in locationIDList)
            {
                locationList.Add(_locationAppService.GetLocationFullCode(long.Parse(item)));
                locIDList.Add(long.Parse(item));
            }
        }

        private void ExtractParameterExcel(string startDate, string endDate, string startWeek, string endWeek, string locationIDList, string granularity, out DateTime startDateFL, out DateTime endDateFL, out string weekNumber, out string endWeekNumber, out List<long> locIDList, out List<string> locationList)
        {
            weekNumber = "";
            endWeekNumber = "";

            if (granularity == "Shiftly" || granularity == "Daily")
            {
                startDateFL = DateTime.Parse(startDate);
                endDateFL = DateTime.Parse(endDate);
            }
            else
            {
                string[] startWeekFilter = startWeek.Split('-');

                weekNumber = startWeekFilter[1].Replace("W", "").Replace("0", "");

                startDateFL = Helper.FirstDateOfWeek(startWeekFilter[0], weekNumber);

                string[] endWeekFilter = endWeek.Split('-');

                endWeekNumber = endWeekFilter[1].Replace("W", "").Replace("0", "");

                endDateFL = Helper.LastDateOfWeek(endWeekFilter[0], endWeekNumber);
            }

            locIDList = new List<long>();
            locationList = new List<string>();

            foreach (var item in locationIDList.Split(','))
            {
                locationList.Add(_locationAppService.GetLocationFullCode(long.Parse(item)));
                locIDList.Add(long.Parse(item));
            }
        }

        private ReportProdActualPlanningModel GetWppByLU(string productType, DateTime startDateFL, DateTime endDateFL, List<string> weekList, List<long> locIDList, string granularity, List<string> locationList = null)
        {
            ReportProdActualPlanningModel model = new ReportProdActualPlanningModel();
            model.Granularity = granularity;

            List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList();
            machineList = machineList.Where(x => x.LinkUp != null).ToList();

            List<WppStpModel> wpps = GetWPPList(startDateFL, endDateFL, productType, locIDList);
            wpps = wpps.Where(x => !string.IsNullOrEmpty(x.Packer)).ToList();

            var pcList = DropDownHelper.GetMultiProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            List<string> dateList = new List<string>();
            for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
            {
                dateList.Add(day.ToString("dd-MMM-yy"));
            }

            foreach (var locID in locIDList)
            {
                string location = pcList.Where(x => x.Value == locID.ToString()).First().Text;

                List<string> packerList = wpps.Where(x => x.LocationID == locID).OrderBy(x => x.Date).ThenByDescending(x => x.Shift1).ThenByDescending(x => x.Shift2).ThenByDescending(x => x.Shift3).Select(x => x.Packer.Trim()).Distinct().ToList();
                foreach (var packerCode in packerList)
                {
                    var wppTempList = wpps.Where(x => x.Packer.Trim() == packerCode).OrderBy(x => x.Date).ThenByDescending(x => x.Shift1).ThenByDescending(x => x.Shift2).ThenByDescending(x => x.Shift3).ToList();

                    var machineLU = machineList.Where(x => x.Code == packerCode).FirstOrDefault();
                    string packerLU = machineLU == null ? "-" : machineLU.LinkUp;

                    List<string> brandList = new List<string>();

                    foreach (var wpp in wppTempList)
                    {
                        if (brandList.Any(x => x == wpp.Brand) || (wpp.Shift1 == 0 && wpp.Shift2 == 0 && wpp.Shift3 == 0))
                            continue;
                        else
                            brandList.Add(wpp.Brand);

                        ReportActualPlanningShiftModel planReport = new ReportActualPlanningShiftModel
                        {
                            Location = location,
                            LocationID = locID,
                            ItemCode = wpp.Brand,
                            Description = wpp.Description,
                            Type = "Plan",
                            LU = packerLU,
                            UOM = WebConstants.DEFAULT_STICK,
                            Market = wpp.Others,
                            Remark = string.Empty,
                            ShiftList = new List<ShiftModel>()
                        };

                        ReportActualPlanningShiftModel actualReport = new ReportActualPlanningShiftModel
                        {
                            Location = location,
                            LocationID = locID,
                            ItemCode = wpp.Brand,
                            Description = wpp.Description,
                            Type = "Actual",
                            LU = packerLU,
                            UOM = WebConstants.DEFAULT_STICK,
                            Market = wpp.Others,
                            Remark = string.Empty,
                            ShiftList = new List<ShiftModel>()
                        };

                        ReportActualPlanningShiftModel totalReport = new ReportActualPlanningShiftModel
                        {
                            Location = location,
                            LocationID = locID,
                            ItemCode = wpp.Brand,
                            Description = wpp.Description,
                            Type = "% (Actual / Plan)",
                            LU = packerLU,
                            UOM = WebConstants.DEFAULT_STICK,
                            Market = wpp.Others,
                            Remark = string.Empty,
                            ShiftList = new List<ShiftModel>()
                        };

                        if (granularity == "Shiftly" || granularity == "Daily")
                        {
                            for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                            {
                                var wppTemp = wpps.Where(x => x.Date.Date == day.Date && x.Brand == wpp.Brand).ToList();

                                ShiftModel planShift = new ShiftModel
                                {
                                    Date = day,
                                    Shift1 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1),
                                    Shift2 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift2),
                                    Shift3 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift3),
                                    AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift2) + wppTemp.Sum(x => x.Shift3)
                                };
                                planReport.ShiftList.Add(planShift);

                                ShiftModel actualShift = new ShiftModel
                                {
                                    Date = day,
                                    Shift1 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1),
                                    Shift2 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual2),
                                    Shift3 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual3),
                                    AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1) + wppTemp.Sum(x => x.Actual2) + wppTemp.Sum(x => x.Actual3)
                                };
                                actualReport.ShiftList.Add(actualShift);

                                if (wppTemp.Count == 0)
                                {
                                    ShiftModel totalShift = new ShiftModel
                                    {
                                        Date = day,
                                        Shift1 = 0,
                                        Shift2 = 0,
                                        Shift3 = 0,
                                        AllShift = 0
                                    };

                                    totalShift.Shift1Str = "";
                                    totalShift.Shift2Str = "";
                                    totalShift.Shift3Str = "";
                                    totalShift.AllShiftStr = "";

                                    totalReport.ShiftList.Add(totalShift);
                                }
                                else
                                {
                                    ShiftModel totalShift = new ShiftModel
                                    {
                                        Date = day,
                                        Shift1 = wppTemp.Sum(x => x.Actual1) == 0 || wppTemp.Sum(x => x.Shift1) == 0 ? 0 : (wppTemp.Sum(x => x.Actual1) / wppTemp.Sum(x => x.Shift1)) * 100,
                                        Shift2 = wppTemp.Sum(x => x.Actual2) == 0 || wppTemp.Sum(x => x.Shift2) == 0 ? 0 : (wppTemp.Sum(x => x.Actual2) / wppTemp.Sum(x => x.Shift2)) * 100,
                                        Shift3 = wppTemp.Sum(x => x.Actual3) == 0 || wppTemp.Sum(x => x.Shift3) == 0 ? 0 : (wppTemp.Sum(x => x.Actual3) / wppTemp.Sum(x => x.Shift3)) * 100,
                                        AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                                    };
                                    totalShift.Shift1Str = totalShift.Shift1 == 0 ? "" : totalShift.Shift1.ToString("F") + " %";
                                    totalShift.Shift2Str = totalShift.Shift2 == 0 ? "" : totalShift.Shift2.ToString("F") + " %";
                                    totalShift.Shift3Str = totalShift.Shift3 == 0 ? "" : totalShift.Shift3.ToString("F") + " %";
                                    totalShift.AllShiftStr = totalShift.AllShift == 0 ? "" : totalShift.AllShift.ToString("F") + " %";

                                    totalReport.ShiftList.Add(totalShift);
                                }
                            }
                        }
                        else
                        {
                            DateTime startDate;
                            DateTime endDate;

                            int index = 0;
                            foreach (var weekNumber in weekList)
                            {
                                startDate = startDateFL.AddDays(index++ * 7);
                                endDate = startDate.AddDays(6);

                                var wppTemp = wpps.Where(x => x.Date.Date >= startDate.Date && x.Date.Date <= endDate.Date && x.Brand == wpp.Brand).ToList();

                                ShiftModel planShift = new ShiftModel
                                {
                                    Date = startDate,
                                    AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift1)
                                };
                                planReport.ShiftList.Add(planShift);

                                ShiftModel actualShift = new ShiftModel
                                {
                                    Date = startDate,
                                    AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1) + wppTemp.Sum(x => x.Actual2) + wppTemp.Sum(x => x.Actual3)
                                };
                                actualReport.ShiftList.Add(actualShift);

                                if (wppTemp.Count == 0)
                                {
                                    ShiftModel totalShift = new ShiftModel
                                    {
                                        Date = startDate,
                                        AllShift = 0
                                    };

                                    totalShift.AllShiftStr = "";

                                    totalReport.ShiftList.Add(totalShift);
                                }
                                else
                                {
                                    ShiftModel totalShift = new ShiftModel
                                    {
                                        Date = startDate,
                                        AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                                    };

                                    totalShift.AllShiftStr = totalShift.AllShift == 0 ? "" : totalShift.AllShift.ToString("F") + " %";

                                    totalReport.ShiftList.Add(totalShift);
                                }
                            }
                        }

                        decimal totalShift1 = planReport.ShiftList.Sum(x => x.Shift1);
                        decimal totalShift2 = planReport.ShiftList.Sum(x => x.Shift2);
                        decimal totalShift3 = planReport.ShiftList.Sum(x => x.Shift3);
                        decimal totalAllShift = planReport.ShiftList.Sum(x => x.AllShift);

                        planReport.Total = granularity == "Daily" ? totalShift1 + totalShift2 + totalShift3 : totalAllShift;

                        totalShift1 = actualReport.ShiftList.Sum(x => x.Shift1);
                        totalShift2 = actualReport.ShiftList.Sum(x => x.Shift2);
                        totalShift3 = actualReport.ShiftList.Sum(x => x.Shift3);
                        totalAllShift = actualReport.ShiftList.Sum(x => x.AllShift);

                        actualReport.Total = granularity == "Daily" ? totalShift1 + totalShift2 + totalShift3 : totalAllShift;

                        totalReport.Total = actualReport.Total == 0 || planReport.Total == 0 ? 0 : (actualReport.Total / planReport.Total) * 100;
                        totalReport.TotalStr = totalReport.Total == 0 ? "" : totalReport.Total.ToString("F") + " %";

                        model.ShiftReports.Add(planReport);
                        model.ShiftReports.Add(actualReport);
                        model.ShiftReports.Add(totalReport);
                    }

                    // add sub total
                    ReportActualPlanningShiftModel planSubTotalReport = new ReportActualPlanningShiftModel
                    {
                        Location = string.Empty,
                        LocationID = locID,
                        ItemCode = string.Empty,
                        Description = string.Empty,
                        Type = "Plan",
                        LU = packerLU,
                        UOM = WebConstants.DEFAULT_STICK,
                        Market = string.Empty,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    ReportActualPlanningShiftModel actualSubTotalReport = new ReportActualPlanningShiftModel
                    {
                        Location = string.Empty,
                        LocationID = locID,
                        ItemCode = string.Empty,
                        Description = string.Empty,
                        Type = "Actual",
                        LU = packerLU,
                        UOM = WebConstants.DEFAULT_STICK,
                        Market = string.Empty,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    ReportActualPlanningShiftModel totalSubTotalReport = new ReportActualPlanningShiftModel
                    {
                        Location = "Sub Total",
                        LocationID = locID,
                        ItemCode = string.Empty,
                        Description = string.Empty,
                        Type = "% (Actual / Plan)",
                        LU = packerLU,
                        UOM = WebConstants.DEFAULT_STICK,
                        Market = string.Empty,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    if (granularity == "Shiftly" || granularity == "Daily")
                    {
                        for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                        {
                            var shiftPlanList = model.ShiftReports.Where(x => x.LU == packerLU && x.Type == "Plan").ToList();
                            var shiftActualList = model.ShiftReports.Where(x => x.LU == packerLU && x.Type == "Actual").ToList();

                            decimal s1 = 0, s2 = 0, s3 = 0, allS = 0;
                            foreach (var item in shiftPlanList)
                            {
                                var slist = item.ShiftList.Where(x => x.Date.Date == day.Date).ToList();
                                s1 += slist.Sum(x => x.Shift1);
                                s2 += slist.Sum(x => x.Shift2);
                                s3 += slist.Sum(x => x.Shift3);
                                allS += slist.Sum(x => x.AllShift);
                            }

                            ShiftModel planShift = new ShiftModel
                            {
                                Date = day,
                                Shift1 = s1,
                                Shift2 = s2,
                                Shift3 = s3,
                                AllShift = allS
                            };

                            planSubTotalReport.ShiftList.Add(planShift);

                            s1 = 0;
                            s2 = 0;
                            s3 = 0;
                            allS = 0;

                            foreach (var item in shiftActualList)
                            {
                                var slist = item.ShiftList.Where(x => x.Date.Date == day.Date).ToList();
                                s1 += slist.Sum(x => x.Shift1);
                                s2 += slist.Sum(x => x.Shift2);
                                s3 += slist.Sum(x => x.Shift3);
                                allS += slist.Sum(x => x.AllShift);
                            }

                            ShiftModel actualShift = new ShiftModel
                            {
                                Date = day,
                                Shift1 = s1,
                                Shift2 = s2,
                                Shift3 = s3,
                                AllShift = allS
                            };
                            actualSubTotalReport.ShiftList.Add(actualShift);

                            ShiftModel totalSubTotalShift = new ShiftModel
                            {
                                Date = day,
                                Shift1 = actualShift.Shift1 == 0 || planShift.Shift1 == 0 ? 0 : (actualShift.Shift1 / planShift.Shift1) * 100,
                                Shift2 = actualShift.Shift2 == 0 || planShift.Shift2 == 0 ? 0 : (actualShift.Shift2 / planShift.Shift2) * 100,
                                Shift3 = actualShift.Shift3 == 0 || planShift.Shift3 == 0 ? 0 : (actualShift.Shift3 / planShift.Shift3) * 100,
                                AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                            };

                            totalSubTotalShift.Shift1Str = totalSubTotalShift.Shift1 == 0 ? "" : totalSubTotalShift.Shift1.ToString("F") + " %";
                            totalSubTotalShift.Shift2Str = totalSubTotalShift.Shift2 == 0 ? "" : totalSubTotalShift.Shift2.ToString("F") + " %";
                            totalSubTotalShift.Shift3Str = totalSubTotalShift.Shift3 == 0 ? "" : totalSubTotalShift.Shift3.ToString("F") + " %";
                            totalSubTotalShift.AllShiftStr = totalSubTotalShift.AllShift == 0 ? "" : totalSubTotalShift.AllShift.ToString("F") + " %";

                            totalSubTotalReport.ShiftList.Add(totalSubTotalShift);
                        }
                    }
                    else
                    {
                        DateTime startDate;
                        DateTime endDate;

                        int index = 0;
                        foreach (var weekNumber in weekList)
                        {
                            startDate = startDateFL.AddDays(index++ * 7);
                            endDate = startDate.AddDays(6);

                            var shiftPlanList = model.ShiftReports.Where(x => x.LU == packerLU && x.Type == "Plan").ToList();
                            var shiftActualList = model.ShiftReports.Where(x => x.LU == packerLU && x.Type == "Actual").ToList();

                            decimal allS = 0;
                            foreach (var item in shiftPlanList)
                            {
                                var slist = item.ShiftList.Where(x => x.Date.Date >= startDate.Date && x.Date.Date <= endDate.Date).ToList();
                                allS = slist.Sum(x => x.AllShift);
                            }

                            ShiftModel planShift = new ShiftModel
                            {
                                AllShift = allS
                            };

                            planSubTotalReport.ShiftList.Add(planShift);

                            allS = 0;

                            foreach (var item in shiftActualList)
                            {
                                var slist = item.ShiftList.Where(x => x.Date.Date >= startDate.Date && x.Date.Date <= endDate.Date).ToList();
                                allS += slist.Sum(x => x.AllShift);
                            }

                            ShiftModel actualShift = new ShiftModel
                            {
                                AllShift = allS
                            };
                            actualSubTotalReport.ShiftList.Add(actualShift);

                            ShiftModel totalSubTotalShift = new ShiftModel
                            {
                                AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                            };

                            totalSubTotalShift.Shift1Str = totalSubTotalShift.Shift1 == 0 ? "" : totalSubTotalShift.Shift1.ToString("F") + " %";
                            totalSubTotalShift.Shift2Str = totalSubTotalShift.Shift2 == 0 ? "" : totalSubTotalShift.Shift2.ToString("F") + " %";
                            totalSubTotalShift.Shift3Str = totalSubTotalShift.Shift3 == 0 ? "" : totalSubTotalShift.Shift3.ToString("F") + " %";
                            totalSubTotalShift.AllShiftStr = totalSubTotalShift.AllShift == 0 ? "" : totalSubTotalShift.AllShift.ToString("F") + " %";

                            totalSubTotalReport.ShiftList.Add(totalSubTotalShift);
                        }
                    }

                    decimal totalShiftx1 = planSubTotalReport.ShiftList.Sum(x => x.Shift1);
                    decimal totalShiftx2 = planSubTotalReport.ShiftList.Sum(x => x.Shift2);
                    decimal totalShiftx3 = planSubTotalReport.ShiftList.Sum(x => x.Shift3);
                    decimal totalAllShiftx = planSubTotalReport.ShiftList.Sum(x => x.AllShift);

                    planSubTotalReport.Total = granularity == "Daily" ? totalShiftx1 + totalShiftx2 + totalShiftx3 : totalAllShiftx;

                    totalShiftx1 = actualSubTotalReport.ShiftList.Sum(x => x.Shift1);
                    totalShiftx2 = actualSubTotalReport.ShiftList.Sum(x => x.Shift2);
                    totalShiftx3 = actualSubTotalReport.ShiftList.Sum(x => x.Shift3);
                    totalAllShiftx = actualSubTotalReport.ShiftList.Sum(x => x.AllShift);

                    actualSubTotalReport.Total = granularity == "Granularity" ? totalShiftx1 + totalShiftx2 + totalShiftx3 : totalAllShiftx;

                    totalSubTotalReport.Total = actualSubTotalReport.Total == 0 || planSubTotalReport.Total == 0 ? 0 : (actualSubTotalReport.Total / planSubTotalReport.Total) * 100;
                    totalSubTotalReport.TotalStr = totalSubTotalReport.Total == 0 ? "" : totalSubTotalReport.Total.ToString("F") + " %";

                    model.ShiftReports.Add(planSubTotalReport);
                    model.ShiftReports.Add(actualSubTotalReport);
                    model.ShiftReports.Add(totalSubTotalReport);

                    model.ShiftReportBlocks.Add(new ShiftReportBlock { Plan = planSubTotalReport, Actual = actualSubTotalReport, Total = totalSubTotalReport });
                }
            }

            return model;
        }

        private ReportProdActualPlanningModel GetWppByBrand(string productType, DateTime startDateFL, DateTime endDateFL, List<string> weekList, List<long> locIDList, string granularity, List<string> locList = null)
        {
            ReportProdActualPlanningModel model = new ReportProdActualPlanningModel();
            model.Granularity = granularity;

            List<WppStpModel> wpps = GetWPPList(startDateFL, endDateFL, productType, locIDList);

            List<string> locationList = wpps.Select(x => x.Location).Distinct().ToList();
            locList = locationList;

            foreach (var location in locationList)
            {
                List<string> itemCodeList = wpps.Where(x => x.Location == location).OrderBy(x => x.Date).ThenByDescending(x => x.Shift1).ThenByDescending(x => x.Shift2).ThenByDescending(x => x.Shift3).Select(x => x.Brand).Distinct().ToList();
                foreach (var itemCode in itemCodeList)
                {
                    var wpp = wpps.Where(x => x.Brand == itemCode).First();

                    ReportActualPlanningShiftModel planReport = new ReportActualPlanningShiftModel
                    {
                        Location = location,
                        ItemCode = wpp.Brand,
                        Description = wpp.Description,
                        Type = "Plan",
                        UOM = WebConstants.DEFAULT_STICK,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    ReportActualPlanningShiftModel actualReport = new ReportActualPlanningShiftModel
                    {
                        Location = location,
                        ItemCode = wpp.Brand,
                        Description = wpp.Description,
                        Type = "Actual",
                        UOM = WebConstants.DEFAULT_STICK,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    ReportActualPlanningShiftModel totalReport = new ReportActualPlanningShiftModel
                    {
                        Location = location,
                        ItemCode = wpp.Brand,
                        Description = wpp.Description,
                        Type = "% (Actual / Plan)",
                        UOM = WebConstants.DEFAULT_STICK,
                        Remark = string.Empty,
                        ShiftList = new List<ShiftModel>()
                    };

                    if (granularity == "Shiftly" || granularity == "Daily")
                    {
                        for (var day = startDateFL.Date; day.Date <= endDateFL.Date; day = day.AddDays(1))
                        {
                            var wppTemp = wpps.Where(x => x.Date.Date == day.Date && x.Brand == itemCode && x.Location == location).ToList();

                            ShiftModel planShift = new ShiftModel
                            {
                                Date = day,
                                Shift1 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1),
                                Shift2 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift2),
                                Shift3 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift3),
                                AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift2) + wppTemp.Sum(x => x.Shift3)
                            };
                            planReport.ShiftList.Add(planShift);

                            ShiftModel actualShift = new ShiftModel
                            {
                                Date = day,
                                Shift1 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1),
                                Shift2 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual2),
                                Shift3 = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual3),
                                AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1) + wppTemp.Sum(x => x.Actual2) + wppTemp.Sum(x => x.Actual3)
                            };
                            actualReport.ShiftList.Add(actualShift);

                            if (wppTemp.Count == 0)
                            {
                                ShiftModel totalShift = new ShiftModel
                                {
                                    Date = day,
                                    Shift1 = 0,
                                    Shift2 = 0,
                                    Shift3 = 0,
                                    AllShift = 0
                                };

                                totalShift.Shift1Str = "";
                                totalShift.Shift2Str = "";
                                totalShift.Shift3Str = "";
                                totalShift.AllShiftStr = "";

                                totalReport.ShiftList.Add(totalShift);
                            }
                            else
                            {
                                ShiftModel totalShift = new ShiftModel
                                {
                                    Date = day,
                                    Shift1 = wppTemp.Sum(x => x.Actual1) == 0 || wppTemp.Sum(x => x.Shift1) == 0 ? 0 : (wppTemp.Sum(x => x.Actual1) / wppTemp.Sum(x => x.Shift1)) * 100,
                                    Shift2 = wppTemp.Sum(x => x.Actual2) == 0 || wppTemp.Sum(x => x.Shift2) == 0 ? 0 : (wppTemp.Sum(x => x.Actual2) / wppTemp.Sum(x => x.Shift2)) * 100,
                                    Shift3 = wppTemp.Sum(x => x.Actual3) == 0 || wppTemp.Sum(x => x.Shift3) == 0 ? 0 : (wppTemp.Sum(x => x.Actual3) / wppTemp.Sum(x => x.Shift3)) * 100,
                                    AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                                };
                                totalShift.Shift1Str = totalShift.Shift1 == 0 ? "" : totalShift.Shift1.ToString("F") + " %";
                                totalShift.Shift2Str = totalShift.Shift2 == 0 ? "" : totalShift.Shift2.ToString("F") + " %";
                                totalShift.Shift3Str = totalShift.Shift3 == 0 ? "" : totalShift.Shift3.ToString("F") + " %";
                                totalShift.AllShiftStr = totalShift.AllShift == 0 ? "" : totalShift.AllShift.ToString("F") + " %";

                                totalReport.ShiftList.Add(totalShift);
                            }
                        }
                    }
                    else
                    {
                        DateTime startDate;
                        DateTime endDate;

                        int index = 0;
                        foreach (var weekNumber in weekList)
                        {
                            startDate = startDateFL.AddDays(index++ * 7);
                            endDate = startDate.AddDays(6);

                            var wppTemp = wpps.Where(x => x.Date.Date >= startDate.Date && x.Date.Date <= endDate.Date && x.Brand == itemCode && x.Location == location).ToList();

                            ShiftModel planShift = new ShiftModel
                            {
                                AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift1) + wppTemp.Sum(x => x.Shift1)
                            };
                            planReport.ShiftList.Add(planShift);

                            ShiftModel actualShift = new ShiftModel
                            {
                                AllShift = wppTemp.Count == 0 ? 0 : wppTemp.Sum(x => x.Actual1) + wppTemp.Sum(x => x.Actual2) + wppTemp.Sum(x => x.Actual3)
                            };
                            actualReport.ShiftList.Add(actualShift);

                            if (wppTemp.Count == 0)
                            {
                                ShiftModel totalShift = new ShiftModel
                                {
                                    AllShift = 0
                                };

                                totalShift.AllShiftStr = "";

                                totalReport.ShiftList.Add(totalShift);
                            }
                            else
                            {
                                ShiftModel totalShift = new ShiftModel
                                {
                                    AllShift = actualShift.AllShift == 0 || planShift.AllShift == 0 ? 0 : (actualShift.AllShift / planShift.AllShift) * 100,
                                };

                                totalShift.AllShiftStr = totalShift.AllShift == 0 ? "" : totalShift.AllShift.ToString("F") + " %";

                                totalReport.ShiftList.Add(totalShift);
                            }
                        }
                    }

                    decimal totalShift1 = planReport.ShiftList.Sum(x => x.Shift1);
                    decimal totalShift2 = planReport.ShiftList.Sum(x => x.Shift2);
                    decimal totalShift3 = planReport.ShiftList.Sum(x => x.Shift3);
                    decimal totalAllShift = planReport.ShiftList.Sum(x => x.AllShift);

                    planReport.Total = totalAllShift;

                    totalShift1 = actualReport.ShiftList.Sum(x => x.Shift1);
                    totalShift2 = actualReport.ShiftList.Sum(x => x.Shift2);
                    totalShift3 = actualReport.ShiftList.Sum(x => x.Shift3);
                    totalAllShift = actualReport.ShiftList.Sum(x => x.AllShift);

                    actualReport.Total = totalAllShift;

                    totalReport.Total = actualReport.Total == 0 || planReport.Total == 0 ? 0 : (actualReport.Total / planReport.Total) * 100;
                    totalReport.TotalStr = totalReport.Total == 0 ? "" : totalReport.Total.ToString("F") + " %";

                    model.ShiftReports.Add(planReport);
                    model.ShiftReports.Add(actualReport);
                    model.ShiftReports.Add(totalReport);
                }
            }

            return model;
        }

        private List<WppStpModel> GetWPPList(DateTime startDate, DateTime endDateFL, string productType, List<long> locIDList)
        {
            // Filter Search
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Date", startDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
            filters.Add(new QueryFilter("Date", endDateFL.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));

            string wppList = _wppStpAppService.Find(filters);
            List<WppStpModel> wpps = wppList.DeserializeToWppStpList().Where(x => locIDList.Any(y => y == x.LocationID)).OrderBy(x => x.Date).ToList();

            if (productType == "Finish")
            {
                wpps = wpps.Where(x => x.Brand.StartsWith("F") && !x.IsDeleted).ToList();
            }
            else if (productType == "Intermediate")
            {
                wpps = wpps.Where(x => !x.Brand.StartsWith("F") && !x.IsDeleted).ToList();
            }
            else if (productType == "Primary")
            {
                wpps = wpps.Where(x => x.IsDeleted).ToList();
            }

            return wpps;
        }
        #endregion
    }
}
