using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Threading;
using System.Globalization;
using ReportModel = Fast.Web.Models.Report.ReportModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("reportrsd")]
    public class ShiftDailyController : BaseController<LPHModel>
    {
        private readonly IReportRemarksAppService _reportRemarksAppService;
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBrandAppService _brandAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly IInputTargetAppService _inputTargetAppService;
        private readonly IInputDailyAppService _inputDailyAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly IMachineAppService _machineAppService;
        public ShiftDailyController(
          IReportRemarksAppService reportRemarksService,
          ILPHAppService lphAppService,
          IReferenceAppService referenceAppService,
          ILocationAppService locationAppService,
          ILoggerAppService logger,
          IBrandAppService brandAppService,
          ILPHSubmissionsAppService lPHSubmissionsAppService,
          ILPHApprovalsAppService lPHApprovalsAppService,
          IInputTargetAppService inputTargetAppService,
          IInputDailyAppService inputDailyAppService,
          IReferenceDetailAppService referenceDetailAppService,
          IMachineAppService machineAppService)
        {
            _reportRemarksAppService = reportRemarksService;
            _lphAppService = lphAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _logger = logger;
            _brandAppService = brandAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _lphApprovalAppService = lPHApprovalsAppService;
            _inputTargetAppService = inputTargetAppService;
            _inputDailyAppService = inputDailyAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _machineAppService = machineAppService;
        }
        private class ShiftlyDailyModel : BaseModel
        {
            public string type { get; set; }
            public decimal trend { get; set; }
            public ShiftlyDailyDayModel sum { get; set; }
            public decimal WTD { get; set; }
            public decimal MTD { get; set; }
            public decimal MTG { get; set; }
            public ShiftlyDailyDayModel Shift1 { get; set; }
            public ShiftlyDailyDayModel Shift2 { get; set; }
            public ShiftlyDailyDayModel Shift3 { get; set; }
            public decimal Target { get; set; }
        }
        private class ShiftlyDailyDayModel : BaseModel
        {
            public decimal H { get; set; }
            public decimal H1 { get; set; }
            public decimal H2 { get; set; }
        }
        public ActionResult Index()
        {

            ExecuteQuery(@"UPDATE InputDailies SET[MTBFValue] = 0 WHERE[MTBFValue] IS NULL;
                            UPDATE InputDailies SET[CPQIValue] = 0 WHERE[CPQIValue] IS NULL;
                            UPDATE InputDailies SET[VQIValue] = 0 WHERE[VQIValue] IS NULL;
                            UPDATE InputDailies SET[LocationID] = 0 WHERE[LocationID] IS NULL;
                            UPDATE InputDailies SET [ModifiedDate] = GETDATE() WHERE [ModifiedDate] IS NULL; ");
            GetTempData();
            ViewBag.CountryList = DropDownHelper.BindDropDownCountry(_referenceAppService);
            // ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            //ViewBag.LinkUpList = GetLinkUpList();
            ViewBag.ProductionCenterList = GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, true);
            return View();
        }
        public void ExecuteQuery(string query)
        {
            string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(strConString))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    SqlCommand command = new SqlCommand(query, connection, transaction);
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                }
            }
        }
        [HttpPost]
        public ActionResult GetLinkUpList(string prodcenterID) {

            //Dictionary<string, List<string>> resultList = new Dictionary<string, List<string>>();
            ////Get Prodcenter list
            //string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            //List<LocationModel> prodcenterList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            //List<QueryFilter> filters = new List<QueryFilter>();
            //filters.Add(new QueryFilter("LinkUp", null, Operator.NotEqual));
            //filters.Add(new QueryFilter("LinkUp", "", Operator.NotEqual));
            //filters.Add(new QueryFilter("IsDeleted", "0"));

            List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList().Where(x => x.LinkUp != null && x.LinkUp != "").ToList();
            //foreach (LocationModel lm in prodcenterList)
            //{
            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(long.Parse(prodcenterID), "productioncenter");
            List<string> listLinkUp = machineList.Where(x =>
                    locationIdList.Any(y => y == x.LocationID)
                ).Select(x => x.LinkUp).Distinct().ToList();
                //resultList.Add(lm.ID.ToString(), listLinkUp);
            //}
            return Json(new { Status = "True", Data = listLinkUp }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult SendEmail()
        {
            //try
            //{
            //    return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            //}
            //catch (Exception ex)
            //{
            //    ViewBag.Result = false;
            //    _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            //    return Json(new { Error = ex.Message });
            //}
            return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult GetUptimeDefault()
        {
            try
            {
                DateTime dateStart = DateTime.Now.AddDays(-2);
                string month = DateTime.Now.ToString("MMM").ToUpper();

                ICollection<QueryFilter> filterUptime = new List<QueryFilter>();
                filterUptime = new List<QueryFilter>();
                filterUptime.Add(new QueryFilter("IsDeleted", "0"));
                filterUptime.Add(new QueryFilter("ProdCenterID", AccountProdCenterID));
                filterUptime.Add(new QueryFilter("Date", DateTime.Now.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

                string uptimeDaily = _inputDailyAppService.Find(filterUptime);
                List<InputDailyModel> uptimeDailyRes = uptimeDaily.DeserializeToInputDailyList();

                List<UptimeModel> listUptime = uptimeDailyRes.AsEnumerable()
                          .Select(o => new UptimeModel
                          {
                              Shift = o.Shift,
                              UptimeValue = o.UptimeValue,
                              UptimeFocus = o.UptimeFocus,
                              UptimeActPlan = o.UptimeActPlan
                          }).ToList();


                return Json(new { Status = "True", ResultUptime = listUptime }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Error = ex.Message });
            }
        }
        [HttpPost]
        public ActionResult GetUptimeWithParam(DateTime dtFilter, string prodCenter, string machine)
        {
            try
            {
                long pcID = long.Parse(prodCenter);

                string month = dtFilter.ToString("MMM").ToUpper();

                DateTime dateStart = dtFilter.AddDays(-2);

                ICollection<QueryFilter> filterUptime = new List<QueryFilter>();
                filterUptime = new List<QueryFilter>();
                filterUptime.Add(new QueryFilter("IsDeleted", "0"));
                filterUptime.Add(new QueryFilter("ProdCenterID", pcID));
                filterUptime.Add(new QueryFilter("LinkUp", machine));
                filterUptime.Add(new QueryFilter("Date", dtFilter.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

                string uptimeDaily = _inputDailyAppService.Find(filterUptime);
                List<InputDailyModel> uptimeDailyRes = uptimeDaily.DeserializeToInputDailyList();

                List<UptimeModel> listUptime = uptimeDailyRes.AsEnumerable()
                          .Select(o => new UptimeModel
                          {
                              Shift = o.Shift,
                              UptimeValue = o.UptimeValue,
                              UptimeFocus = o.UptimeFocus,
                              UptimeActPlan = o.UptimeActPlan
                          }).ToList();

                return Json(new { Status = "True", ResultUptime = listUptime }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Error = ex.Message });
            }
        }
        public double CalculateLinest(double[] y)
        {
            double linest = 0;
            double[] x = new double[] { 0, 1, 2 };
            if (y.Length == x.Length)
            {
                double avgY = y.Average();
                double avgX = x.Average();
                double[] dividend = new double[y.Length];
                double[] divisor = new double[y.Length];
                for (int i = 0; i < y.Length; i++)
                {
                    dividend[i] = (x[i] - avgX) * (y[i] - avgY);
                    divisor[i] = Math.Pow((x[i] - avgX), 2);
                }
                linest = dividend.Sum() / divisor.Sum();
            }
            return linest;
        }
        public (List<Dictionary<string, object>>, List<Dictionary<string, string>>) GetDataWithParam(DateTime dtFilter, string prodCenter, string machine)
        {
            List<string> listLinkUp = new List<string>();

            string month = dtFilter.ToString("MMM").ToUpper();
            DateTime dateStart = dtFilter.AddDays(-2);

            if (machine == "All")
            {
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(long.Parse(prodCenter), "productioncenter");
                List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList().Where(x => x.LinkUp != null && x.LinkUp != "").ToList();
                listLinkUp = machineList.Where(x =>
                       locationIdList.Any(y => y == x.LocationID)
                   ).Select(x => x.LinkUp).Distinct().ToList();
                listLinkUp.Insert(0, "All");
            }
            else
            {
                listLinkUp.Add(machine);
            }
            int diff = (7 + (dtFilter.DayOfWeek - DayOfWeek.Monday)) % 7;
            DateTime startOfWeek = dtFilter.AddDays(-1 * diff).Date;
            DateTime firstDayOfMonth = new DateTime(dtFilter.Year, dtFilter.Month, 1);

            ICollection<QueryFilter> filterInputTarget = new List<QueryFilter>();
            filterInputTarget = new List<QueryFilter>();
            filterInputTarget.Add(new QueryFilter("IsDeleted", "0"));
            filterInputTarget.Add(new QueryFilter("ProdCenterID", prodCenter));
            filterInputTarget.Add(new QueryFilter("Month", month));

            string inputTarget = _inputTargetAppService.Find(filterInputTarget);
            List<InputTargetModel> inputTargetRes = inputTarget.DeserializeToInputTargetList();


            var mtbfCheck = inputTargetRes.Where(x => x.KPI == "MTBF").ToList();
            var cpqiCheck = inputTargetRes.Where(x => x.KPI == "CPQI").ToList();
            var vqiCheck = inputTargetRes.Where(x => x.KPI == "VQI").ToList();
            var workingCheck = inputTargetRes.Where(x => x.KPI == "Working Time").ToList();
            var uptimeCheck = inputTargetRes.Where(x => x.KPI == "Uptime").ToList();
            var strsCheck = inputTargetRes.Where(x => x.KPI == "STRS").ToList();
            var prodCheck = inputTargetRes.Where(x => x.KPI == "Production Volume").ToList();
            var crrCheck = inputTargetRes.Where(x => x.KPI == "CRR").ToList();

            List<string> listKPI = new List<string>()
                {
                    "Volume",
                    "Uptime",
                    "MTBF",
                    "CRR",
                    "VQI",
                    "CPQI",
                    "STRS",
                    "Working Time"
                };
            List<string> listVersion = new List<string>()
                {
                    "RF11",
                    "RF09",
                    "RF06",
                    "RF02",
                    "Internal Target",
                    "OB"
                };
            InputTargetModel mtbfTarget = null;
            InputTargetModel cpqiTarget = null;
            InputTargetModel vqiTarget = null;
            InputTargetModel workingTarget = null;
            InputTargetModel uptimeTarget = null;
            InputTargetModel strsTarget = null;
            InputTargetModel prodTarget = null;
            InputTargetModel crrTarget = null;

            foreach (string version in listVersion)
            {
                if (mtbfTarget == null)
                    mtbfTarget = mtbfCheck.Where(x => x.Version == version).FirstOrDefault();

                if (cpqiTarget == null)
                    cpqiTarget = cpqiCheck.Where(x => x.Version == version).FirstOrDefault();

                if (vqiTarget == null)
                    vqiTarget = vqiCheck.Where(x => x.Version == version).FirstOrDefault();

                if (workingTarget == null)
                    workingTarget = workingCheck.Where(x => x.Version == version).FirstOrDefault();

                if (uptimeTarget == null)
                    uptimeTarget = uptimeCheck.Where(x => x.Version == version).FirstOrDefault();

                if (strsTarget == null)
                    strsTarget = strsCheck.Where(x => x.Version == version).FirstOrDefault();

                if (prodTarget == null)
                    prodTarget = prodCheck.Where(x => x.Version == version).FirstOrDefault();

                if (crrTarget == null)
                    crrTarget = crrCheck.Where(x => x.Version == version).FirstOrDefault();
            }


            ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
            filterInputDaily = new List<QueryFilter>();
            filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
            filterInputDaily.Add(new QueryFilter("ProdCenterID", prodCenter));
            filterInputDaily.Add(new QueryFilter("Date", dateStart.AddHours(-12).ToString(), Operator.GreaterThanOrEqual));
            filterInputDaily.Add(new QueryFilter("Date", dtFilter.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

            string inputDaily = _inputDailyAppService.Find(filterInputDaily);
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift != "Daily").ToList();

            List<Dictionary<string, object>> linkupList = new List<Dictionary<string, object>>();

            foreach (string lu in listLinkUp)
            {
                List<ShiftlyDailyModel> dataDic = new List<ShiftlyDailyModel>();

                WTDModel modelWTD = GetWTD(startOfWeek, dtFilter, long.Parse(prodCenter), lu);
                MTDModel modelMTD = GetMTD(firstDayOfMonth, dtFilter, long.Parse(prodCenter), lu);
                List<InputDailyModel> inputDailyLinkUp = new List<InputDailyModel>();
                if (lu.Equals("All"))
                {
                    inputDailyLinkUp = inputDailyRes;
                }
                else
                {
                    inputDailyLinkUp = inputDailyRes.Where(x => x.LinkUp == lu).ToList();
                }
                foreach (string kpi in listKPI)
                {
                    ShiftlyDailyModel typeDic = new ShiftlyDailyModel();
                    typeDic.sum = new ShiftlyDailyDayModel();
                    typeDic.Shift1 = new ShiftlyDailyDayModel();
                    typeDic.Shift2 = new ShiftlyDailyDayModel();
                    typeDic.Shift3 = new ShiftlyDailyDayModel();
                    InputTargetModel thisTarget = null;


                    List<InputDailyModel> inputThisKPIH = inputDailyLinkUp.Where(x => x.Date == dtFilter).ToList();
                    List<InputDailyModel> inputThisKPIHmin1 = inputDailyLinkUp.Where(x => x.Date == dtFilter.AddDays(-1)).ToList();
                    List<InputDailyModel> inputThisKPIHmin2 = inputDailyLinkUp.Where(x => x.Date == dtFilter.AddDays(-2)).ToList();

                    Dictionary<string, object> sumInput = new Dictionary<string, object>();
                    Dictionary<string, object> shiftInput = new Dictionary<string, object>();

                    typeDic.WTD = decimal.Parse(getWTDWithKey(kpi, modelWTD));
                    typeDic.MTD = decimal.Parse(getMTDWithKey(kpi, modelMTD));

                    switch (kpi)
                    {
                        case "Volume":
                            thisTarget = prodTarget != null ? prodTarget : new InputTargetModel { Value = 0 };
                            typeDic.MTG = thisTarget.Value - typeDic.MTD;
                            typeDic.sum.H = inputThisKPIH.Sum(x => x.ProdVolumeValue);
                            typeDic.sum.H1 = inputThisKPIHmin1.Sum(x => x.ProdVolumeValue);
                            typeDic.sum.H2 = inputThisKPIHmin2.Sum(x => x.ProdVolumeValue);
                            typeDic.trend = decimal.Parse(CalculateLinest(new double[] { decimal.ToDouble(typeDic.sum.H2), decimal.ToDouble(typeDic.sum.H1), decimal.ToDouble(typeDic.sum.H) }).ToString());

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Sum(x => x.ProdVolumeValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Sum(x => x.ProdVolumeValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Sum(x => x.ProdVolumeValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Sum(x => x.ProdVolumeValue);

                            break;
                        case "Uptime":
                            thisTarget = uptimeTarget != null ? uptimeTarget : new InputTargetModel { Value = 0 };

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.UptimeValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.UptimeValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.UptimeValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.UptimeValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.UptimeValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.UptimeValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.UptimeValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.UptimeValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.UptimeValue);
                            
                            break;
                        case "MTBF":
                            thisTarget = mtbfTarget != null ? mtbfTarget : new InputTargetModel { Value = 0 };
                            //typeDic.MTG = thisTarget.Value - typeDic.MTD;

                            //typeDic.sum.H = inputThisKPIH.Sum(x => x.MTBFValue);
                            //typeDic.sum.H1 = inputThisKPIHmin1.Sum(x => x.MTBFValue);
                            //typeDic.sum.H2 = inputThisKPIHmin2.Sum(x => x.MTBFValue);
                            //typeDic.trend = decimal.Parse(CalculateLinest(new double[] { decimal.ToDouble(typeDic.sum.H2), decimal.ToDouble(typeDic.sum.H1), decimal.ToDouble(typeDic.sum.H) }).ToString());

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.MTBFValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.MTBFValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.MTBFValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.MTBFValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.MTBFValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.MTBFValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.MTBFValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.MTBFValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.MTBFValue);

                            break;
                        case "CRR":
                            thisTarget = crrTarget != null ? crrTarget : new InputTargetModel { Value = 0 };

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.CRRValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.CRRValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.CRRValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.CRRValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.CRRValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.CRRValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.CRRValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.CRRValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.CRRValue);

                            break;
                        case "VQI":
                            thisTarget = vqiTarget != null ? vqiTarget : new InputTargetModel { Value = 0 };


                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.VQIValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.VQIValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.VQIValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.VQIValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.VQIValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.VQIValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.VQIValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.VQIValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.VQIValue);

                            break;
                        case "CPQI":
                            thisTarget = cpqiTarget != null ? cpqiTarget : new InputTargetModel { Value = 0 };
                            

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.CPQIValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.CPQIValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.CPQIValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.CPQIValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.CPQIValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.CPQIValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.CPQIValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.CPQIValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.CPQIValue);

                            break;
                        case "STRS":
                            thisTarget = strsTarget != null ? strsTarget : new InputTargetModel { Value = 0 };
                           
                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "1").Average(x => x.STRSValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Average(x => x.STRSValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Average(x => x.STRSValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "2").Average(x => x.STRSValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Average(x => x.STRSValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Average(x => x.STRSValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIH.Where(x => x.Shift.Trim() == "3").Average(x => x.STRSValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Average(x => x.STRSValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Count() == 0 ? 0 : inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Average(x => x.STRSValue);

                            break;
                        case "Working Time":
                            thisTarget = workingTarget != null ? workingTarget : new InputTargetModel { Value = 0 };
                            typeDic.MTG = thisTarget.Value - typeDic.MTD;

                            typeDic.sum.H = inputThisKPIH.Sum(x => x.WorkingValue);
                            typeDic.sum.H1 = inputThisKPIHmin1.Sum(x => x.WorkingValue);
                            typeDic.sum.H2 = inputThisKPIHmin2.Sum(x => x.WorkingValue);
                            typeDic.trend = decimal.Parse(CalculateLinest(new double[] { decimal.ToDouble(typeDic.sum.H2), decimal.ToDouble(typeDic.sum.H1), decimal.ToDouble(typeDic.sum.H) }).ToString());

                            typeDic.Shift1.H = inputThisKPIH.Where(x => x.Shift.Trim() == "1").Sum(x => x.WorkingValue);
                            typeDic.Shift1.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "1").Sum(x => x.WorkingValue);
                            typeDic.Shift1.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "1").Sum(x => x.WorkingValue);

                            typeDic.Shift2.H = inputThisKPIH.Where(x => x.Shift.Trim() == "2").Sum(x => x.WorkingValue);
                            typeDic.Shift2.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "2").Sum(x => x.WorkingValue);
                            typeDic.Shift2.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "2").Sum(x => x.WorkingValue);

                            typeDic.Shift3.H = inputThisKPIH.Where(x => x.Shift.Trim() == "3").Sum(x => x.WorkingValue);
                            typeDic.Shift3.H1 = inputThisKPIHmin1.Where(x => x.Shift.Trim() == "3").Sum(x => x.WorkingValue);
                            typeDic.Shift3.H2 = inputThisKPIHmin2.Where(x => x.Shift.Trim() == "3").Sum(x => x.WorkingValue);

                            break;
                    }

                    decimal target = thisTarget.Value;
                    typeDic.type = kpi;
                    typeDic.Target = target;
                    dataDic.Add(typeDic);
                }

                ShiftlyDailyModel volume = dataDic.Where(x => x.type == "Volume").FirstOrDefault();
                ShiftlyDailyDayModel volumeS1 = volume.Shift1;
                ShiftlyDailyDayModel volumeS2 = volume.Shift2;
                ShiftlyDailyDayModel volumeS3 = volume.Shift3;
                ShiftlyDailyDayModel volumeSum = volume.sum;
                ShiftlyDailyModel uptime = dataDic.Where(x => x.type == "Uptime").FirstOrDefault();
                ShiftlyDailyDayModel uptimeS1 = uptime.Shift1;
                ShiftlyDailyDayModel uptimeS2 = uptime.Shift2;
                ShiftlyDailyDayModel uptimeS3 = uptime.Shift3;

                //H
                decimal TheorVals1 = uptimeS1.H == 0 ? 0 : volumeS1.H / uptimeS1.H * 100;
                decimal TheorVals2 = uptimeS2.H == 0 ? 0 : volumeS2.H / uptimeS2.H * 100;
                decimal TheorVals3 = uptimeS3.H == 0 ? 0 : volumeS3.H / uptimeS3.H * 100;
                decimal TheorSumH = TheorVals1 + TheorVals2 + TheorVals3;
                decimal sumUptimeH = TheorSumH == 0 ? 0 : volumeSum.H / TheorSumH * 100;
                //H
                TheorVals1 = uptimeS1.H1 == 0 ? 0 : volumeS1.H1 / uptimeS1.H1 * 100;
                TheorVals2 = uptimeS2.H1 == 0 ? 0 : volumeS2.H1 / uptimeS2.H1 * 100;
                TheorVals3 = uptimeS3.H1 == 0 ? 0 : volumeS3.H1 / uptimeS3.H1 * 100;
                decimal TheorSumH1 = TheorVals1 + TheorVals2 + TheorVals3;
                decimal sumUptimeH1 = TheorSumH1 == 0 ? 0 : volumeSum.H1 / TheorSumH1 * 100;
                //H
                TheorVals1 = uptimeS1.H2 == 0 ? 0 : volumeS1.H2 / uptimeS1.H2 * 100;
                TheorVals2 = uptimeS2.H2 == 0 ? 0 : volumeS2.H2 / uptimeS2.H2 * 100;
                TheorVals3 = uptimeS3.H2 == 0 ? 0 : volumeS3.H2 / uptimeS3.H2 * 100;
                decimal TheorSumH2 = TheorVals1 + TheorVals2 + TheorVals3;
                decimal sumUptimeH2 = TheorSumH2 == 0 ? 0 : volumeSum.H2 / TheorSumH2 * 100;

                decimal mtdTheor = uptime.MTD == 0 ? 0 : volume.MTD / uptime.MTD * 100;
                decimal targetTheor = uptime.Target == 0 ? 0 : volume.Target / uptime.Target * 100;
                decimal mtgTheor = targetTheor - mtdTheor;


                dataDic.Where(x => x.type == "Uptime").FirstOrDefault().sum = new ShiftlyDailyDayModel { H = sumUptimeH, H1 = sumUptimeH1, H2 = sumUptimeH2 };
                dataDic.Where(x => x.type == "Uptime").FirstOrDefault().MTG = mtgTheor == 0 ? 0 : volume.MTG / mtgTheor * 100;

                dataDic.Where(x => x.type == "Uptime").FirstOrDefault().trend = decimal.Parse(CalculateLinest(new double[] { decimal.ToDouble(sumUptimeH2), decimal.ToDouble(sumUptimeH1), decimal.ToDouble(sumUptimeH) }).ToString());

                List<string> specialSUM = new List<string>() { "CPQI", "VQI", "STRS", "CRR", "MTBF" };
                foreach (string kpi in specialSUM)
                {
                    ShiftlyDailyModel CPQI = dataDic.Where(x => x.type == kpi).FirstOrDefault();
                    decimal sumCPQIH = volumeSum.H == 0 ? 0 : SumProduct(new decimal[] { CPQI.Shift1.H, CPQI.Shift2.H, CPQI.Shift3.H }, new decimal[] { volume.Shift1.H, volume.Shift2.H, volume.Shift3.H }) / volumeSum.H;
                    decimal sumCPQIH1 = volumeSum.H1 == 0 ? 0 : SumProduct(new decimal[] { CPQI.Shift1.H1, CPQI.Shift2.H1, CPQI.Shift3.H1 }, new decimal[] { volume.Shift1.H1, volume.Shift2.H1, volume.Shift3.H1 }) / volumeSum.H1;
                    decimal sumCPQIH2 = volumeSum.H2 == 0 ? 0 : SumProduct(new decimal[] { CPQI.Shift1.H2, CPQI.Shift2.H2, CPQI.Shift3.H2 }, new decimal[] { volume.Shift1.H2, volume.Shift2.H2, volume.Shift3.H2 }) / volumeSum.H2;
                    dataDic.Where(x => x.type == kpi).FirstOrDefault().sum = new ShiftlyDailyDayModel { H = sumCPQIH, H1 = sumCPQIH1, H2 = sumCPQIH2 };
                    dataDic.Where(x => x.type == kpi).FirstOrDefault().trend = decimal.Parse(CalculateLinest(new double[] { decimal.ToDouble(sumCPQIH2), decimal.ToDouble(sumCPQIH1), decimal.ToDouble(sumCPQIH) }).ToString());
                    decimal valMTG = volume.MTG == 0 ? 0 : ((volume.Target * CPQI.Target) - (volume.MTD * CPQI.MTD)) / volume.MTG;
                    dataDic.Where(x => x.type == kpi).FirstOrDefault().MTG = valMTG;
                }
                Dictionary<string, object> linkupDic = new Dictionary<string, object>();
                linkupDic.Add("linkup", lu);
                linkupDic.Add("data", dataDic);
                linkupList.Add(linkupDic);


            }


            //Get Remark

            List<InputDailyModel> inputThisDay = inputDailyRes.Where(x => x.Date == dtFilter).OrderBy(x => x.Shift).ToList();

            if (machine != "All")
                inputThisDay = inputThisDay.Where(x => x.LinkUp == machine).OrderBy(x => x.Shift).ToList();

            List<Dictionary<string, string>> remarkList = new List<Dictionary<string, string>>();
            foreach (InputDailyModel idm in inputThisDay)
            {
                foreach (string kpi in listKPI)
                {
                    Dictionary<string, string> lineDic = new Dictionary<string, string>();
                    lineDic.Add("shift", idm.Shift);
                    lineDic.Add("type", kpi);
                    switch (kpi)
                    {
                        case "Volume":
                            lineDic.Add("remark", idm.ProdVolumeFocus);
                            lineDic.Add("actionplan", idm.ProdVolumeActPlan);
                            break;
                        case "Uptime":
                            lineDic.Add("remark", idm.UptimeFocus);
                            lineDic.Add("actionplan", idm.UptimeActPlan);
                            break;
                        case "MTBF":
                            lineDic.Add("remark", idm.MTBFFocus);
                            lineDic.Add("actionplan", idm.MTBFActPlan);
                            break;
                        case "CRR":
                            lineDic.Add("remark", idm.CRRFocus);
                            lineDic.Add("actionplan", idm.CRRActPlan);
                            break;
                        case "VQI":
                            lineDic.Add("remark", idm.VQIFocus);
                            lineDic.Add("actionplan", idm.VQIActPlan);
                            break;
                        case "CPQI":
                            lineDic.Add("remark", idm.CPQIFocus);
                            lineDic.Add("actionplan", idm.CPQIActPlan);
                            break;
                        case "STRS":
                            lineDic.Add("remark", idm.STRSFocus);
                            lineDic.Add("actionplan", idm.STRSActPlan);
                            break;
                        case "Working Time":
                            lineDic.Add("remark", idm.WorkingFocus);
                            lineDic.Add("actionplan", idm.WorkingActPlan);
                            break;
                    }

                    remarkList.Add(lineDic);
                }
            }
            return (linkupList, remarkList);

        }
        decimal SumProduct(decimal[] arrayA, decimal[] arrayB)
        {
            decimal result = 0;
            for (int i = 0; i < arrayA.Count(); i++)
                result += arrayA[i] * arrayB[i];
            return result;
        }
        [HttpPost]
        public ActionResult GetReportWithParam(DateTime dtFilter, string prodCenter, string machine)
        {
            try
            {
                var (linkupList, remarkList) = GetDataWithParam(dtFilter, prodCenter, machine);
                
                return Json(new { Status = "True", linkupList, remarkList }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False", Error = ex.Message });
            }
        }
        public ActionResult GenerateShiftlyDailyWithParam(DateTime dtFilter, string prodCenter, string machine)
        {
            try
            {
                System.Drawing.Color colFromHex = ColorTranslator.FromHtml("#99ccff");
                string[] higherBetter = new string[] { "MTBF", "Uptime", "STRS", "Volume" };
                string[] lowerBetter = new string[] { "CRR", "VQI", "CPQI" };
                var (linkupList, remarkList) = GetDataWithParam(dtFilter, prodCenter, machine);
                ExcelPackage Ep = new ExcelPackage();

                foreach (Dictionary<string, object> linkup in linkupList)
                {
                    ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(linkup["linkup"].ToString());
                    int rowPos = 1;

                    using (var range = Sheet.Cells[rowPos, 1, rowPos, 9])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }
                    Sheet.Cells[rowPos, 1].Value = "KPI";
                    Sheet.Cells[rowPos, 2].Value = "Trend";
                    Sheet.Cells[rowPos, 3].Value = dtFilter.AddDays(-2).ToString("dd-MMM-yyyy");
                    Sheet.Cells[rowPos, 4].Value = dtFilter.AddDays(-1).ToString("dd-MMM-yyyy");
                    Sheet.Cells[rowPos, 5].Value = dtFilter.ToString("dd-MMM-yyyy");
                    Sheet.Cells[rowPos, 6].Value = "WTD";
                    Sheet.Cells[rowPos, 7].Value = "MTD";
                    Sheet.Cells[rowPos, 8].Value = "MTG";
                    Sheet.Cells[rowPos++, 9].Value = "Target";
                    foreach (ShiftlyDailyModel data in (List<ShiftlyDailyModel>)linkup["data"])
                    {
                        // ▲ ► ▼
                        string satuan = " (%)";
                        switch (data.type)
                        {
                            case "Volume":
                                satuan = " (Mio Stick)";
                                break;
                            case "MTBF":
                                satuan = " (Min)";
                                break;
                            case "VQI":
                            case "CPQI":
                                satuan = " (Points)";
                                break;
                            case "Working Time":
                                satuan = " (Hours)";
                                break;
                        }
                        string type = data.type;
                        double h2 = (double)data.sum.H2;
                        double h1 = (double)data.sum.H1;
                        double h = (double)data.sum.H;
                        double target = (double)data.Target;
                        string trendColor = "#000";
                        if (higherBetter.Any(x => x == type)) trendColor = target == h ? "#000" : target > h ? "#F00" : "#0F0";
                        else if (lowerBetter.Any(x => x == type)) trendColor = target == h ? "#000" : target < h ? "#F00" : "#0F0";
                        String valueFormat = "{0:N}";
                        if (data.type.Equals("Volume"))
                        {
                            valueFormat = "{0:N0}";
                        }
                        double trend = CalculateLinest(new double[] { h2, h1, h });
                        Sheet.Cells[rowPos, 1].Value = type + satuan;
                        Sheet.Cells[rowPos, 2].Value = trend > 0 ? "▲" : trend < 0 ? "▼" : "►";
                        Sheet.Cells[rowPos, 2].Style.Font.Color.SetColor(ColorTranslator.FromHtml(trendColor));
                        Sheet.Cells[rowPos, 3].Value = String.Format(valueFormat, h2);
                        Sheet.Cells[rowPos, 4].Value = String.Format(valueFormat, h1);
                        Sheet.Cells[rowPos, 5].Value = String.Format(valueFormat, h);
                        Sheet.Cells[rowPos, 6].Value = String.Format(valueFormat, data.WTD);
                        Sheet.Cells[rowPos, 7].Value = String.Format(valueFormat, data.MTD);
                        Sheet.Cells[rowPos, 8].Value = String.Format(valueFormat, data.MTG);
                        Sheet.Cells[rowPos++, 9].Value = String.Format(valueFormat, data.Target);

                        Sheet.Cells[rowPos, 2].Value = "Shift 1";
                        Sheet.Cells[rowPos, 3].Value = String.Format(valueFormat, data.Shift1.H2);
                        Sheet.Cells[rowPos, 4].Value = String.Format(valueFormat, data.Shift1.H1);
                        Sheet.Cells[rowPos++, 5].Value = String.Format(valueFormat, data.Shift1.H);

                        Sheet.Cells[rowPos, 2].Value = "Shift 2";
                        Sheet.Cells[rowPos, 3].Value = String.Format(valueFormat, data.Shift2.H2);
                        Sheet.Cells[rowPos, 4].Value = String.Format(valueFormat, data.Shift2.H1);
                        Sheet.Cells[rowPos++, 5].Value = String.Format(valueFormat, data.Shift2.H);

                        Sheet.Cells[rowPos, 2].Value = "Shift 3";
                        Sheet.Cells[rowPos, 3].Value = String.Format(valueFormat, data.Shift3.H2);
                        Sheet.Cells[rowPos, 4].Value = String.Format(valueFormat, data.Shift3.H1);
                        Sheet.Cells[rowPos++, 5].Value = String.Format(valueFormat, data.Shift3.H);
                    }

                    Sheet.Cells[1, 1, rowPos - 1, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[1, 1, rowPos - 1, 9].AutoFitColumns();
                }
                ExcelWorksheet remarkSheet = Ep.Workbook.Worksheets.Add("Remarks");
                int remarkRowPos = 1;

                using (var range = remarkSheet.Cells[remarkRowPos, 1, remarkRowPos, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                remarkSheet.Cells[remarkRowPos, 1].Value = "Shift";
                remarkSheet.Cells[remarkRowPos, 2].Value = "Type";
                remarkSheet.Cells[remarkRowPos, 3].Value = "Focus";
                remarkSheet.Cells[remarkRowPos++, 4].Value = "Action Plan";
                foreach (Dictionary<string, string> remark in remarkList)
                {
                    if (remark["remark"] != null && remark["remark"] != "" && remark["remark"] != "0")
                    {
                        remarkSheet.Cells[remarkRowPos, 1].Value = remark["shift"];
                        remarkSheet.Cells[remarkRowPos, 2].Value = remark["type"];
                        remarkSheet.Cells[remarkRowPos, 3].Value = remark["remark"];
                        remarkSheet.Cells[remarkRowPos++, 4].Value = remark["actionplan"];
                    }
                }
                remarkSheet.Cells[1, 1, remarkRowPos - 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                remarkSheet.Cells[1, 1, remarkRowPos - 1, 4].AutoFitColumns();
                byte[] dataExcel = Ep.GetAsByteArray();
                //return File(dataExcel, "application/octet-stream", "ReportShiftlyDaily.xlsx");
                Session["DownloadExcel_ShiftDaily"] = dataExcel;
                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        private string getWTDWithKey(string key, WTDModel model)
        {
            switch (key)
            {
                case "Volume":
                    return model.ProdVolume_WTD;
                case "Uptime":
                    return model.Uptime_WTD;
                case "MTBF":
                    return model.MTBF_WTD;
                case "CRR":
                    return model.CRR_WTD;
                case "VQI":
                    return model.VQI_WTD;
                case "CPQI":
                    return model.CPQI_WTD;
                case "STRS":
                    return model.STRS_WTD;
                case "Working Time":
                    return model.Working_WTD;
                default:
                    return "0";
            }
        }
        private string getMTDWithKey(string key, MTDModel model)
        {
            switch (key)
            {
                case "Volume":
                    return model.ProdVolume_MTD;
                case "Uptime":
                    return model.Uptime_MTD;
                case "MTBF":
                    return model.MTBF_MTD;
                case "CRR":
                    return model.CRR_MTD;
                case "VQI":
                    return model.VQI_MTD;
                case "CPQI":
                    return model.CPQI_MTD;
                case "STRS":
                    return model.STRS_MTD;
                case "Working Time":
                    return model.Working_MTD;
                default:
                    return "0";
            }
        }
        public ActionResult GetReportDefault()
        {
            var model = new DailyModel();
            DateTime dateStart = DateTime.Now.AddDays(-2);
            string month = DateTime.Now.ToString("MMM").ToUpper();

            DateTime startOfWeek = DateTime.Now.StartOfWeek();
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            WTDModel modelWTD = GetWTD(startOfWeek, DateTime.Now, AccountProdCenterID, "");
            MTDModel modelMTD = GetMTD(firstDayOfMonth, DateTime.Now, AccountProdCenterID, "");


            ICollection<QueryFilter> filterInputTarget = new List<QueryFilter>();
            filterInputTarget = new List<QueryFilter>();
            filterInputTarget.Add(new QueryFilter("IsDeleted", "0"));
            filterInputTarget.Add(new QueryFilter("ProdCenterID", AccountProdCenterID));
            filterInputTarget.Add(new QueryFilter("Month", month));

            string inputTarget = _inputTargetAppService.Find(filterInputTarget);
            List<InputTargetModel> inputTargetRes = inputTarget.DeserializeToInputTargetList();

            List<TargetModel> listTarget = new List<TargetModel>();

            var mtbfCheck = inputTargetRes.Where(x => x.KPI == "MTBF").ToList();
            var cpqiCheck = inputTargetRes.Where(x => x.KPI == "CPQI").ToList();
            var vqiCheck = inputTargetRes.Where(x => x.KPI == "VQI").ToList();
            var workingCheck = inputTargetRes.Where(x => x.KPI == "Working Time").ToList();
            var uptimeCheck = inputTargetRes.Where(x => x.KPI == "Uptime").ToList();
            var strsCheck = inputTargetRes.Where(x => x.KPI == "STRS").ToList();
            var prodCheck = inputTargetRes.Where(x => x.KPI == "Production Volume").ToList();
            var crrCheck = inputTargetRes.Where(x => x.KPI == "CRR").ToList();

            if (mtbfCheck.Count() > 0)
            {
                var rf11_M = mtbfCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = mtbfCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = mtbfCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = mtbfCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = mtbfCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = mtbfCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }
            if (cpqiCheck.Count() > 0)
            {
                var rf11_M = cpqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = cpqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = cpqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = cpqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = cpqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = cpqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }

            }
            if (vqiCheck.Count() > 0)
            {
                var rf11_M = vqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = vqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = vqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = vqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = vqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = vqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }
            if (workingCheck.Count() > 0)
            {
                var rf11_M = workingCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = workingCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = workingCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = workingCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = workingCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = workingCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }
            if (uptimeCheck.Count() > 0)
            {
                var rf11_M = uptimeCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = uptimeCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = uptimeCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = uptimeCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = uptimeCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = uptimeCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }

            }
            if (strsCheck.Count() > 0)
            {
                var rf11_M = strsCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = strsCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = strsCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = strsCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = strsCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = strsCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }
            if (prodCheck.Count() > 0)
            {
                var rf11_M = prodCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = prodCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = prodCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = prodCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = prodCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = prodCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }
            if (crrCheck.Count() > 0)
            {
                var rf11_M = crrCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                if (rf11_M == null)
                {
                    var rf09_M = crrCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                    if (rf09_M == null)
                    {
                        var rf06_M = crrCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                        if (rf06_M == null)
                        {
                            var rf02_M = crrCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                            if (rf02_M == null)
                            {
                                var rfInternal_M = crrCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                if (rfInternal_M == null)
                                {
                                    var rfOB_M = crrCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                    if (rfOB_M == null)
                                    {
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfOB_M.KPI,
                                            Version = rfOB_M.Version,
                                            ValueTarget = rfOB_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rfInternal_M.KPI,
                                        Version = rfInternal_M.Version,
                                        ValueTarget = rfInternal_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf02_M.KPI,
                                    Version = rf02_M.Version,
                                    ValueTarget = rf02_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf06_M.KPI,
                                Version = rf06_M.Version,
                                ValueTarget = rf06_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf09_M.KPI,
                            Version = rf09_M.Version,
                            ValueTarget = rf09_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                else
                {
                    TargetModel mTarget = new TargetModel
                    {
                        KPI = rf11_M.KPI,
                        Version = rf11_M.Version,
                        ValueTarget = rf11_M.Value
                    };
                    listTarget.Add(mTarget);
                }
            }



            ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
            filterInputDaily = new List<QueryFilter>();
            filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
            filterInputDaily.Add(new QueryFilter("ProdCenterID", AccountProdCenterID));
            filterInputDaily.Add(new QueryFilter("Date", dateStart.AddHours(-12).ToString(), Operator.GreaterThanOrEqual));
            filterInputDaily.Add(new QueryFilter("Date", DateTime.Now.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

            string inputDaily = _inputDailyAppService.Find(filterInputDaily);
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift != "Daily").ToList(); ;          

            List<DailyModel> listDaily = inputDailyRes.AsEnumerable()
                      .Select(o => new DailyModel
                      {
                          LinkUp = o.LinkUp,
                          DateValue = o.Date,
                          Shift = o.Shift,
                          MTBFValue = o.MTBFValue,
                          CPQIValue = o.CPQIValue,
                          VQIValue = o.VQIValue,
                          WorkingValue = o.WorkingValue,
                          UptimeValue = o.UptimeValue,
                          STRSValue = o.STRSValue,
                          ProdVolumeValue = o.ProdVolumeValue,
                          CRRValue = o.CRRValue
                      }).ToList();

            var tupleModel = new Tuple<List<DailyModel>, List<TargetModel>, WTDModel, MTDModel>(listDaily, listTarget, modelWTD, modelMTD);

            return PartialView("_ReportTable", tupleModel);
        }

        #region Get data Lama
        /*
        [HttpPost]
        public ActionResult GetReportWithParam(DateTime dtFilter, string prodCenter, string machine)
        {
            try
            {              
                string refDetail = _referenceDetailAppService.GetAll(true);
                List<ReferenceDetailModel> refModelList = refDetail.DeserializeToRefDetailList();

                long pcID = GetRefID(prodCenter.ToUpper(), refModelList);            

                string month = dtFilter.ToString("MMM").ToUpper();

                string dayName = dtFilter.DayOfWeek.ToString();
                
                int diff = (7 + (dtFilter.DayOfWeek - DayOfWeek.Monday)) % 7;
                DateTime startOfWeek = dtFilter.AddDays(-1 * diff).Date;                
                DateTime firstDayOfMonth = new DateTime(dtFilter.Year, dtFilter.Month, 1);

                WTDModel modelWTD = GetWTD(startOfWeek, dtFilter, pcID, machine);
                MTDModel modelMTD = GetMTD(firstDayOfMonth, dtFilter, pcID, machine);

                ICollection<QueryFilter> filterInputTarget = new List<QueryFilter>();
                filterInputTarget = new List<QueryFilter>();
                filterInputTarget.Add(new QueryFilter("IsDeleted", "0"));
                filterInputTarget.Add(new QueryFilter("ProdCenterID", pcID));
                filterInputTarget.Add(new QueryFilter("Month", month));

                string inputTarget = _inputTargetAppService.Find(filterInputTarget);
                List<InputTargetModel> inputTargetRes = inputTarget.DeserializeToInputTargetList();                              

                List<TargetModel> listTarget = new List<TargetModel>();

                var mtbfCheck = inputTargetRes.Where(x => x.KPI == "MTBF").ToList();
                var cpqiCheck = inputTargetRes.Where(x => x.KPI == "CPQI").ToList();
                var vqiCheck = inputTargetRes.Where(x => x.KPI == "VQI").ToList();
                var workingCheck = inputTargetRes.Where(x => x.KPI == "Working Time").ToList();
                var uptimeCheck = inputTargetRes.Where(x => x.KPI == "Uptime").ToList();
                var strsCheck = inputTargetRes.Where(x => x.KPI == "STRS").ToList();
                var prodCheck = inputTargetRes.Where(x => x.KPI == "Production Volume").ToList();
                var crrCheck = inputTargetRes.Where(x => x.KPI == "CRR").ToList();

                if (mtbfCheck.Count() > 0)
                {
                    var rf11_M = mtbfCheck.Where(x=>x.Version=="RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = mtbfCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = mtbfCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = mtbfCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = mtbfCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = mtbfCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }                                  
                }
                if (cpqiCheck.Count() > 0)
                {
                    var rf11_M = cpqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = cpqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = cpqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = cpqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = cpqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = cpqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }

                }
                if (vqiCheck.Count() > 0)
                {
                    var rf11_M = vqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = vqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = vqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = vqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = vqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = vqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (workingCheck.Count() > 0)
                {
                    var rf11_M = workingCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = workingCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = workingCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = workingCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = workingCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = workingCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (uptimeCheck.Count() > 0)
                {
                    var rf11_M = uptimeCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = uptimeCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = uptimeCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = uptimeCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = uptimeCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = uptimeCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }

                }
                if (strsCheck.Count() > 0)
                {
                    var rf11_M = strsCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = strsCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = strsCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = strsCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = strsCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = strsCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (prodCheck.Count() > 0)
                {
                    var rf11_M = prodCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = prodCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = prodCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = prodCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = prodCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = prodCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (crrCheck.Count() > 0)
                {
                    var rf11_M = crrCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = crrCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = crrCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = crrCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = crrCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = crrCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }

             

                DateTime dateStart = dtFilter.AddDays(-2);                            

                ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
                filterInputDaily = new List<QueryFilter>();
                filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
                filterInputDaily.Add(new QueryFilter("ProdCenterID", pcID));
                filterInputDaily.Add(new QueryFilter("LinkUp", machine));
                filterInputDaily.Add(new QueryFilter("Date", dateStart.AddHours(-12).ToString(), Operator.GreaterThanOrEqual));
                filterInputDaily.Add(new QueryFilter("Date", dtFilter.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

                string inputDaily = _inputDailyAppService.Find(filterInputDaily);
                List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift != "Daily").ToList();

                List<DailyModel> listDaily = inputDailyRes.AsEnumerable()
                        .Select(o => new DailyModel
                        {
                            LinkUp = o.LinkUp,
                            DateValue = o.Date,
                            Shift = o.Shift,
                            MTBFValue = o.MTBFValue,
                            CPQIValue = o.CPQIValue,
                            VQIValue = o.VQIValue,
                            WorkingValue = o.WorkingValue,
                            UptimeValue = o.UptimeValue,
                            STRSValue = o.STRSValue,
                            ProdVolumeValue = o.ProdVolumeValue,
                            CRRValue = o.CRRValue                        
                        }).ToList();

                listDaily = listDaily.OrderBy(x => x.DateValue).ToList();

                var tupleModel = new Tuple<List<DailyModel>, List<TargetModel>, WTDModel, MTDModel>(listDaily, listTarget, modelWTD, modelMTD);

                return PartialView("_ReportTable", tupleModel);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Error = ex.Message });
            }
        }
        */
        #endregion
        [HttpPost]
        public ActionResult GeneratePDF(DateTime dtFilter, string prodCenter, string machine)
        {
            Document pdfDoc = new Document(PageSize.A4.Rotate(), 20, 20, 20, 20);
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfDoc.Open();


            pdfWriter.CloseStream = false;
            pdfDoc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=ShiftlyDaily.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Write(pdfDoc);
            Response.End();

            return View();
        }

        [HttpPost]
        public ActionResult GenerateExcelShift(DateTime dtFilter, string prodCenter, string machine) // don't forget the param
        {
            try
            {
                long pcID = long.Parse(prodCenter);
                string month = dtFilter.ToString("MMM").ToUpper();

                ICollection<QueryFilter> filterInputTarget = new List<QueryFilter>();
                filterInputTarget = new List<QueryFilter>();
                filterInputTarget.Add(new QueryFilter("IsDeleted", "0"));
                filterInputTarget.Add(new QueryFilter("ProdCenterID", pcID));
                filterInputTarget.Add(new QueryFilter("Month", month));

                string inputTarget = _inputTargetAppService.Find(filterInputTarget);
                List<InputTargetModel> inputTargetRes = inputTarget.DeserializeToInputTargetList();

                List<TargetModel> listTarget = new List<TargetModel>();

                var mtbfCheck = inputTargetRes.Where(x => x.KPI == "MTBF").ToList();
                var cpqiCheck = inputTargetRes.Where(x => x.KPI == "CPQI").ToList();
                var vqiCheck = inputTargetRes.Where(x => x.KPI == "VQI").ToList();
                var workingCheck = inputTargetRes.Where(x => x.KPI == "Working Time").ToList();
                var uptimeCheck = inputTargetRes.Where(x => x.KPI == "Uptime").ToList();
                var strsCheck = inputTargetRes.Where(x => x.KPI == "STRS").ToList();
                var prodCheck = inputTargetRes.Where(x => x.KPI == "Production Volume").ToList();
                var crrCheck = inputTargetRes.Where(x => x.KPI == "CRR").ToList();

                if (mtbfCheck.Count() > 0)
                {
                    var rf11_M = mtbfCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = mtbfCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = mtbfCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = mtbfCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = mtbfCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = mtbfCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (cpqiCheck.Count() > 0)
                {
                    var rf11_M = cpqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = cpqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = cpqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = cpqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = cpqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = cpqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }

                }
                if (vqiCheck.Count() > 0)
                {
                    var rf11_M = vqiCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = vqiCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = vqiCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = vqiCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = vqiCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = vqiCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (workingCheck.Count() > 0)
                {
                    var rf11_M = workingCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = workingCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = workingCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = workingCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = workingCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = workingCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (uptimeCheck.Count() > 0)
                {
                    var rf11_M = uptimeCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = uptimeCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = uptimeCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = uptimeCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = uptimeCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = uptimeCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }

                }
                if (strsCheck.Count() > 0)
                {
                    var rf11_M = strsCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = strsCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = strsCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = strsCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = strsCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = strsCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (prodCheck.Count() > 0)
                {
                    var rf11_M = prodCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = prodCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = prodCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = prodCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = prodCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = prodCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }
                if (crrCheck.Count() > 0)
                {
                    var rf11_M = crrCheck.Where(x => x.Version == "RF11").FirstOrDefault();
                    if (rf11_M == null)
                    {
                        var rf09_M = crrCheck.Where(x => x.Version == "RF09").FirstOrDefault();
                        if (rf09_M == null)
                        {
                            var rf06_M = crrCheck.Where(x => x.Version == "RF06").FirstOrDefault();
                            if (rf06_M == null)
                            {
                                var rf02_M = crrCheck.Where(x => x.Version == "RF02").FirstOrDefault();
                                if (rf02_M == null)
                                {
                                    var rfInternal_M = crrCheck.Where(x => x.Version == "Internal Target").FirstOrDefault();
                                    if (rfInternal_M == null)
                                    {
                                        var rfOB_M = crrCheck.Where(x => x.Version == "OB").FirstOrDefault();
                                        if (rfOB_M == null)
                                        {
                                        }
                                        else
                                        {
                                            TargetModel mTarget = new TargetModel
                                            {
                                                KPI = rfOB_M.KPI,
                                                Version = rfOB_M.Version,
                                                ValueTarget = rfOB_M.Value
                                            };
                                            listTarget.Add(mTarget);
                                        }
                                    }
                                    else
                                    {
                                        TargetModel mTarget = new TargetModel
                                        {
                                            KPI = rfInternal_M.KPI,
                                            Version = rfInternal_M.Version,
                                            ValueTarget = rfInternal_M.Value
                                        };
                                        listTarget.Add(mTarget);
                                    }
                                }
                                else
                                {
                                    TargetModel mTarget = new TargetModel
                                    {
                                        KPI = rf02_M.KPI,
                                        Version = rf02_M.Version,
                                        ValueTarget = rf02_M.Value
                                    };
                                    listTarget.Add(mTarget);
                                }
                            }
                            else
                            {
                                TargetModel mTarget = new TargetModel
                                {
                                    KPI = rf06_M.KPI,
                                    Version = rf06_M.Version,
                                    ValueTarget = rf06_M.Value
                                };
                                listTarget.Add(mTarget);
                            }
                        }
                        else
                        {
                            TargetModel mTarget = new TargetModel
                            {
                                KPI = rf09_M.KPI,
                                Version = rf09_M.Version,
                                ValueTarget = rf09_M.Value
                            };
                            listTarget.Add(mTarget);
                        }
                    }
                    else
                    {
                        TargetModel mTarget = new TargetModel
                        {
                            KPI = rf11_M.KPI,
                            Version = rf11_M.Version,
                            ValueTarget = rf11_M.Value
                        };
                        listTarget.Add(mTarget);
                    }
                }



                DateTime dateStart = dtFilter.AddDays(-2);

                ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
                filterInputDaily = new List<QueryFilter>();
                filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
                filterInputDaily.Add(new QueryFilter("ProdCenterID", pcID));
                filterInputDaily.Add(new QueryFilter("LinkUp", machine));
                filterInputDaily.Add(new QueryFilter("Date", dateStart.AddHours(-12).ToString(), Operator.GreaterThanOrEqual));
                filterInputDaily.Add(new QueryFilter("Date", dtFilter.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

                string inputDaily = _inputDailyAppService.Find(filterInputDaily);
                List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();

                List<DailyModel> listDaily = inputDailyRes.AsEnumerable()
                        .Select(o => new DailyModel
                        {
                            LinkUp = o.LinkUp,
                            DateValue = o.Date,
                            Shift = o.Shift,
                            MTBFValue = o.MTBFValue,
                            CPQIValue = o.CPQIValue,
                            VQIValue = o.VQIValue,
                            WorkingValue = o.WorkingValue,
                            UptimeValue = o.UptimeValue,
                            STRSValue = o.STRSValue,
                            ProdVolumeValue = o.ProdVolumeValue,
                            CRRValue = o.CRRValue
                        }).ToList();

                listDaily = listDaily.OrderBy(x => x.DateValue).ToList();

                ICollection<QueryFilter> filterUptime = new List<QueryFilter>();
                filterUptime = new List<QueryFilter>();
                filterUptime.Add(new QueryFilter("IsDeleted", "0"));
                filterUptime.Add(new QueryFilter("ProdCenterID", pcID));
                filterUptime.Add(new QueryFilter("LinkUp", machine));
                filterUptime.Add(new QueryFilter("Date", dtFilter.AddHours(-12).ToString(), Operator.GreaterThanOrEqual));
                filterUptime.Add(new QueryFilter("Date", dtFilter.AddHours(12).ToString(), Operator.LessThanOrEqualTo));

                string uptimeDaily = _inputDailyAppService.Find(filterUptime);
                List<InputDailyModel> uptimeDailyRes = uptimeDaily.DeserializeToInputDailyList();

                List<UptimeModel> listUptime = uptimeDailyRes.AsEnumerable()
                          .Select(o => new UptimeModel
                          {
                              Shift = o.Shift,
                              UptimeValue = o.UptimeValue,
                              UptimeFocus = o.UptimeFocus,
                              UptimeActPlan = o.UptimeActPlan
                          }).ToList();

                DateTime startOfWeek = dtFilter.StartOfWeek();
                DateTime firstDayOfMonth = new DateTime(dtFilter.Year, dtFilter.Month, 1);

                WTDModel modelWTD = GetWTD(startOfWeek, dtFilter, pcID, machine);
                MTDModel modelMTD = GetMTD(firstDayOfMonth, dtFilter, pcID, machine);

                var tupleModel = new Tuple<List<DailyModel>, List<TargetModel>, List<UptimeModel>, WTDModel, MTDModel>(listDaily, listTarget, listUptime, modelWTD, modelMTD);

                ReferenceDetailModel refDesc = GetRefDetail(pcID);
                string pcDesc = refDesc.Code + "-" + refDesc.Description;

                string dateVal = dtFilter.ToString("dd-MMM-yy");
                
                Session["DownloadExcel_ShiftDaily"] = ExcelGenerator.ExportShiftDaily(AccountName, tupleModel, pcDesc, machine, dateVal );               
                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.GenerateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Download()
        {

            if (Session["DownloadExcel_ShiftDaily"] != null)
            {
                byte[] data = Session["DownloadExcel_ShiftDaily"] as byte[];
                return File(data, "application/octet-stream", "ShiftDaily.xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }
        [HttpPost]
        public ActionResult SendEmail(DateTime dtFilter, string prodCenter, string machine)
        {
            try
            {


                return Json(new { Status = "True"}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Error = ex.Message });
            }          
        }
        [HttpPost]
        public ActionResult GetAllRemarksWithParam(string dateFilter, string shift)
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

            // Getting all data report remark   			
            string remarksList = _reportRemarksAppService.GetAll(true);
            List<ReportRemarksModel> remarks = remarksList.DeserializeToReportRemarksList().OrderByDescending(x => x.Date).ToList();

            int recordsTotal = remarks.Count();

            // Filter by param

            shift = shift == "" ? "" : shift;

            if (!string.IsNullOrEmpty(dateFilter))
            {
                DateTime dateFL = DateTime.Parse(dateFilter);
                if (shift == "1")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date == dateFL.Date).ToList();
                }
                if (shift == "2")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date == dateFL.Date).ToList();
                }
                if (shift == "3")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date == dateFL.Date).ToList();
                }
            }

                                                    
            if (string.IsNullOrEmpty(dateFilter))
            {
                DateTime dateFL = DateTime.Now;  
                if (shift == "1")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date ==dateFL.Date).ToList();
                }
                if (shift == "2")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date == dateFL.Date).ToList();
                }
                if (shift == "3")
                {
                    remarks = remarks.Where(m => m.Shift.Contains(shift) && m.Date == dateFL.Date).ToList();
                }
            }
              
            // Search part
            if (!string.IsNullOrEmpty(searchValue))
            {
                remarks = remarks.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
                                              (m.RSA != null ? m.RSA.ToLower().Contains(searchValue.ToLower()) : false) ||
                                               (m.Focus != null ? m.Focus.ToLower().Contains(searchValue.ToLower()) : false) ||
                                                (m.ActionPlan != null ? m.ActionPlan.ToLower().Contains(searchValue.ToLower()) : false) ||
                                              (m.Date != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false)).ToList();

            }

            // Sort part
            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
            {
                if (sortColumnDir == "asc")
                {
                    switch (sortColumn.ToLower())
                    {
                        case "id":
                            remarks = remarks.OrderBy(x => x.ID).ToList();
                            break;
                        case "shift":
                            remarks = remarks.OrderBy(x => x.Shift).ToList();
                            break;
                        case "rsa":
                            remarks = remarks.OrderBy(x => x.RSA).ToList();
                            break;
                        case "date":
                            remarks = remarks.OrderBy(x => x.Date).ToList();
                            break;
                        case "focus":
                            remarks = remarks.OrderBy(x => x.Focus).ToList();
                            break;
                        case "actionplan":
                            remarks = remarks.OrderBy(x => x.ActionPlan).ToList();
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
                            remarks = remarks.OrderBy(x => x.ID).ToList();
                            break;
                        case "shift":
                            remarks = remarks.OrderBy(x => x.Shift).ToList();
                            break;
                        case "rsa":
                            remarks = remarks.OrderBy(x => x.RSA).ToList();
                            break;
                        case "date":
                            remarks = remarks.OrderBy(x => x.Date).ToList();
                            break;
                        case "focus":
                            remarks = remarks.OrderBy(x => x.Focus).ToList();
                            break;
                        case "actionplan":
                            remarks = remarks.OrderBy(x => x.ActionPlan).ToList();
                            break;
                        default:
                            break;
                    }
                }
            }

            // total number of rows count     
            int recordsFiltered = remarks.Count();

            // Paging     
            var data = remarks.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();

            // Returning Json Data    
            return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
        }

        #region Helper       
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

        public static List<SelectListItem> GetProductionCenterInIndonesia(ILocationAppService _locationAppService, IReferenceAppService _referenceAppService, bool isIncludePleaseSelect)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            if (isIncludePleaseSelect)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Please Select -",
                    Value = ""
                });
            }

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = text.Code + " - " + text.Description,
                        Value = item.ID.ToString()
                    });
                }
            }

            return _menuList;
        }
        #endregion
        public string getCurrentWeekNumber(DateTime dateTime)
        {
            var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dateTime, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weeknum.ToString();/*dateTime.ToString($"{weeknum}")*/
        }
        private ReferenceDetailModel GetRefDetail(long idRef)
        {
            string refDetail = _referenceDetailAppService.GetById(idRef, true);

            return refDetail.DeserializeToRefDetail();
        }
        public WTDModel GetWTD(DateTime startOfWeek, DateTime currentDate, long pcID, string machine)
        {
            WTDModel model = new WTDModel();

            ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
            filterInputDaily = new List<QueryFilter>();
            filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
            filterInputDaily.Add(new QueryFilter("ProdCenterID", pcID));
            if (machine != "All")
            {
                filterInputDaily.Add(new QueryFilter("LinkUp", machine));
            }
            filterInputDaily.Add(new QueryFilter("Date", startOfWeek.ToString(), Operator.GreaterThanOrEqual));
            filterInputDaily.Add(new QueryFilter("Date", currentDate.ToString(), Operator.LessThanOrEqualTo));

            string inputDaily = _inputDailyAppService.Find(filterInputDaily);
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift != "Daily").ToList();

            List<DailyModel> listDaily = inputDailyRes.AsEnumerable()
                    .Select(o => new DailyModel
                    {
                        LinkUp = o.LinkUp,
                        DateValue = o.Date,
                        Shift = o.Shift,
                        MTBFValue = o.MTBFValue,
                        CPQIValue = o.CPQIValue,
                        VQIValue = o.VQIValue,
                        WorkingValue = o.WorkingValue,
                        UptimeValue = o.UptimeValue,
                        STRSValue = o.STRSValue,
                        ProdVolumeValue = o.ProdVolumeValue,
                        CRRValue = o.CRRValue
                    }).ToList();

            listDaily = listDaily.OrderBy(x => x.DateValue).ToList();

            decimal ProdVolumeValue = 0;
            decimal WorkingValue = 0;

            for (int i = 0; i < listDaily.Count(); i++)
            {
                //MTBFValue += listDaily[i].MTBFValue;
                ProdVolumeValue += listDaily[i].ProdVolumeValue;
                WorkingValue += listDaily[i].WorkingValue;
            }

            //model.MTBF_WTD = MTBFValue.ToString("n");
            model.ProdVolume_WTD = ProdVolumeValue.ToString("n");
            model.Working_WTD = WorkingValue.ToString("n");

            decimal MTBFValue = 0;
            decimal CPQIValue = 0;
            decimal VQIValue = 0;
            decimal STRSValue = 0;
            decimal CRRValue = 0;
            decimal UptimeValue = 0;

            List<DateTime> listDinoe = listDaily.Select(x => x.DateValue).Distinct().ToList();

            decimal[] VolumeDaily = new decimal[listDinoe.Count()];
            decimal[] MTBFDaily = new decimal[listDinoe.Count()];
            decimal[] CPQIDaily = new decimal[listDinoe.Count()];
            decimal[] VQIDaily = new decimal[listDinoe.Count()];
            decimal[] STRSDaily = new decimal[listDinoe.Count()];
            decimal[] CRRDaily = new decimal[listDinoe.Count()];

            int indexDate = 0;
            DailyModel nullShift = new DailyModel()
            {
                ProdVolumeValue = 0,
                MTBFValue = 0,
                CPQIValue = 0,
                VQIValue = 0,
                STRSValue = 0,
                CRRValue = 0,
                UptimeValue = 0
            };
            foreach (DateTime dt in listDinoe)
            {
                DailyModel shift1 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").FirstOrDefault() == null ? nullShift: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").FirstOrDefault();
                DailyModel shift2 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").FirstOrDefault() == null ? nullShift : listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").FirstOrDefault();
                DailyModel shift3 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").FirstOrDefault() == null ? nullShift : listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").FirstOrDefault();

                decimal[] listVolume = new decimal[] { shift1.ProdVolumeValue, shift2.ProdVolumeValue, shift3.ProdVolumeValue };
                decimal[] listMTBF = new decimal[] { shift1.MTBFValue, shift2.MTBFValue, shift3.MTBFValue };
                decimal[] listCPQI = new decimal[] { shift1.CPQIValue, shift2.CPQIValue, shift3.CPQIValue };
                decimal[] listVQI = new decimal[] { shift1.VQIValue, shift2.VQIValue, shift3.VQIValue };
                decimal[] listSTRS = new decimal[] { shift1.STRSValue, shift2.STRSValue, shift3.STRSValue };
                decimal[] listCRR = new decimal[] { shift1.CRRValue, shift2.CRRValue, shift3.CRRValue };
                decimal[] listUptime = new decimal[] { shift1.UptimeValue, shift2.UptimeValue, shift3.UptimeValue };



                if (machine == "All")
                {
                    listVolume = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Sum(x=>x.ProdVolumeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Sum(x=>x.ProdVolumeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Sum(x=>x.ProdVolumeValue)
                    };
                    listMTBF = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.MTBFValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.MTBFValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.MTBFValue)
                    };

                    listCPQI = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.CPQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.CPQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.CPQIValue)
                    };

                    listVQI = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.VQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.VQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.VQIValue) 
                    };

                    listSTRS = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.STRSValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.STRSValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.STRSValue) 
                    };

                    listCRR = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.CRRValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.CRRValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.CRRValue)
                    };
                    listUptime = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.UptimeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.UptimeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.UptimeValue)
                    };
                }

                VolumeDaily[indexDate] = listVolume.Sum();

                MTBFDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listMTBF, listVolume) / listVolume.Sum() : 0;
                CPQIDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listCPQI, listVolume) / listVolume.Sum() : 0;
                VQIDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listVQI, listVolume) / listVolume.Sum() : 0;
                STRSDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listSTRS, listVolume) / listVolume.Sum() : 0;
                CRRDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listCRR, listVolume) / listVolume.Sum() : 0;
                /*
                CPQIValue += (((listDaily[j].ProdVolumeValue * listDaily[j].CPQIValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].CPQIValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].CPQIValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue))*listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                VQIValue += (((listDaily[j].ProdVolumeValue * listDaily[j].VQIValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].VQIValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].VQIValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                STRSValue += (((listDaily[j].ProdVolumeValue * listDaily[j].STRSValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].STRSValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].STRSValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                CRRValue += (((listDaily[j].ProdVolumeValue * listDaily[j].CRRValue)
                                + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].CRRValue)
                                + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].CRRValue))
                                / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;
                */
                UptimeValue += ((listUptime[0] != 0 ? (listVolume[0] / listUptime[0] * 100) : 0)
                              + (listUptime[1] != 0 ? (listVolume[1] / listUptime[1] * 100) : 0)
                              + (listUptime[2] != 0 ? (listVolume[2] / listUptime[2] * 100) : 0)
                              );
                indexDate++;
            }
            if (ProdVolumeValue != 0)
            {
                MTBFValue = VolumeDaily.Sum() != 0 ? SumProduct(MTBFDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.MTBF_WTD = MTBFValue.ToString("n");

                CPQIValue = VolumeDaily.Sum() != 0 ? SumProduct(CPQIDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.CPQI_WTD = CPQIValue.ToString("n");

                VQIValue = VolumeDaily.Sum() != 0 ? SumProduct(VQIDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.VQI_WTD = VQIValue.ToString("n");

                STRSValue = VolumeDaily.Sum() != 0 ? SumProduct(STRSDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.STRS_WTD = STRSValue.ToString("n");

                CRRValue = VolumeDaily.Sum() != 0 ? SumProduct(CRRDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.CRR_WTD = CRRValue.ToString("n");

                UptimeValue = UptimeValue == 0 ? 0 : (ProdVolumeValue / UptimeValue) * 100;
                model.Uptime_WTD = UptimeValue.ToString("n");
            }
            if (ProdVolumeValue == 0)
            {
                model.MTBF_WTD = "0";
                model.CPQI_WTD = "0";
                model.VQI_WTD = "0";
                model.STRS_WTD = "0";
                model.CRR_WTD = "0";
                model.Uptime_WTD = "0";
            }          

            return model;
        }
        public MTDModel GetMTD(DateTime startOfMonth, DateTime currentDate, long pcID, string machine)
        {
            MTDModel model = new MTDModel();

            ICollection<QueryFilter> filterInputDaily = new List<QueryFilter>();
            filterInputDaily = new List<QueryFilter>();
            filterInputDaily.Add(new QueryFilter("IsDeleted", "0"));
            filterInputDaily.Add(new QueryFilter("ProdCenterID", pcID)); 
            if (machine != "All")
            {
                filterInputDaily.Add(new QueryFilter("LinkUp", machine));
            }
            filterInputDaily.Add(new QueryFilter("Date", startOfMonth.ToString(), Operator.GreaterThanOrEqual));
            filterInputDaily.Add(new QueryFilter("Date", currentDate.ToString(), Operator.LessThanOrEqualTo));

            string inputDaily = _inputDailyAppService.Find(filterInputDaily);
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift != "Daily").ToList();

            List<DailyModel> listDaily = inputDailyRes.AsEnumerable()
                    .Select(o => new DailyModel
                    {
                        LinkUp = o.LinkUp,
                        DateValue = o.Date,
                        Shift = o.Shift,
                        MTBFValue = o.MTBFValue,
                        CPQIValue = o.CPQIValue,
                        VQIValue = o.VQIValue,
                        WorkingValue = o.WorkingValue,
                        UptimeValue = o.UptimeValue,
                        STRSValue = o.STRSValue,
                        ProdVolumeValue = o.ProdVolumeValue,
                        CRRValue = o.CRRValue
                    }).ToList();

            listDaily = listDaily.OrderBy(x => x.DateValue).ToList();

            decimal ProdVolumeValue = 0;
            decimal WorkingValue = 0;

            for (int i = 0; i < listDaily.Count(); i++)
            {
                //MTBFValue += listDaily[i].MTBFValue;
                ProdVolumeValue += listDaily[i].ProdVolumeValue;
                WorkingValue += listDaily[i].WorkingValue;
            }

            //model.MTBF_MTD = MTBFValue.ToString("n");
            model.ProdVolume_MTD = ProdVolumeValue.ToString("n");
            model.Working_MTD = WorkingValue.ToString("n");

            decimal MTBFValue = 0;
            decimal CPQIValue = 0;
            decimal VQIValue = 0;
            decimal STRSValue = 0;
            decimal CRRValue = 0;
            decimal UptimeValue = 0;

            List<DateTime> listDinoe = listDaily.Select(x => x.DateValue).Distinct().ToList();
            decimal[] VolumeDaily = new decimal[listDinoe.Count()];
            decimal[] MTBFDaily = new decimal[listDinoe.Count()];
            decimal[] CPQIDaily = new decimal[listDinoe.Count()];
            decimal[] VQIDaily = new decimal[listDinoe.Count()];
            decimal[] STRSDaily = new decimal[listDinoe.Count()];
            decimal[] CRRDaily = new decimal[listDinoe.Count()];
            int indexDate = 0;
            DailyModel nullShift = new DailyModel()
            {
                ProdVolumeValue = 0,
                MTBFValue = 0,
                CPQIValue = 0,
                VQIValue = 0,
                STRSValue = 0,
                CRRValue = 0,
                UptimeValue = 0
            };
            foreach (DateTime dt in listDinoe)
            {
                DailyModel shift1 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").FirstOrDefault() == null ? nullShift : listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").FirstOrDefault();
                DailyModel shift2 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").FirstOrDefault() == null ? nullShift : listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").FirstOrDefault();
                DailyModel shift3 = listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").FirstOrDefault() == null ? nullShift : listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").FirstOrDefault();

                decimal[] listVolume = new decimal[] { shift1.ProdVolumeValue, shift2.ProdVolumeValue, shift3.ProdVolumeValue };
                decimal[] listMTBF = new decimal[] { shift1.MTBFValue, shift2.MTBFValue, shift3.MTBFValue };
                decimal[] listCPQI = new decimal[] { shift1.CPQIValue, shift2.CPQIValue, shift3.CPQIValue };
                decimal[] listVQI = new decimal[] { shift1.VQIValue, shift2.VQIValue, shift3.VQIValue };
                decimal[] listSTRS = new decimal[] { shift1.STRSValue, shift2.STRSValue, shift3.STRSValue };
                decimal[] listCRR = new decimal[] { shift1.CRRValue, shift2.CRRValue, shift3.CRRValue };
                decimal[] listUptime = new decimal[] { shift1.UptimeValue, shift2.UptimeValue, shift3.UptimeValue };
                
                if (machine == "All")
                {
                    listVolume = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Sum(x=>x.ProdVolumeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Sum(x=>x.ProdVolumeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Sum(x=>x.ProdVolumeValue)
                    };
                    listMTBF = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.MTBFValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.MTBFValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.MTBFValue)
                    };

                    listCPQI = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.CPQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.CPQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.CPQIValue)
                    };

                    listVQI = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.VQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.VQIValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.VQIValue)
                    };

                    listSTRS = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.STRSValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.STRSValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.STRSValue)
                    };

                    listCRR = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.CRRValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.CRRValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.CRRValue)
                    };

                    listUptime = new decimal[] {
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "1").ToList().Average(x=>x.UptimeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "2").ToList().Average(x=>x.UptimeValue),
                        listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").Count() == 0 ? 0: listDaily.Where(x => x.DateValue == dt && x.Shift.Trim() == "3").ToList().Average(x=>x.UptimeValue)
                    };
                }

                VolumeDaily[indexDate] = listVolume.Sum();

                MTBFDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listMTBF, listVolume) / listVolume.Sum() : 0;
                CPQIDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listCPQI, listVolume) / listVolume.Sum() : 0;
                VQIDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listVQI, listVolume) / listVolume.Sum() : 0;
                STRSDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listSTRS, listVolume) / listVolume.Sum() : 0;
                CRRDaily[indexDate] = listVolume.Sum() != 0 ? SumProduct(listCRR, listVolume) / listVolume.Sum() : 0;
                /*
                CPQIValue += (((listDaily[j].ProdVolumeValue * listDaily[j].CPQIValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].CPQIValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].CPQIValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue))*listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                VQIValue += (((listDaily[j].ProdVolumeValue * listDaily[j].VQIValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].VQIValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].VQIValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                STRSValue += (((listDaily[j].ProdVolumeValue * listDaily[j].STRSValue)
                                    + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].STRSValue)
                                    + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].STRSValue))
                                    / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;

                CRRValue += (((listDaily[j].ProdVolumeValue * listDaily[j].CRRValue)
                                + (listDaily[j+1].ProdVolumeValue * listDaily[j+1].CRRValue)
                                + (listDaily[j+2].ProdVolumeValue * listDaily[j+2].CRRValue))
                                / (listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue)) * listDaily[j].ProdVolumeValue + listDaily[j+1].ProdVolumeValue + listDaily[j+2].ProdVolumeValue;
                */

                UptimeValue += ((listUptime[0] != 0 ? (listVolume[0] / listUptime[0] * 100) : 0)
                              + (listUptime[1] != 0 ? (listVolume[1] / listUptime[1] * 100) : 0)
                              + (listUptime[2] != 0 ? (listVolume[2] / listUptime[2] * 100) : 0)
                              );
                indexDate++;
            }
            if (ProdVolumeValue != 0)
            {
                MTBFValue = VolumeDaily.Sum() != 0 ? SumProduct(MTBFDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.MTBF_MTD = MTBFValue.ToString("n");

                CPQIValue = VolumeDaily.Sum() != 0 ? SumProduct(CPQIDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.CPQI_MTD = CPQIValue.ToString("n");

                VQIValue = VolumeDaily.Sum() != 0 ? SumProduct(VQIDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.VQI_MTD = VQIValue.ToString("n");

                STRSValue = VolumeDaily.Sum() != 0 ? SumProduct(STRSDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.STRS_MTD = STRSValue.ToString("n");

                CRRValue = VolumeDaily.Sum() != 0 ? SumProduct(CRRDaily, VolumeDaily) / VolumeDaily.Sum() : 0;
                model.CRR_MTD = CRRValue.ToString("n");

                UptimeValue = UptimeValue == 0 ? 0 : (ProdVolumeValue / UptimeValue) * 100;
                model.Uptime_MTD = UptimeValue.ToString("n");
            }

            if (ProdVolumeValue == 0)
            {
                model.MTBF_MTD = "0";
                model.CPQI_MTD = "0";
                model.VQI_MTD = "0";
                model.STRS_MTD = "0";
                model.CRR_MTD = "0";
                model.Uptime_MTD = "0";
            }
            return model;
        }

    }
}
