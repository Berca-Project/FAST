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
using System.Threading.Tasks;

namespace Fast.Web.Controllers.Report
{
    public class ShiftDailyUrlExcelController : BaseController<LPHModel>
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
        public ShiftDailyUrlExcelController(
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
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift.Trim() != "Daily").ToList();

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
        public ActionResult GenerateShiftlyDailyWithParam(string key, DateTime date, string prodCenter, string machine)
        {
            try
            {
                if (key.Equals("S2F0YWthbmxhaCAoTXVoYW1tYWQpLCAnRGlhbGFoIEFsbGFoLCBZYW5nIE1haGEgRXNhJy4gQWxsYWggdGVtcGF0IG1lbWludGEgc2VnYWxhIHNlc3VhdHUuIChBbGxhaCkgdGlkYWsgYmVyYW5hayBkYW4gdGlkYWsgcHVsYSBkaXBlcmFuYWtrYW4uIERhbiB0aWRhayBhZGEgc2VzdWF0dSB5YW5nIHNldGFyYSBkZW5nYW4gRGlhLg=="))
                {
                    System.Drawing.Color colFromHex = ColorTranslator.FromHtml("#99ccff");
                    string[] higherBetter = new string[] { "MTBF", "Uptime", "STRS", "Volume" };
                    string[] lowerBetter = new string[] { "CRR", "VQI", "CPQI" };
                    var (linkupList, remarkList) = GetDataWithParam(date, prodCenter, machine);
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
                        Sheet.Cells[rowPos, 3].Value = date.AddDays(-2).ToString("dd-MMM-yyyy");
                        Sheet.Cells[rowPos, 4].Value = date.AddDays(-1).ToString("dd-MMM-yyyy");
                        Sheet.Cells[rowPos, 5].Value = date.ToString("dd-MMM-yyyy");
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
                            Sheet.Cells[rowPos, 1].Value = type+ satuan;
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

                    using (var range = remarkSheet.Cells[remarkRowPos, 1, remarkRowPos, 9])
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
                        remarkSheet.Cells[remarkRowPos, 1].Value = remark["shift"];
                        remarkSheet.Cells[remarkRowPos, 2].Value = remark["type"];
                        remarkSheet.Cells[remarkRowPos, 3].Value = remark["remark"];
                        remarkSheet.Cells[remarkRowPos++, 4].Value = remark["actionplan"];
                    }
                    remarkSheet.Cells[1, 1, remarkRowPos - 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    remarkSheet.Cells[1, 1, remarkRowPos - 1, 4].AutoFitColumns();
                    byte[] dataExcel = Ep.GetAsByteArray();
                    return File(dataExcel, "application/octet-stream", "ReportShiftlyDaily.xlsx");
                }
                else
                {
                    return Json(new { Status = "False", Error = "Unauthorized" }, JsonRequestBehavior.AllowGet);
                }
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
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift.Trim() != "Daily").ToList();

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
            List<InputDailyModel> inputDailyRes = inputDaily.DeserializeToInputDailyList();//.Where(x => x.Shift.Trim() != "Daily").ToList();

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
        public async Task<ActionResult> SentDailyEmailWithParam(string key, DateTime date, string prodcenter, string lu)
        {
            try
            {
                if (key.Equals("S2F0YWthbmxhaCwg4oCcQWt1IGJlcmxpbmR1bmcga2VwYWRhIFR1aGFubnlhIG1hbnVzaWEsClJhamEgbWFudXNpYSwKc2VtYmFoYW4gbWFudXNpYSwKZGFyaSBrZWphaGF0YW4gKGJpc2lrYW4pIHNldGFuIHlhbmcgYmVyc2VtYnVueWksCnlhbmcgbWVtYmlzaWtrYW4gKGtlamFoYXRhbikga2UgZGFsYW0gZGFkYSBtYW51c2lhLApkYXJpIChnb2xvbmdhbikgamluIGRhbiBtYW51c2lhLuKAnQ=="))
                {
                    //string emailTo = "AchmadYusuf.GBU@contracted.sampoerna.com";

                    string getIDReference = _referenceAppService.FindBy("Name", "SP Shiftly Daily", true);
                    List<ReferenceModel> referenceList = getIDReference.DeserializeToReferenceList();
                    if (referenceList.Count > 0)
                    {
                        string resultCloveCon = _referenceDetailAppService.FindBy("ReferenceID", referenceList[0].ID, true);
                        List<ReferenceDetailModel> emailList = resultCloveCon.DeserializeToRefDetailList();
                        foreach (ReferenceDetailModel rdm in emailList)
                        {
                            await EmailSender.SendEmailReportShiftlyDaily(date, prodcenter, lu, rdm.Description);
                        }
                    }
                    return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { Status = "False", Error = "Unauthorized" }, JsonRequestBehavior.AllowGet);

                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}
