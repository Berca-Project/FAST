using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("ReportPPDetailBatch")]
    public class ReportPPDetailBatchController : BaseController<PPLPHModel>
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

        public ReportPPDetailBatchController(
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
        public ActionResult GetData(string startDate, string endDate, long prodCenterID, List<string> lphHeader, string batch)
        {
            try
            {
                ExecuteQuery("UPDATE [PPLPHExtras] SET [IsDeleted] = 1 WHERE [ValueType] = 'json' AND [Value] NOT LIKE '%]';");
                DateTime dtStart = DateTime.ParseExact(startDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEnd = DateTime.ParseExact(endDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                if (dtStart > dtEnd)
                {
                    SetFalseTempData("Start Date must be less than End Date");
                    return RedirectToAction("Index");
                }
                dtStart = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 0, 0, 0);
                dtEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 23, 59, 59);

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenterID, "productioncenter");
                List<QueryFilter> submissionsFilter = new List<QueryFilter>();
                locationIdList.ForEach(loc => submissionsFilter.Add(new QueryFilter("LocationID", loc.ToString(), Operator.Equals, Operation.OrElse)));
                if (locationIdList.Count() % 2 == 1) submissionsFilter.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                submissionsFilter.Add(new QueryFilter("Date", dtStart.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                submissionsFilter.Add(new QueryFilter("Date", dtEnd.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
                List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.Find(submissionsFilter).DeserializeToPPLPHSubmissionsList()
                    .Where(x => lphHeader.Any(y => y == x.LPHHeader)).ToList();
                submissions = submissions.Where(s => s == null ? false : _ppLphApprovalAppService.FindBy("LPHSubmissionID", s.ID, false).DeserializeToPPLPHApprovalList().Any(x =>
                {
                    string status = x.Status.Trim().ToLower();
                    return status == "approved" || status == "submitted";
                })).ToList();

                List<QueryFilter> filterByComponent = new List<QueryFilter>();
                submissions.ForEach(sub => filterByComponent.Add(new QueryFilter("LPHID", sub.LPHID.ToString(), Operator.Equals, Operation.OrElse)));
                if (submissions.Count() % 2 == 1) filterByComponent.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                filterByComponent.Add(new QueryFilter("IsDeleted", "0"));
                List<PPLPHComponentsModel> componentList = submissions.Count == 0 ? new List<PPLPHComponentsModel>() : _ppLphComponentsAppService.Find(filterByComponent).DeserializeToPPLPHComponentList();
                List<QueryFilter> filterForValues = new List<QueryFilter>();
                componentList.ForEach(compo => filterForValues.Add(new QueryFilter("LPHComponentID", compo.ID.ToString(), Operator.Equals, Operation.OrElse)));
                if (componentList.Count() % 2 == 1) filterForValues.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                filterForValues.Add(new QueryFilter("IsDeleted", "0"));
                List<PPLPHValuesModel> valueList = submissions.Count == 0 ? new List<PPLPHValuesModel>() : _ppLphValuesAppService.Find(filterForValues).DeserializeToPPLPHValueList();
                List<QueryFilter> filterForExtras = new List<QueryFilter>();
                submissions.ForEach(sub => filterForExtras.Add(new QueryFilter("LPHID", sub.LPHID.ToString(), Operator.Equals, Operation.OrElse)));
                if (submissions.Count() % 2 == 1) filterForExtras.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                filterForExtras.Add(new QueryFilter("IsDeleted", "0"));
                List<PPLPHExtrasModel> extraList = submissions.Count == 0 ? new List<PPLPHExtrasModel>() : _ppLphExtrasAppService.Find(filterForExtras).DeserializeToPPLPHExtrasList();

                Dictionary<string, (Dictionary<string, string>, Dictionary<string, string>)> renaming = new Dictionary<string, (Dictionary<string, string>, Dictionary<string, string>)>()
                {
                    {"LPHPrimaryCloveInfeedConditioningController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNoInfeed", "Batch No." },
                                { "BlendInfeed", "Blend" },
                                { "SAPBatchInfeed", "SAP Batch" },
                                { "FormulaInfeed", "Formula Rev" },
                                { "StartInfeed", "Start" },
                                { "StopInfeed", "Stop" },
                                { "TpcInfeed", "No TPC" },
                                { "SiloDestinationInfeed", "SILO Destination" },
                                { "QuantityWeighterInfeed", "Quantity Weighter" },
                                { "DurationInfeed", "Duration" },
                                { "BatchNoDCC", "Batch No." },
                                { "BlendDCC", "Blend" },
                                { "SAPBatchDCC", "SAP Batch" },
                                { "FormulaDCC", "Formula Rev" },
                                { "StartDCC", "Start" },
                                { "StopDCC", "Stop" },
                                { "SiloDestinationDCC", "Silo Destination" },
                                { "FlowRateInputDCC", "Flow Rate Input" },
                                { "QtyWeigcon1DCC", "Quantity WeighCon 1" },
                                { "QtyWeigcon2DCC", "Quantity WeighCon 2" },
                                { "CCS1DCC", "Casing Clove Screw 1" },
                                { "CCS2DCC", "Casing Clove Screw 2" },
                                { "TWA1DCC", "Totalizer Water App 1" },
                                { "TWA2DCC", "Totalizer Water App 2" },
                                { "SteamPreasure1DCC", "Steam Preasure Screw 1(Bar)" },
                                { "SteamPreasure2DCC", "Steam Preasure Screw 2(Bar)" },
                                { "CloveTemp1DCC", "Clove Temp Exit Cond 1 Screw (&deg;C)" },
                                { "CloveTemp2DCC", "Clove Temp Exit Cond 2 Screw (&deg;C)" },
                                { "ResidenceTimeDCC", "Residence Time (detik)" },
                                { "DurationDCC", "Duration" },
                                { "numberLot", "Number" },
                                { "qtyLot", "Qty" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryCloveCutDryPackingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SAPBatch", "SAP Batch" },
                                { "AsalSilo", "Asal Silo" },
                                { "TujuanSilo", "Tujuan Silo" },
                                { "CutterNo", "Cutter No" },
                                { "FlowRateInput", "Clove Flow Rate Input Cutter" },
                                { "WeighconBefore", "Total Weighcon Before Cutter" },
                                { "CONoTangki", "No Tangki (Clove Oil)" },
                                { "COKg", "Kg (Clove Oil)" },
                                { "COCleve", "Cleve (Clove Oil)" },
                                { "WOC1", "WOC Cutter 1" },
                                { "WOC2", "WOC Cutter 2" },
                                { "WOC3", "WOC Cutter 3" },
                                { "WOC4", "WOC Cutter 4" },
                                { "WOC5", "WOC Cutter 5" },
                                { "WOC6", "WOC Cutter 6" },
                                { "WOC7", "WOC Cutter 7" },
                                { "WOC8", "WOC Cutter 8" },
                                { "WOC9", "WOC Cutter 9" },
                                { "WOC10", "WOC Cutter 10" },
                                { "CondensateOPS", "Ops Batch" },
                                { "CondensateNoTangki", "No Tangki" },
                                { "CondensateKg", "Kg" },
                                { "FlowRateIn0", "Flow Rate Input Dryer (4650/0027)" },
                                { "FlowRateIn1", "Flow Rate Input Dryer (4651/0006)" },
                                { "AirTemp0", "Air Temp Dryer (4650/0027)" },
                                { "AirTemp1", "Air Temp Dryer (4651/0006)" },
                                { "HoodPressure0", "Hood Pressure (4650/0027)" },
                                { "HoodPressure1", "Hood Pressure (4651/0006)" },
                                { "ResidenceTime0", "Residence Time In Dryer (4650/0027)" },
                                { "ResidenceTime1", "Residence Time In Dryer (4651/0006)" },
                                { "TotalWeyconFeed0", "Total Weycon Infeed Dryer (4650/0027)" },
                                { "TotalWeyconFeed1", "Total Weycon Infeed Dryer (4651/0006)" },
                                { "TempExit0", "Temp Exit Dryer (4650/0027)" },
                                { "TempExit1", "Temp Exit Dryer (4651/0006)" },
                                { "MoistureInput0", "Moisture Input Dryer (4650/0027)" },
                                { "MoistureInput1", "Moisture Input Dryer (4651/0006)" },
                                { "MoistureExit0", "Moisture Exit Dryer (4650/0027)" },
                                { "MoistureExit1", "Moisture Exit Dryer (4651/0006)" },
                                { "TotalWeyconMaster0", "Total Weighcon Master (4650/0027)" },
                                { "AddbackAngkup0", "Addback Angkup (4650/0027)" },
                                { "MCFinal0", "MC Final (4650/0027)" },
                                { "TopBandPreasure", "Top band pressure (Bar)" },
                                { "MouthHeight", "Mouth Height (mm)" },
                                { "WidthOfCut", "Width of cut (mm)" },
                                { "DetectedMetal", "Detected metal (times)" },
                                { "CutterStop", "Cutter stop (times)" },
                                { "SharpKnives", "Sharp knives (OK/Not OK)" },
                                { "CounterTraverse1", "Counter traverse gerinda (Impulse)" },
                                { "CounterTraverse2", "Counter traverse pisau (Impulse)" },
                                { "AddbackMaterial", "Addback Material reverse cutter (kg)" },
                                { "packingclovenoopspacking", "No Ops Packing" },
                                { "packingclovestart", "Start" },
                                { "packingclovestop", "Stop" },
                                { "packingclovetotal", "Total" },
                                { "packingclovetotalbox", "Total Box" },
                                { "packingcloveqty", "QTY (Kg)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryCSFInfeedConditioningController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "infeedBatchNo", "Batch No." },
                                { "infeedSapBatch", "SAP Batch" },
                                { "infeedBlend", "Blend" },
                                { "infeedFormulaRev", "Formula Rev." },
                                { "infeedSiloDestination", "SILO Destination" },
                                { "infeedStart", "Start" },
                                { "infeedStop", "Stop" },
                                { "infeedTpc", "No TPC" },
                                { "infeedQuantityWeighter", "Quantity Weighter" },
                                { "infeedDuration", "Duration" },
                                { "conditioningBatchNo", "Batch No." },
                                { "conditioningStart", "Start" },
                                { "conditioningStop", "Stop" },
                                { "conditioningDuration", "Duration" },
                                { "conditioningSapBatch", "SAP Batch" },
                                { "conditioningBlend", "Blend" },
                                { "conditioningFormulaRev", "Formula Rev." },
                                { "conditioningFri1", "Flow Rate Input 1" },
                                { "conditioningFri2", "Flow Rate Input 2" },
                                { "conditioningSiloDestination", "SILO Destination" },
                                { "conditioningTwa1", "Totalizer Water App 1" },
                                { "conditioningTwa2", "Totalizer Water App 2" },
                                { "conditioningQw1", "Quantity WeightCon 1" },
                                { "conditioningQw2", "Quantity WeightCon 2" },
                                { "Steam", "Steam Pressure Screw 1 (Bar)" },
                                { "CloveTemp", "Clove Temp Exit Cond 1 Screw (&deg;C)" },
                                { "ResidenceTime", "Residence Time (detik)" },
                                { "lotnumber", "Number" },
                                { "lotqty", "Qty" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryCSFCutDryPackingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "cuttingBatchNo", "Batch No." },
                                { "cuttingSapBatch", "SAP Batch" },
                                { "cuttingBlend", "Blend" },
                                { "cuttingFormulaRev", "Formula Rev." },
                                { "cuttingStart", "Start" },
                                { "cuttingStop", "Stop" },
                                { "cuttingDuration", "Duration" },
                                { "cuttingSilodestination", "Silo Destination" },
                                { "cuttingTotalizermaterial", "Totalizer Material" },
                                { "cuttingTotalizerflavour", "Totalizer Flavour" },
                                { "cuttingMcoutflavour", "MC Out Flavour" },
                                { "cuttingTwv1", "Totalizer Weigher Vibro (Line 1)" },
                                { "cuttingMom1", "MC Out Microwave (Line 1)" },
                                { "cuttingRemark1", "Remark (Line 1)" },
                                { "cuttingTwv2", "Totalizer Weigher Vibro (Line 2)" },
                                { "cuttingMom2", "MC Out Microwave (Line 2)" },
                                { "cuttingRemark2", "Remark (Line 2)" },
                                { "cuttingTwv3", "Totalizer Weigher Vibro (Line 3)" },
                                { "cuttingMom3", "MC Out Microwave (Line 3)" },
                                { "cuttingRemark3", "Remark (Line 3)" },
                                { "cuttingTwv4", "Totalizer Weigher Vibro (Line 4)" },
                                { "cuttingMom4", "MC Out Microwave (Line 4)" },
                                { "cuttingRemark4", "Remark (Line 4)" },
                                { "cuttingTwv1n2", "Totalizer Weigher Vibro (Weicon Line 1&2)" },
                                { "cuttingMom1n2", "MC Out Microwave (Weicon Line 1&2)" },
                                { "cuttingRemark1n2", "Remark (Weicon Line 1&2)" },
                                { "cuttingTwv3n4", "Totalizer Weigher Vibro (Weicon Line 3&4)" },
                                { "cuttingMom3n4", "MC Out Microwave (Weicon Line 3&4)" },
                                { "cuttingRemark3n4", "Remark (Weicon Line 3&4)" },
                                { "cuttingTwvtfw", "Totalizer Weigher Vibro (Total Final Weigher)" },
                                { "cuttingMomtfw", "MC Out Microwave (Total Final Weigher)" },
                                { "cuttingRemarktfw", "Remark (Total Final Weigher)" },
                                { "tl1_awal1", "Thickness line 1 (mm) - Awal 1" },
                                { "tl1_awal2", "Thickness line 1 (mm) - Awal 2" },
                                { "tl1_awal3", "Thickness line 1 (mm) - Awal 3" },
                                { "tl1_tengah1", "Thickness line 1 (mm) - Tengah 1" },
                                { "tl1_tengah2", "Thickness line 1 (mm) - Tengah 2" },
                                { "tl1_tengah3", "Thickness line 1 (mm) - Tengah 3" },
                                { "tl1_akhir1", "Thickness line 1 (mm) - Akhir 1" },
                                { "tl1_akhir2", "Thickness line 1 (mm) - Akhir 2" },
                                { "tl1_akhir3", "Thickness line 1 (mm) - Akhir 3" },
                                { "tl2_awal1", "Thickness line 2 (mm) - Awal 1" },
                                { "tl2_awal2", "Thickness line 2 (mm) - Awal 2" },
                                { "tl2_awal3", "Thickness line 2 (mm) - Awal 3" },
                                { "tl2_tengah1", "Thickness line 2 (mm) - Tengah 1" },
                                { "tl2_tengah2", "Thickness line 2 (mm) - Tengah 2" },
                                { "tl2_tengah3", "Thickness line 2 (mm) - Tengah 3" },
                                { "tl2_akhir1", "Thickness line 2 (mm) - Akhir 1" },
                                { "tl2_akhir2", "Thickness line 2 (mm) - Akhir 2" },
                                { "tl2_akhir3", "Thickness line 2 (mm) - Akhir 3" },
                                { "tl3_awal1", "Thickness line 3 (mm) - Awal 1" },
                                { "tl3_awal2", "Thickness line 3 (mm) - Awal 2" },
                                { "tl3_awal3", "Thickness line 3 (mm) - Awal 3" },
                                { "tl3_tengah1", "Thickness line 3 (mm) - Tengah 1" },
                                { "tl3_tengah2", "Thickness line 3 (mm) - Tengah 2" },
                                { "tl3_tengah3", "Thickness line 3 (mm) - Tengah 3" },
                                { "tl3_akhir1", "Thickness line 3 (mm) - Akhir 1" },
                                { "tl3_akhir2", "Thickness line 3 (mm) - Akhir 2" },
                                { "tl3_akhir3", "Thickness line 3 (mm) - Akhir 3" },
                                { "tl4_awal1", "Thickness line 4 (mm) - Awal 1" },
                                { "tl4_awal2", "Thickness line 4 (mm) - Awal 2" },
                                { "tl4_awal3", "Thickness line 4 (mm) - Awal 3" },
                                { "tl4_tengah1", "Thickness line 4 (mm) - Tengah 1" },
                                { "tl4_tengah2", "Thickness line 4 (mm) - Tengah 2" },
                                { "tl4_tengah3", "Thickness line 4 (mm) - Tengah 3" },
                                { "tl4_akhir1", "Thickness line 4 (mm) - Akhir 1" },
                                { "tl4_akhir2", "Thickness line 4 (mm) - Akhir 2" },
                                { "tl4_akhir3", "Thickness line 4 (mm) - Akhir 3" },
                                { "mwrl1_1", "M.wave Running Line 1 (%) - Result 1" },
                                { "mwrl1_2", "M.wave Running Line 1 (%) - Result 2" },
                                { "mwrl1_3", "M.wave Running Line 1 (%) - Result 3" },
                                { "mwrl2_1", "M.wave Running Line 2 (%) - Result 1" },
                                { "mwrl2_2", "M.wave Running Line 2 (%) - Result 2" },
                                { "mwrl2_3", "M.wave Running Line 2 (%) - Result 3" },
                                { "mwrl3_1", "M.wave Running Line 3 (%) - Result 1" },
                                { "mwrl3_2", "M.wave Running Line 3 (%) - Result 2" },
                                { "mwrl3_3", "M.wave Running Line 3 (%) - Result 3" },
                                { "mwrl4_1", "M.wave Running Line 4 (%) - Result 1" },
                                { "mwrl4_2", "M.wave Running Line 4 (%) - Result 2" },
                                { "mwrl4_3", "M.wave Running Line 4 (%) - Result 3" },
                                { "packinglinenoopspacking", "No Ops Packing" },
                                { "packinglinestart", "Start" },
                                { "packinglinestop ", "Stop" },
                                { "packinglinetotal", "Total" },
                                { "packinglinetotalbox", "Total Box" },
                                { "packinglineqty", "QTY (Kg)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryRTCController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "gonesack1", "Pemakaian Guargum - Sack" },
                                { "gonepg1", "Pemakaian Guargum - Kg" },
                                { "gonesack2", "Flovour Gorden (Tepung Tapioka) - Sack" },
                                { "gonefg", "Flovour Gorden (Tepung Tapioka) - Kg" },
                                { "ACRoystonAwal", "AC Royston Awal (Kg)" },
                                { "ACRoystonAkhir", "AC Royston Akhir (Kg)" },
                                { "ACRoystonTotal", "AC Royston Total (Kg)" },
                                { "AlanisAwal", "Alanis Awal (Kg)" },
                                { "AlanisAkhir", "Alanis Akhir (Kg)" },
                                { "AlanisTotal", "Alanis Total (Kg)" },
                                { "Bilr1", "Batch ID Liquid RTC 1" },
                                { "Bilr2", "Batch ID Liquid RTC 2" },
                                { "Bilr3", "Batch ID Liquid RTC 3" },
                                { "Bilr4", "Batch ID Liquid RTC 4" },
                                { "Bilr5", "Batch ID Liquid RTC 5" },
                                { "QtyLiquid1", "Qty Liquid 1" },
                                { "QtyLiquid2", "Qty Liquid 2" },
                                { "QtyLiquid3", "Qty Liquid 3" },
                                { "QtyLiquid4", "Qty Liquid 4" },
                                { "QtyLiquid5", "Qty Liquid 5" },
                                { "Bir1", "Batch ID Royston 1" },
                                { "Bir2", "Batch ID Royston 2" },
                                { "Bir3", "Batch ID Royston 3" },
                                { "QtyRoyston1", "Qty Royston 1" },
                                { "QtyRoyston2", "Qty Royston 2" },
                                { "QtyRoyston3", "Qty Royston 3" },
                                { "PackingBatchNo", "Batch No" },
                                { "PackingStart", "No Start" },
                                { "PackingStop", "No Stop" },
                                { "PackingTotalBox", "Total (Box)" },
                                { "PackingTotalKg", "Total (Kg)" },
                                { "FeedingBatchNo", "No Batch" },
                                { "RevisiBatchNo", "No Revisi" },
                                { "Line1234OB1", "Oval Bin 1 - Line 3&4 / 1&2 (kg)" },
                                { "Line1234OB2", "Oval Bin 2 - Line 3&4 / 1&2 (kg)" },
                                { "Line1234OB3", "Oval Bin 3 - Line 3&4 / 1&2 (kg)" },
                                { "Line1234OB4", "Oval Bin 4 - Line 3&4 / 1&2 (kg)" },
                                { "Line1234Total", "Total - Line 3&4 / 1&2 (kg)" },
                                { "Line56OB1", "Oval Bin 1 - Line 5&6 (kg)" },
                                { "Line56OB2", "Oval Bin 2 - Line 5&6 (kg)" },
                                { "Line56OB3", "Oval Bin 3 - Line 5&6 (kg)" },
                                { "Line56OB4", "Oval Bin 4 - Line 5&6 (kg)" },
                                { "Line56Total", "Total - Line 5&6 (kg)" },
                                { "proporline", "Line" },
                                { "proporstart", "Start" },
                                { "proporstop", "Stop" },
                                { "proportotal", "Total Menit" },
                                { "proporpms", "Powder Material Setting" },
                                { "proporlms", "Liquid Material Setting" },
                                { "propormixingtime", "Mixing Time (detik)" },
                                { "proporidletime", "Idle Time (detik)" },
                                { "propor41sd47", "41 sd 47" },
                                { "moldingline", "Line" },
                                { "moldingnobatch", "No Batch" },
                                { "moldingstart", "Start" },
                                { "moldingstop", "Stop" },
                                { "moldingduration", "Duration" },
                                { "moldingroll1", "Material Thickness (mm) - Roll 1" },
                                { "moldingroll2", "Material Thickness (mm) - Roll 2" },
                                { "moldingroll3", "Material Thickness (mm) - Roll 3" },
                                { "moldingoutputdrying", "Material Thickness (mm) - Output Drying" },
                                { "moldingclo", "Cut Length Output (mm)" },
                                { "moldingcwo", "Cut Width Output (mm)" },
                                { "moldingmac", "Mc After Cutr (%)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryDietController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "ImpregnationNoOps", "No.OPS" },
                                { "ImpregnationBatchTo", "Batch To" },
                                { "ImpregnationSiloOrigin", "Silo Origin" },
                                { "ImpregnationWeightCon", "Weighcon Qty Weighter Real" },
                                { "ImpregnationMCInput", "MC Input (%)" },
                                { "ImpregnationStartBatch", "Start Batch" },
                                { "ImpregnationStopBatch", "Stop Batch" },
                                { "ImpregnationCycleTime", "Cycle Time" },
                                { "ImpregnationTempV23", "Temp.V23 (&deg;C)" },
                                { "ImpregnationCO2Start", "CO2 Start" },
                                { "ImpregnationCO2Stop", "CO2 Stop" },
                                { "ImpregnationCO2Usage", "CO2 Usage" },
                                { "ImpregnationPressure", "Pressure (bar)" },
                                { "ReorderingBlendCode", "Blend Code" },
                                { "ReorderingNoOps", "No OPS" },
                                { "ReorderingSiloOrigin", "No Silo Asal" },
                                { "ReorderingStartBatch", "Start Batch" },
                                { "ReorderingStopBatch", "Stop Batch" },
                                { "ReorderingDuration", "Duration" },
                                { "ReorderingFlowrate", "Flowrate (Kg/h)" },
                                { "ReorderingPressurePG", "Pressure PG" },
                                { "ReorderingTempLoop", "Temp. Loop." },
                                { "ReorderingTempOxidizer", "Temp. Oxidizer" },
                                { "ReorderingMCAfterRC80", "MC After RC80" },
                                { "ReorderingCCV", "CCV" },
                                { "ReorderingGasCNGStart", "Gas CNG Start" },
                                { "ReorderingGasCNGStop", "Gas CNG Stop" },
                                { "ReorderingGasCNGUsage", "Gas CNG Usage" },
                                { "ReorderingSteamStart", "Steam Start" },
                                { "ReorderingSteamStop", "Steam Stop" },
                                { "ReorderingSteamUsage", "Steam Usage" },
                                { "PackingBlendCode", "Blend Code" },
                                { "PackingSiloOrigin", "No Silo Asal" },
                                { "PackingStartPacking", "Start" },
                                { "PackingStopPacking", "Stop" },
                                { "PackingDuration", "Duration" },
                                { "PackingJenisBox", "Jenis Box" },
                                { "PackingVisual", "Visual" },
                                { "PackingQtyBox", "Qty (Box)" },
                                { "PackingQtyKg", "Qty (Kg)" },
                                { "PackingSamplingNo", "No - Sampling Timbang Ulang" },
                                { "PackingSamplingNoBox", "No Box - Sampling Timbang Ulang" },
                                { "PackingSamplingQty", "Quantity Label - Sampling Timbang Ulang" },
                                { "PackingSamplingTimbang", "Timbang Ulang - Sampling Timbang Ulang" },
                                { "PackingSamplingError", "Error - Sampling Timbang Ulang" },
                                { "WasteBlend", "Blend" },
                                { "WasteNoSKJ", "No. SKJ - OPS" },
                                { "WasteNoPJ", "No. PJ - OPS" },
                                { "WasteRM", "RM" },
                                { "WasteAddback", "Addback" },
                                { "WastePackingBox", "Packing Box" },
                                { "WastePackingKg", "Packing Kg" },
                                { "WasteYield", "Yield" },
                                { "WasteIndexFinalOV", "Index Final OV" },
                                { "WasteCVIB0069", "CVIB 0069 - Waste" },
                                { "WasteCSFR0022", "CSFR 0022 - Waste" },
                                { "WasteDSCL0034", "DSCL 0034 - Waste" },
                                { "WasteCVIB0070", "CVIB 0070 - Waste" },
                                { "WasteRV0054", "RV 0054 - Waste" },
                                { "WasteRC80Dry", "RC80 Dry - Waste" },
                                { "WasteCO2Awal", "CO2 Awal" },
                                { "WasteCO2Terima", "CO2 Terima" },
                                { "WasteCO2Akhir", "CO2 Akhir" },
                                { "WasteCO2Pakai", "CO2 Pakai" },
                                { "WasteConsumptionCO2", "Consumption CO2" },
                                { "WasteTotalGAS", "Total GAS" },
                                { "WasteTotalSTEAM", "TOTAL STEAM" },
                                { "WastePRIMAT", "PRIMAT" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKretekLineFeedingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "fbatchno", "Batch No." },
                                { "fblend", "Blend" },
                                { "fsapBatch", "SAP Batch" },
                                { "fformulaRev", "Formula Rev" },
                                { "fkrosokStart", "Start - Krosok" },
                                { "fkrosokStop", "Stop - Krosok" },
                                { "fkrosokSiloDestination", "SILO Destination - Krosok" },
                                { "fkrosokTotalizer", "Totalizer" },
                                { "frajanganStart", "Start - Rajangan" },
                                { "frajanganStop", "Stop - Rajangan" },
                                { "frajanganSiloDestination", "SILO Destination - Rajangan" },
                                { "frajanganTotalizer", "Totalizer - Rajangan" },
                                { "ResultCm", "Target 16-23(cm)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "Produksi Aktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKretekLineConditioningController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "krosokBatchNo", "Batch No." },
                                { "krosokBlend", "Blend" },
                                { "krosokSapBatch", "SAP Batch" },
                                { "krosokSiloPreblend", "Silo Preblend" },
                                { "krosokSiloTotalBlend", "Silo Total Blend" },
                                { "krosokStart", "" },
                                { "krosokStop", "" },
                                { "krosokDuration", "Duration" },
                                { "Total Tob Wet", "krosokTotalTobWet" },
                                { "Total Tob Dry", "krosokTotalTobDry" },
                                { "krosokCylinderAirTempAvg1", "Cylinder Air Temp Avg 1" },
                                { "krosokCylinderAirTempAvg2", "Cylinder Air Temp Avg 2" },
                                { "krosokTobTempAvg1", "Tob Temp Avg 1" },
                                { "krosokTobTempAvg2", "Tob Temp Avg 2" },
                                { "krosokTotalWater1", "Total Water 1" },
                                { "krosokTotalWater2", "Total Water 2" },
                                { "krosokMcExitOv", "MC Exit OV" },
                                { "krosokCasingAppRate", "Casing App Rate" },
                                { "krosokTotalizer", "Totalizer" },
                                { "rajanganBatchNo", "Batch No." },
                                { "rajanganBlend", "Blend" },
                                { "rajanganSapBatch", "SAP Batch" },
                                { "rajanganSiloTotalBlend", "Silo Total Blend" },
                                { "rajanganStart", "Start" },
                                { "rajanganStop", "Stop" },
                                { "rajanganDuration", "Duration" },
                                { "rajanganTotalTobWet", "Total Tob Wet" },
                                { "rajanganTotalTobDry", "Total Tob Dry" },
                                { "rajanganMcExitOv", "MC Exit OV" },
                                { "rajanganCylinderAirTempAvg1", "Cylinder Air Temp Avg 1" },
                                { "rajanganCylinderAirTempAvg2", "Cylinder Air Temp Avg 2" },
                                { "rajanganTobTempAvg1", "Tob Temp Avg 1" },
                                { "rajanganTobTempAvg2", "Tob Temp Avg 2" },
                                { "rajanganTotalWater1", "Total Water 1" },
                                { "rajanganTotalWater2", "Total Water 2" },
                                { "rajanganCasingAppRate", "Casing App Rate" },
                                { "rajanganTotalizer", "Totalizer" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKretekLineCuttingDryingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "CutterInUse", "Cutter in use" },
                                { "CuttingBatchNo", "Batch No." },
                                { "Cuttingblend", "Blend" },
                                { "Cuttingsapbatch", "SAP Batch" },
                                { "Cuttingstart", "Start" },
                                { "Cuttingstop", "Stop" },
                                { "Cuttingduration", "Duration" },
                                { "Cuttingawal1", "WOC Awal 1" },
                                { "Cuttingawal2", "WOC Awal 2" },
                                { "Cuttingawal3", "WOC Awal 3" },
                                { "Cuttingtengah1", "WOC Tengah 1" },
                                { "Cuttingtengah2", "WOC Tengah 2" },
                                { "Cuttingtengah3", "WOC Tengah 3" },
                                { "Cuttingakhir1", "WOC Akhir 1" },
                                { "Cuttingakhir2", "WOC Akhir 2" },
                                { "Cuttingakhir3", "WOC Akhir 3" },
                                { "Cuttingaveragewoc", "Average" },
                                { "TotalWeigherCombine", "Total Weigher Combine (Kg)" },
                                { "DetectedMaterial2400", "Detected Material (2400)" },
                                { "DetectedMaterial2430", "Detected Material (2430)" },
                                { "DetectedMaterialActual", "Detected Material (Actual)" },
                                { "CuttingMP1", "Mouth Piece Pressure Cutter (bar)" },
                                { "CuttingCS", "Cutter Stop (Stop)" },
                                { "CuttingFTD", "FTD Feeder Stop Frequency (Stop)" },
                                { "CuttingSK", "Sharp Knives (Visual)" },
                                { "DryingBatchNo", "Batch No." },
                                { "Dryingblend", "Blend" },
                                { "Dryingsapbatch", "SAP Batch" },
                                { "Dryingsiloatb", "Silo Asal Total Blend" },
                                { "Dryingsilotcf", "Silo Tujuan CF" },
                                { "DryingTowerType", "Tower Type" },
                                { "Dryingstart", "Start" },
                                { "Dryingstop", "Stop" },
                                { "Dryingduration", "Duration" },
                                { "Dryingtobweigher", "Tob Weigher" },
                                { "Dryingmcinlet", "MC Inlet" },
                                { "Dryingpgt", "PG Temperature" },
                                { "Dryingpgspeed", "PG Speed" },
                                { "Dryingccfs", "Cooling Conv Fan Speed" },
                                { "Dryingmced", "MC Exit Dryer" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKretekLineAddbackController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "SiloAsal", "SILO Asal" },
                                { "SiloTujuan", "SILO Tujuan" },
                                { "TotalWet1", "Total Wet" },
                                { "TotalDry1", "Total Dry" },
                                { "Ov1", "OV" },
                                { "FlowAvg1", "Flow Avg" },
                                { "WtFlowrate", "Flowrate Avg" },
                                { "WtMC", "MC" },
                                { "WtSDMC", "SD MC" },
                                { "FaFlowAvg", "Flowrate Avg - Flavour Application" },
                                { "FaBatchNo", "Batch No - Flavour Application" },
                                { "FaTotQty", "Tot Qty - Flavour Application" },
                                { "CloveFlowAvg", "Flowrate Avg - Clove" },
                                { "CloveBatchNo", "Batch No - Clove" },
                                { "CloveTotalizer", "Totalizer - Clove" },
                                { "CloveMCClove", "MC Clove - Clove" },
                                { "CsfFlowAvg", "Flowrate Avg - CSF" },
                                { "CsfBatchNo", "Batch No - CSF" },
                                { "CsfTotalizer", "Totalizer - CSF" },
                                { "CsfMCCSF", "MC CSF - CSF" },
                                { "DslFlowAvg", "Flowrate Avg - Diet" },
                                { "DslBatchNo", "Batch No - Diet" },
                                { "DslTotalizer", "Totalizer - Diet" },
                                { "DslMCDiet", "MC Diet - Diet" },
                                { "SmallLaminaFlowAvg", "Flowrate Avg - Small Lamina" },
                                { "SmallLaminaBatchNo", "Batch No - Small Lamina" },
                                { "SmallLaminaTotalizer", "Totalizer - Small Lamina" },
                                { "SmallLaminaMC", "MC Small Lamina - Small Lamina" },
                                { "RsFlowAvg", "Flowrate Avg - Ripper Short" },
                                { "RsBatchNo", "Batch No - Ripper Short" },
                                { "RsTotalizer", "Totalizer - Ripper Short" },
                                { "RsMCRs", "MC RS - Ripper Short" },
                                { "CresFlowAvg", "Flowrate Avg - CRES" },
                                { "CresBatchNo", "Batch No - CRES" },
                                { "CresTotalizer", "Totalizer - CRES" },
                                { "CresMCCres", "MC CRES - CRES" },
                                { "RtcFlowAvg", "Flowrate Avg - RTC" },
                                { "RtcBatchNo", "Batch No - RTC" },
                                { "RtcTotalizer", "Totalizer - RTC" },
                                { "RtcMCRtc", "MC RTC - RTC" },
                                { "OosFlowAvg", "Flowrate Avg - Out of Spec" },
                                { "OosBatchNo", "Batch No - Out of Spec" },
                                { "OosTotalizer", "Totalizer - Out of Spec" },
                                { "OosMCRs", "MC RS - Out of Spec" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKretekLinePackingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "AsalSilo", "Asal Silo" },
                                { "TipeBox", "Tipe Box Packing" },
                                { "JumlahBox", "Jumlah Box" },
                                { "Kg", "@Kg" },
                                { "SisaPacking", "Sisa Packing" },
                                { "TotNetto", "Tot Netto" },
                                { "NoBox", "No Box" },
                                { "TimbangUlang", "Timbang Ulang" },
                                { "TanggalReclaim", "Tanggal Reclaim" },
                                { "SesuaiAddback", "Sesuai WPP Addback" },
                                { "TotalSesuai", "Total Sesuai WPP" },
                                { "BebasNTRM", "Bebas NTRM" },
                                { "Barcode", "Barcode Benar, Dilepas & Dikumpulkan ke Admin" },
                                { "Kondisi", "Kondisi Tipper, Doffer & Mini Silo" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryCresInfeedConditioningController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "FormulaRev", "Formula Rev" },
                                { "SiloDestination", "Silo Destination" },
                                { "FlowRateInput", "Flow Rate Input" },
                                { "mcinput", "MC Input" },
                                { "mcexit", "MC Exit" },
                                { "TempExitCond", "Temp Exit Cond" },
                                { "TempWaterApp", "Qty Water App" },
                                { "SiloQuality", "Silo Quantity" },
                                { "WaterSprayNozzle", "Water Spray Nozzle" },
                                { "AtomizingSteam", "Atomizing Steam & Water" },
                                { "ResidenceTime", "Residence Time" },
                                { "abm", "Addback Material Heaviest Cutter" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryCresCutDryPackingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "FormulaRev", "Formula Rev" },
                                { "SiloDestination", "Silo Destination" },
                                { "AirTemp", "Air Temp 1" },
                                { "AirTemp2", "Air Temp 2" },
                                { "AirTemp3", "Air Temp 3" },
                                { "SteamPressure", "Steam Pressure" },
                                { "HoodPressure", "Hood Pressure" },
                                { "MCExitDryer", "MC Exit Dryer" },
                                { "TempExitDryer", "Temp Exit Dryer" },
                                { "WeightAfterDryer", "Weight After Dryer" },
                                { "FlowRateInput", "Flow Rate Input" },
                                { "MCInput", "MC Input" },
                                { "SteamFlowrate", "Steam Flowrate" },
                                { "ElutriatorDumpingOpening", "Elutriator Dumping Opening" },
                                { "WeightBeforeSTS", "Weight Before STS" },
                                { "SelectionCutter", "Selection Cutter CRES" },
                                { "ThicknessHasil", "Thickness Hasil material Flattener" },
                                { "SteamFlatener", "Steam Flatener" },
                                { "TopBandPressure", "Top Band Pressure" },
                                { "MouthHeight", "Mouth Height" },
                                { "CutterStop", "Cutter Stop" },
                                { "SharpKnives", "Sharp Knives" },
                                { "CounterTraverseGerinda", "Counter traverse gerinda" },
                                { "CounterTraversePisau", "Counter traverse pisau" },
                                { "packingnoopspacking", "No Ops Packing" },
                                { "packingstart", "Start" },
                                { "packingstop", "Stop" },
                                { "packingtotal", "Total" },
                                { "packingtotalbox", "Total Box" },
                                { "packingqty", "QTY (Kg)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryKitchenController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SolutionDetail", "Solution Detail" },
                                { "QtyTarget", "Qty Target (kg)" },
                                { "AddingMaterial", "Adding Material" },
                                { "QtyRecipe", "Qty Recipe" },
                                { "AddingType", "Adding Type" },
                                { "ActualQty", "Actual Qty" },
                                { "CekActual", "Cek Actual" },
                                { "StatusAdding", "Status Adding" },
                                { "BatchMaterial", "Batch Material DIM" },
                                { "RefrationMax", "Refration Max" },
                                { "RefrationActual", "Refration Actual" },
                                { "RefrationMin", "Refration Min" },
                                { "SpecificMax", "Specific Max" },
                                { "SpecificActual", "Specific Actual" },
                                { "SpecificMin", "Specific Min" },
                                { "AperanceRemark", "Aperance Remark" },
                                { "OdourRemark", "Odour Remark" },
                                { "PHRemark", "PH Remark" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineFeedingWhiteController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "fwbatchno", "Batch No." },
                                { "fwblend", "Blend" },
                                { "fwsapBatch", "SAP Batch" },
                                { "fwformulaRev", "Formula Rev." },
                                { "fwstart", "Start" },
                                { "fwstop", "Stop" },
                                { "fwDuration", "Duration" },
                                { "fwburleySlice", "Jumlah Slice - Burley" },
                                { "fwburleyTotal", "Totalizer - Burley" },
                                { "fwcastLeafSlice", "Jumlah Slice - Cast Leaf" },
                                { "fwcastLeafTotal", "Totalizer - Cast Leaf" },
                                { "fwflueCureSlice", "Jumlah Slice - Flue Cure" },
                                { "fwflueCureTotal", "Totalizer - Flue Cure" },
                                { "fwburley1", "Burley 1" },
                                { "fwburley2", "Burley 2" },
                                { "fwcastleaf1", "Cast Leaf 1" },
                                { "fwcastleaf2", "Cast Leaf 2" },
                                { "fwfluecure1", "Flue Cure 1" },
                                { "fwfluecure2", "Flue Cure 2" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineDCCCController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "TotalBlendSilo1", "Total Blend SILO 1" },
                                { "TotalBlendSilo2", "Total Blend SILO 2" },
                                { "totalWaterB", "Total/Water - Burley" },
                                { "airTempAvgB", "Air TempAvg - Burley" },
                                { "tobTempAvgB", "Tob TempAvg - Burley" },
                                { "ovExitDCCB", "OVExitDCC - Burley" },
                                { "tTWB", "Total Tob Wet - Burley" },
                                { "tTDB", "Total Tob Dry - Burley" },
                                { "cCB", "CasingCode - Burley" },
                                { "cTB", "CasingTemp - Burley" },
                                { "cTFC", "CasingTemp - Bright" },
                                { "aRAB", "App RateAvg - Burley" },
                                { "aRAFC", "App RateAvg - Bright" },
                                { "totalWaterCL", "Total/Water - Cast Leaf" },
                                { "airTempAvgCL", "Air TempAvg - Cast Leaf" },
                                { "tobTempAvgCL", "Tob TempAvg - Cast Leaf" },
                                { "ovExitDCCCL", "OVExitDCC - Cast Leaf" },
                                { "totalWaterFC", "Total/Water - Flue Cure" },
                                { "airTempAvgFC", "Air TempAvg - Flue Cure" },
                                { "tobTempAvgFC", "Tob TempAvg - Flue Cure" },
                                { "ovExitDCCFC", "OVExitDCC - Flue Cure" },
                                { "tTWFC", "Total Tob Wet - Flue Cure" },
                                { "tTDFC", "Total Tob Dry - Flue Cure" },
                                { "cCFC", "CasingCode - Flue Cure" },
                                { "bCB", "BatchCasing - Burley"},
                                { "bCB2", "BatchCasing - Burley 2"},
                                { "bCFC", "BatchCasing - Bright"},
                                { "bCFC2", "BatchCasing - Bright2"},
                                { "tCB", "TotalCasing - Burley"},
                                { "tCB2", "TotalCasing - Burley2"},
                                { "tCFC", "TotalCasing - Bright"},
                                { "tCFC2", "TotalCasing - Bright2"},
                                { "totalWaterR", "Total/Water - Recon" },
                                { "airTempAvgR", "Air TempAvg - Recon" },
                                { "tobTempAvgR", "Tob TempAvg - Recon" },
                                { "ovExitDCCR", "OVExitDCC - Recon" },
                                { "totalWaterO", "Total/Water - Oriental" },
                                { "airTempAvgO", "Air TempAvg - Oriental" },
                                { "tobTempAvgO", "Tob TempAvg - Oriental" },
                                { "ovExitDCCO", "OVExitDCC - Oriental" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineCuttingFTDController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SAPBatch", "SAP Batch" },
                                { "StartCutting", "Cutting-Start" },
                                { "StopCutting", "Cutting-Stop" },
                                { "TobFlowAvgCutting", "Cutting-Tob Flow Avg" },
                                { "TotalWetTobCutting", "Cutting-Total Wet Tob" },
                                { "StartFTD", "FTD-Start" },
                                { "StopFTD", "FTD-Stop" },
                                { "TobFlowAvgFTD", "FTD-Tob Flow Avg" },
                                { "DischargeSilo", "FTD-Discharge Silo" },
                                { "TotalWetTobFTD", "FTD-Total Wet Tob" },
                                { "InletOVAvg", "Inlet OV Avg" },
                                { "ExitOVAvg", "Exit OV Avg" },
                                { "PGPressure", "PG Pressure" },
                                { "PGTempAvg", "PG Temp Avg" },
                                { "BurnerTempAvg", "Burner Temp Avg" },
                                { "WOCStart1", "WOC-Awal1" },
                                { "WOCStart2", "WOC-Awal2" },
                                { "WOCStart3", "WOC-Awal3" },
                                { "WOCMiddle1", "WOC-Tengah1" },
                                { "WOCMiddle2", "WOC-Tengah2" },
                                { "WOCMiddle3", "WOC-Tengah3" },
                                { "WOCEnd1", "WOC-Akhir1" },
                                { "WOCEnd2", "WOC-Akhir2" },
                                { "WOCEnd3", "WOC-Akhir3" },
                                { "MPPC140", "Mouth Piece Pressure Cutter 140 (Bar)" },
                                { "MPPC141", "Mouth Piece Pressure Cutter 141 (Bar)" },
                                { "CutterStop", "Cutter Stop (Stop)" },
                                { "FTDFeederStopFreq", "FTD Feeder Stop Frequency (Stop)" },
                                { "SharpKnives", "Sharp Knives (Visual)" },
                                { "SharpKnives140", "Sharp Knives 140 (Visual)" },
                                { "SharpKnives141", "Sharp Knives 141 (Visual)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineAddbackController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                //Master
                                { "BatchNo", "PSS Batch" },
                                { "SapBatch", "SAP Batch" },
                                { "SiloDestination", "SILO Destination" },
                                { "TotalWet1", "Master-Total Wet" },
                                { "TotalDry1", "Master-Total Dry" },
                                { "Ov1", "Master-OV" },
                                { "FlowAvg1", "Master-Flow Avg" },

                                //Ripper Short
                                { "BatchNo1", "Ripper Short-Batch No" },
                                { "TotalWet2", "Ripper Short-Total Wet" },
                                { "TotalDry2", "Ripper Short-Total Dry" },
                                { "Ov2", "Ripper Short-OV" },
                                { "FlowAvg2", "Ripper Short-Flow Avg" },
                                { "Bn1", "Ripper Short-Batch No 1" },
                                { "Bn2", "Ripper Short-Batch No 2" },
                                { "Bn3", "Ripper Short-Batch No 3" },
                                { "Ba1", "Ripper Short-Blend Addback 1" },
                                { "Ba2", "Ripper Short-Blend Addback 2" },
                                { "Ba3", "Ripper Short-Blend Addback 3" },
                                { "Tb1", "Ripper Short-Total Box 1" },
                                { "Tb2", "Ripper Short-Total Box 2" },
                                { "Tb3", "Ripper Short-Total Box 3" },
                                { "TotalQuantity1", "Ripper Short-Total Qty 1" },
                                { "TotalQuantity2", "Ripper Short-Total Qty 2" },
                                { "TotalQuantity3", "Ripper Short-Total Qty 3" },
                                
                                //Small Lamina
                                { "BatchNo2", "Small Lamina-Batch No" },
                                { "TotalWet3", "Small Lamina-Total Wet" },
                                { "TotalDry3", "Small Lamina- Total Dry" },
                                { "Ov3", "Small Lamina-OV" },
                                { "FlowAvg3", "Small Lamina-Flow Avg" },
                                { "BlendAddback1", "Small Lamina-Blend Addback" },
                                { "TotalBox1", "Small Lamina-Total Box" },

                                //Expanded Tob
                                { "BlendAddback3", "Expanded Tob-Blend Addback" },
                                { "TotalWet5", "Expanded Tob-Total Wet" },
                                { "TotalBox3", "Expanded Tob-Total Box" },
                                { "BatchNo4", "Expanded Tob-Batch No" },
                                { "TotalDry5", "Expanded Tob-Total Dry" },
                                { "Ov5", "Expanded Tob-OV" },
                                { "FlowAvg5", "Expanded Tob-Flow Avg" },
                                { "SiLoLinkUpExTob1", "Expanded Tob-Silo LU" },
                                { "SiLoLinkUpExTob2", "Expanded Tob-Silo LU2" },

                                 //IS
                                { "BatchNo3", "IS-Batch No" },
                                { "TotalWet4", "IS-Total Wet" },
                                { "TotalDry4", "IS-Total Dry" },
                                { "Ov4", "IS-OV" },
                                { "FlowAvg4", "IS-Flow Avg" },
                                { "BlendAddback2", "IS-Blend Addback" },
                                { "TotalBox2", "IS-Total Box" },
                                { "SiloLinkUpIS1", "IS-Silo LU" },
                                { "SiloLinkUpIS2", "IS-Silo LU2" },

                                //Flavour
                                { "TotalQtyFlavour1", "Flavour-Tot Qty" },
                                { "FlavourBatchID1", "Flavour-Batch No" },
                                { "TotalQtyFlavour2", "Flavour-Tot Qty" },
                                { "FlavourBatchID2", "Flavour-Batch No" },
                                { "FlavourCode", "Flavour-Blend Addback" },

                                //Riftab
                                { "BatchNoRiftab", "RIFTAB-Batch No" },
                                { "TotalWetRiftab", "RIFTAB-Total Wet" },
                                { "TotalDryRiftab", "RIFTAB-Total Dry" },
                                { "OvRiftab", "RIFTAB-Ov" },
                                { "FlowAvgRiftab", "RIFTAB-RIFTAB-Flow Avg" },
                                { "BlendAddbackRiftab", "RIFTAB-Blend Addback" },
                                { "TotalBoxRiftab", "RIFTAB-Total Box" },
                                { "SiloLinkUpRiftab", "RIFTAB-Silo LU" },
                                { "SiloLinkUpRiftab2", "RIFTAB-Silo LU2" },
                                
                                //Final
                                { "TotalWetFinal", "Total Wet Final" },
                                { "FinalOv", "Final OV" },

                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLinePackingWhiteController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BoxBatchNo", "Batch No." },
                                { "BoxBlend", "Blend" },
                                { "BoxSapBatch", "SAP Batch" },
                                { "BoxSap", "Silo Asal Packing" },
                                { "BoxTypeBox", "Tipe Box Packing" },
                                { "BoxStart", "Waktu Packing-Start" },
                                { "BoxStop", "Waktu Packing-Stop" },
                                { "BoxJumlahBox", "Hasil Packing-Jumlah Box" },
                                { "BoxKg", "Hasil Packing-Kg" },
                                { "BoxTotNetto", "Hasil Packing-Tot Netto" },
                                { "BoxRemarks", "Hasil Packing-Remarks" },
                                { "BoxDuration", "Duration" },
                                { "CLUBatchNo", "Batch No. " },
                                { "CLUBlend", "Blend" },
                                { "CLUSapBatch", "SAP Batch" },
                                { "CLUSiloAsal", "Silo Asal" },
                                { "CLUStart", "Waktu Link Up-Start" },
                                { "CLUStop", "Waktu Link Up-Stop" },
                                { "CLUTopFeeder1", "Jacobi Feeder-Top Feeder 1" },
                                { "CLUTopFeeder2", "Jacobi Feeder-Top Feeder 2" },
                                { "CLUTotKg", "Jacobi Feeder-Total Kg" },
                                { "CLURemarks", "Jacobi Feeder-Remarks" },
                                { "CLUDuration", "Duration" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineFeedingSPMController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "spmbatchno", "Batch No." },
                                { "spmproddate", "Prod Date" },
                                { "spmexpdate", "Exp Date" },
                                { "spmjumlahcase", "Jumlah Case" },
                                { "spmjumlahqty", "Jumlah Qty" },
                                { "spmcaseno", "Case No" },
                                { "spmjam", "Jam" },
                                { "spmberat", "Berat" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryISWhiteFeedingController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No. " },
                                { "SapBatch", "SAP Batch" },
                                { "FormulaRev", "Formula Rev." },
                                { "RowRateInput", "Flow Rate Input" },
                                { "MCInput", "MC Input" },
                                { "MCExit", "MC Exit" },
                                { "TmpExitCond", "Temp Exit Cond" },
                                { "TotalWaterApp", "Total Water App" },
                                { "SiloQuantity", "Silo Quantity" },
                                { "WaterSprayNozzle", "Water Spray nozzle" },
                                { "AtomSteamWater", "Atomizing Steam & water" },
                                { "RecidenceTime", "Recidence Time (detik)" },
                                { "AddbackHeavistCutter", "Add back Material Heavist Cutter (kg)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryISWhiteCutDryController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "BatchNo", "Batch No." },
                                { "SapBatch", "SAP Batch" },
                                { "NumberCutter", "Number Cutter" },
                                { "SelectionCutter", "Selection Cutter" },
                                { "SiloDestination", "Silo Destination" },
                                { "FlowInput", "STS-Flow Input" },
                                { "OVInput", "STS-OV Input" },
                                { "SteamFlowrate", "STS-Steam Flowrate" },
                                { "WeightInput", "STS-Weight Input" },
                                { "TempInput", "FBD-Temp Input" },
                                { "AirTempZona1", "FBD-Air Temp Zone 1" },
                                { "AirTempZona2", "FBD-Air Temp Zone 2" },
                                { "OVExit", "FBD-OV Exit" },
                                { "WeightExit", "FBD-Weight Exit" },
                                { "TF1", "Thickness Flattener 1" },
                                { "TF2", "Thickness Flattener 2" },
                                { "TF3", "Thickness Flattener 3" },
                                { "TF4", "Thickness Flattener 4" },
                                { "TF5", "Thickness Flattener 5" },
                                { "TF6", "Thickness Flattener 6" },
                                { "TF7", "Thickness Flattener 7" },
                                { "TF8", "Thickness Flattener 8" },
                                { "TF9", "Thickness Flattener 9" },
                                { "TF10", "Thickness Flattener 10" },
                                { "TF11", "Thickness Flattener 11" },
                                { "TF12", "Thickness Flattener 12" },
                                { "WC1", "WOC Cutter 1" },
                                { "WC2", "WOC Cutter 2" },
                                { "WC3", "WOC Cutter 3" },
                                { "WC4", "WOC Cutter 4" },
                                { "WC5", "WOC Cutter 5" },
                                { "WC6", "WOC Cutter 6" },
                                { "WC7", "WOC Cutter 7" },
                                { "WC8", "WOC Cutter 8" },
                                { "WC9", "WOC Cutter 9" },
                                { "WC10", "WOC Cutter 10" },
                                { "TargetThickness", "Thickness Exit Flattener (mm)" },
                                { "TargetSteam", "Steam flattener" },
                                { "TargetPressure", "Top Band Pressure (Bar)" },
                                { "TargetHeight", "Mouth Height (mm)" },
                                { "TargetWOC", "WOC (mm)" },
                                { "TargetCutter", "Cutter Stop (Stop)" },
                                { "TargetSharp", "Sharp Knives (Visual)" },
                                { "TargetCounter", "Counter Traverse Pisau (Travense)" },
                                { "TargetHeaviest", "Heaviest (Kg)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                    {"LPHPrimaryWhiteLineOTPController",
                        (
                            new Dictionary<string, string>
                            {
                                { "TeamLeader", "Supervisor" }
                            },
                            new Dictionary<string, string>
                            {
                                { "piBatchID", "Batch No. " },
                                { "piBlendCode", "Blend Code" },
                                { "piPKBatch", "PkBatch" },
                                { "bcoBlendBefore", "BCO Casing-Blend Code dari Batch Sebelumnya" },
                                { "bcoCasingBeda", "BCO Casing-Apakah Code Casing beda dengan sebelumnya ?" },
                                { "bcoCleanTank", "BCO Casing-Bila [Ya]: Bersihkan App Tank dengan air" },
                                { "sccSiloDestination", "Slicer, Conditioning & Casing-Silo Destination" },
                                { "sccCasingBatch", "Slicer, Conditioning & Casing-Casing Batch ID" },
                                { "sccBSBatch", "Slicer, Conditioning & Casing-BS Batch ID" },
                                { "cutdrySiloDestination", "Cutting & Drying-Silo Destination-Silo Destination" },
                                { "cutWidth1", "Cutting & Drying-Cut Width (sampling)1" },
                                { "cutWidth2", "Cutting & Drying-Cut Width (sampling)2" },
                                { "cutWidth3", "Cutting & Drying-Cut Width (sampling)3" },
                                { "cutWidth4", "Cutting & Drying-Cut Width (sampling)4" },
                                { "cutWidth5", "Cutting & Drying-Cut Width (sampling)5" },
                                { "cutselection1", "Cutting & Drying-For RYO - Cutter Selection1" },
                                { "cutselection2", "Cutting & Drying-For RYO - Cutter Selection2" },
                                { "cutselection3", "Cutting & Drying-For RYO - Cutter Selection3" },
                                { "cutselection4", "Cutting & Drying-For RYO - Cutter Selection4" },
                                { "MAFETActual", "Master, Addback & Flavoring-ET Actual Applied" },
                                { "MAFCresActual", "Master, Addback & Flavoring-CRES Actual Applied" },
                                { "MAFFlavorBatch", "Master, Addback & Flavoring-Flavor Batch ID" },
                                { "AddbackBlendCode", "Addback-Blend Code" },
                                { "AddbackBatchID", "Addback-Batch ID" },
                                { "AddbackQty", "Addback-Qty" },
                                { "bcoFlavorBatchSama", "BCO Flavor-Batch Sebelumnya menggunakan Code Flavor berbeda?" },
                                { "bcoFlavorCleanTank", "BCO Flavor-Bersihkan App Tank dengan air" },
                                { "bcoFlavorCleanCilinder", "BCO Flavor-Bersihkan Flavor Cylinder dari sisa tembakau" },
                                { "bcoFlavorCleanConveyor", "BCO Flavor-Bersihkan Conveyor setelah Flavor Cyl dari tembakau" },
                                { "bcoFlavorCleanBlending", "BCO Flavor-Bersihkan Blending Silo dari tembakau" },
                                { "bcoFlavorCleanBulking", "BCO Flavor-Bersihkan Bulking Silo dari tembakau" },
                                { "bcoFlavorNextBatchUseFlavorCleanTank", "BCO Flavor-Bersihkan App Tank dengan air dan PG" },
                                { "bcoFlavorNextBatchUseFlavorCleanCylinder", "BCO Flavor-Bersihkan Flavor Cylinder dengan air dan PG" },
                                { "bcoFlavorNextBatchUseFlavorCleanConveyor", "BCO Flavor-Bersihkan Conveyor setelah FC dengan air dan PG" },
                                { "bcoFlavorNextBatchUseFlavorCleanBlending", "BCO Flavor-Bersihkan Blending Silo dengan air dan PG" },
                                { "bcoFlavorNextBatchUseFlavorCleanBulking", "BCO Flavor-Bersihkan Bulking Silo dengan air dan PG" },
                                { "bcoPackingBatchBerbedaCleanFilling", "BCO Packing-Bersihkan Filling Station dari sisa tembakau" },
                                { "bcoPackingBatchBerbedaCleanConv", "BCO Packing-Bersihkan Conv. setelah Packing Silo dari sisa material" },
                                { "bcoPackingNextUseFlavorCleanFilling", "BCO Packing-Bersihkan Filling Station dengan air dan PG" },
                                { "bcoPackingNextUseFlavorCleanConv", "BCO Packing-Bersihkan Conv setelah Bulking Silo dengan air dan PG" },
                                { "bcoPackingDate", "Packing-Packing date" },
                                { "bcoPackingDateStart", "Packing-Start" },
                                { "bcoPackingDateStop", "Packing-Stop" },
                                { "bcoPackingMCPacking", "Packing-MC Packing" },
                                { "bcoPackingTaraBox", "Packing-Tara Box" },
                                { "bcoPackingTotal", "Packing-Total Packing (Box)" },
                                { "bcoPackingBeratSatuan", "Packing-Total Packing (Kg)" },
                                { "bcoPackingPlusKilo", "Packing-Total Packing (Receh)" },
                                { "WasteType", "Waste Type" },
                                { "Process", "Proses (Blend)" },
                                { "TargetProduksi", "Target Produksi" },
                                { "ProduksiAktual", "ProduksiAktual" },
                                { "InformasiChangeover", "Informasi Changeover" },
                                { "InitialCause", "Initial Root Cause (Deskripsikan issue yang terjadi)" },
                                { "ThisAction", "Action Taken" },
                                { "ThisOwner", "Owner" },
                                { "ThisStatus", "Status" },
                                { "NextAction", "Action Taken" },
                                { "NextOwner", "Owner" },
                                { "NextStartTime", "Start Time" },
                            }
                        )},
                };

                Dictionary<string, List<List<string>>> sheets = new Dictionary<string, List<List<string>>>();

                lphHeader.ForEach(type =>
                {
                    string lphType = type.Replace("LPHPrimary", "").Replace("Controller", "");

                    List<PPLPHSubmissionsModel> subListPerType = submissions.Where(x => x.LPHHeader == type).ToList();
                    if (batch != null && batch != "")
                    {
                        List<long> lphIds = extraList.Where(x => subListPerType.Any(y => y.LPHID == x.LPHID) && 
                            x.FieldName.Trim().ToLower().Contains("batchno") && x.Value == batch).Select(x => x.LPHID).Distinct().ToList();
                        subListPerType = submissions.Where(x => lphIds.Any(y => y == x.LPHID)).ToList();
                    }

                    Dictionary<string, string> renamingExtra = null, renamingComponent = null;
                    if(renaming.ContainsKey(type)) (renamingComponent, renamingExtra) = renaming[type];

                    List<string> filterExtra = extraList.Where(x => subListPerType.Any(y => y.LPHID == x.LPHID)).Select(x => x.HeaderName).Distinct().ToList();
                    List<string> filterComponent = new List<string>
                    {
                        "SelectedWasteWrapper",
                    };

                    List<string> headerComponent = new List<string>();
                    List<Dictionary<string, string>> comvals = new List<Dictionary<string, string>>();
                    foreach (var sub in subListPerType)
                    {
                        List<PPLPHComponentsModel> header = componentList.Where(x => x.LPHID == sub.LPHID && !filterComponent.Any(y => x.ComponentName.StartsWith(y))).ToList();
                        List<string> head = header.Select(x => x.ComponentName).Distinct().ToList();
                        if (head.Count() > headerComponent.Count())
                        {
                            headerComponent = head;
                            headerComponent.Insert(0, "LPH ID");
                        }

                        Dictionary<string, string> comval = new Dictionary<string, string>()
                        {
                            { "LPH ID", sub.LPHID.ToString() }
                        };
                        header.ForEach(h =>
                        {
                            string value = valueList.Where(v => h.ID == v.LPHComponentID).Select(x => x.Value).FirstOrDefault();
                            if (value != null)
                            {
                                if (!comval.ContainsKey(h.ComponentName))
                                    comval.Add(h.ComponentName, value);
                                else comval[h.ComponentName] = value;
                            }
                        });
                        comvals.Add(comval);
                    }
                    List<List<string>> componentRows = new List<List<string>>();
                    comvals.ForEach(cv =>
                    {
                        List<string> row = new List<string>();
                        headerComponent.ForEach(h =>
                        {
                            if (cv.ContainsKey(h)) row.Add(cv[h]);
                            else row.Add(null);
                        });
                        componentRows.Add(row);
                    });
                    if(renamingComponent != null)
                    {
                        headerComponent = headerComponent.Select(h =>
                        {
                            if (h == "LPH ID") return h;
                            h = string.Join("-", h.Split('-').Skip(1));
                            if (renamingComponent.ContainsKey(h))
                                return renamingComponent[h];
                            return h;
                        }).ToList();
                    }
                    componentRows.Insert(0, headerComponent);
                    sheets.Add(lphType, componentRows);

                    filterExtra.ForEach(m =>
                    {
                        List<Dictionary<string, string>> extras = new List<Dictionary<string, string>>();
                        List<PPLPHExtrasModel> extra = extraList.Where(x => x.HeaderName == m).ToList();
                        List<string> headerRow = extra.Select(x => x.FieldName).Distinct().ToList();
                        foreach (var sub in subListPerType)
                        {
                            List<PPLPHExtrasModel> e = extra.Where(x => x.LPHID == sub.LPHID).ToList();
                            for (int i = 0; e.Count != 0; i++)
                            {
                                int tableRowIndex = extras.Count();
                                int subTableRowMaxIndex = extras.Count();
                                extras.Add(new Dictionary<string, string>()
                                {
                                    { "LPH ID", sub.LPHID.ToString() }
                                });
                                for (int j = 0; j < e.Count; j++)
                                {
                                    if (e[j].RowNumber == i)
                                    {
                                        if (e[j].ValueType.Trim() == "Json")
                                        {
                                            string val = e[j].Value;
                                            if (IsValidJson(val))
                                            {
                                                List<Dictionary<string, string>> jsonD =
                                                    JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(val);
                                                int hi = headerRow.FindIndex(x => x == e[j].FieldName);
                                                if (hi >= 0 && jsonD.Count > 0)
                                                {
                                                    headerRow.RemoveAt(hi);
                                                    List<string> keys = jsonD[0].Keys.ToList();
                                                    for (int k = keys.Count - 1; k >= 0; k--)
                                                        if (keys[k] != "id") headerRow.Insert(hi, keys[k]);
                                                }
                                                int subTableRowIndex = tableRowIndex;
                                                int subTableStartFrom = extras[tableRowIndex].Count();
                                                foreach (var json in jsonD)
                                                {
                                                    if (subTableRowMaxIndex >= subTableRowIndex)
                                                    {
                                                        foreach (KeyValuePair<string, string> item in json.OrderBy(x => x.Key))
                                                            if (item.Key != "id" && !extras[subTableRowIndex].ContainsKey(item.Key))
                                                                extras[subTableRowIndex].Add(item.Key, item.Value);
                                                    }
                                                    else
                                                    {
                                                        Dictionary<string, string> newRow = new Dictionary<string, string>();
                                                        foreach (KeyValuePair<string, string> item in json.OrderBy(x => x.Key))
                                                            if (item.Key != "id" && !newRow.ContainsKey(item.Key))
                                                                newRow.Add(item.Key, item.Value);
                                                        extras.Add(newRow);
                                                        subTableRowMaxIndex++;
                                                    }
                                                    subTableRowIndex++;
                                                }
                                            }
                                        }
                                        else if (!extras[tableRowIndex].ContainsKey(e[j].FieldName)) 
                                            extras[tableRowIndex].Add(e[j].FieldName, e[j].Value);
                                        e.RemoveAt(j);
                                        j--;
                                    }
                                }
                            }
                        }
                        headerRow.Reverse();
                        headerRow.Insert(0, "LPH ID");
                        List<List<string>> extraRows = new List<List<string>>();
                        extras.ForEach(e =>
                        {
                            List<string> row = new List<string>();
                            headerRow.ForEach(h =>
                            {
                                if (e.ContainsKey(h)) row.Add(e[h]);
                                else row.Add(null);
                            });
                            extraRows.Add(row);
                        });
                        if (renamingExtra != null)
                        {
                            headerRow = headerRow.Select(h =>
                            {
                                if (renamingExtra.ContainsKey(h))
                                    return renamingExtra[h];
                                return h;
                            }).ToList();
                        }
                        extraRows.Insert(0, headerRow);
                        sheets.Add(lphType.Replace("Conditioning", "DCCC").Replace("CuttingDrying", "CutDry") + "-" + char.ToUpper(m[0]) + m.Substring(1), extraRows);
                    });
                });

                return Json(new { data = sheets }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                return Json(new { error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult ExtractExcel(string data)
        {
            try
            {
                Dictionary<string, List<List<string>>> sheets = string.IsNullOrEmpty(data) ? 
                    new Dictionary<string, List<List<string>>>() : 
                    JsonConvert.DeserializeObject<Dictionary<string, List<List<string>>>>(data);
                Session["ExtractRawDataPP"] = ExcelGenerator.PPRawDataExtract(sheets);

                return Json(new { status = true }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { status = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult DownloadExcel()
        {
            if (Session["ExtractRawDataPP"] != null)
            {
                byte[] data = Session["ExtractRawDataPP"] as byte[];
                Session["ExtractRawDataPP"] = null;
                return File(data, "application/octet-stream", "RawDataPP.xlsx");
            }
            return new EmptyResult();
        }

    }
}
