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
    public class ReportPPExportSAPController : BaseController<PPLPHModel>
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

        public ReportPPExportSAPController(
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

        public class Filter
        {
            public List<ComponentSheet> SheetByComponents;
            public List<ExtraSheet> SheetByExtras;
        }
        public class ComponentSheet
        {
            public string SheetName;
            public List<string> Displayed;
        }
        public class ExtraSheet
        {
            public string SheetName;
            public Dictionary<string, List<string>> JoinBy;
            public List<string> TableNames;
            public List<string> Selected;
            public List<string> Displayed;
            public List<string> Accumulated;
            public List<string> Summarized;
            public Func<
                    List<Dictionary<string, string>>, 
                    IPPLPHSubmissionsAppService,
                    IPPLPHExtrasAppService, 
                    List<Dictionary<string, string>>
                > BeforeDisplay;
        }
        private readonly Dictionary<string, Filter> SapMap = new Dictionary<string, Filter>()
        {
            {"LPHPrimaryRTCController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Feeding RTC",
                            TableNames = new List<string> { "feeding" },
                            JoinBy = new Dictionary<string, List<string>> { { "FeedingBatchNo", new List<string> { } } },
                            Selected = new List<string> { "FeedingBatchNo", "Line1234OB4", "Line1234Total", "Line56OB1", "Line56Total" },
                            Displayed = new List<string> { "LPHID", "Date", "FeedingBatchNo", "Line1234OB4", "Line1234Total", "Line56OB1", "Line56Total" }
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Packing RTC",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "PackingBatchNo", new List<string> { } } },
                            Selected = new List<string> { "PackingBatchNo", "PackingStart", "PackingStop", "PackingTotalBox", "PackingTotalKg" },
                            Displayed = new List<string> { "LPHID", "Date", "PackingBatchNo", "PackingStart", "PackingStop", "PackingTotalBox", "PackingTotalKg" }
                        }
                    }
                } },
            {"LPHPrimaryKretekLinePackingController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Packing",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNo", new List<string> { } } },
                            Selected = new List<string> { "BatchNo", "Blend", "SapBatch", "JumlahBox", "Kg", "SisaPacking", "TotNetto", "Remarks" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNo", "Blend", "SapBatch", "JumlahBox", "Kg", "SisaPacking", "TotNetto", "Remarks" }
                        }
                    }
                } },
            {"LPHPrimaryKretekLineAddbackController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Addback",
                            TableNames = new List<string> { "addback" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNo", new List<string> { } } },
                            Selected = new List<string> { "BatchNo", "Blend", "No PO", "smallLamina", "flavourApplication", "clove", "csf", "cres", "rtc", "oos" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNo", "Blend", "No PO", "SmallLaminaFlowAvg", "SmallLaminaBatchNo", "SmallLaminaTotalizer", "SmallLaminaMC", "FaTotQty", "CloveTotalizer", "CsfTotalizer", "CresTotalizer", "RtcTotalizer", "OosTotalizer", "Cassing" }
                        },
                    }
                } },
            {"LPHPrimaryDietController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Packing",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "PackingBlendCode", new List<string> { } } },
                            Selected = new List<string> { "PackingSamplingError", "PackingSamplingTimbang", "PackingSamplingQty", "PackingSamplingNoBox", "PackingSamplingNo", "PackingQtyKg", "PackingQtyBox", "PackingVisual", "PackingJenisBox", "PackingDuration", "PackingStopPacking", "PackingStartPacking", "PackingSiloOrigin", "PackingBlendCode" },
                            Displayed = new List<string> { "LPHID", "Date", "PackingSamplingError", "PackingSamplingTimbang", "PackingSamplingQty", "PackingSamplingNoBox", "PackingSamplingNo", "PackingQtyKg", "PackingQtyBox", "PackingVisual", "PackingJenisBox", "PackingDuration", "PackingStopPacking", "PackingStartPacking", "PackingSiloOrigin", "PackingBlendCode" }
                        },
                    }
                } },
            {"LPHPrimaryWhiteLineOTPController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Whiteline OTP",
                            TableNames = new List<string> { "sis" },
                            JoinBy = new Dictionary<string, List<string>> { { "bcoPackingPlusKilo", new List<string> { } } },
                            Selected = new List<string> { "bcoPackingPlusKilo", "addback", "bcoPackingTotal", "bcoPackingTaraBox", "bcoBlendBefore", "piPKBatch", "piBlendCode", "piBatchID" },
                            Displayed = new List<string> { "LPHID", "Date", "bcoPackingPlusKilo", "AddbackQty", "AddbackBatchID", "AddbackBlendCode", "bcoPackingTotal", "bcoPackingTaraBox", "bcoBlendBefore", "piPKBatch", "piBlendCode", "piBatchID" }
                        },
                    }
                } },
            {"LPHPrimaryKitchenController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Kitchen",
                            TableNames = new List<string> { "kitchen" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNo", new List<string> { } } },
                            Selected = new List<string> { "BatchNo", "Solution", "QtyTarget", "ingredient", "ingredient", "Blend", "Ops" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNo", "Solution", "QtyTarget", "AddingMaterial", "QtyRecipe", "ActualQty", "Satuan", "Blend", "Ops" }
                        },
                    }
                } },
            {"LPHPrimaryCSFCutDryPackingController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Packing",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "packinglinenoopspacking", new List<string> { } } },
                            Selected = new List<string> { "packinglinenoopspacking", "packinglinestart", "packinglinestop", "packinglinetotal", "packinglinetotalbox", "packinglineqty" },
                            Displayed = new List<string> { "LPHID", "Date", "packinglinenoopspacking", "packinglinestart", "packinglinestop", "packinglinetotal", "packinglinetotalbox", "packinglineqty" }
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Cutting Drying + Packing",
                            TableNames = new List<string> { "cuttingCsf" },
                            JoinBy = new Dictionary<string, List<string>> { { "cuttingTotalizerflavour", new List<string> { } } },
                            Selected = new List<string> { "cuttingTotalizerflavour", "cuttingTotalizermaterial", "cuttingSilodestination", "cuttingStart", "cuttingStop", "cuttingBlend", "cuttingSapBatch", "cuttingBatchNo" },
                            Displayed = new List<string> { "LPHID", "Date", "cuttingTotalizerflavour", "cuttingTotalizermaterial", "cuttingSilodestination", "cuttingStart", "cuttingStop", "cuttingBlend", "cuttingSapBatch", "cuttingBatchNo" }
                        },
                    }
                } },
            {"LPHPrimaryCloveInfeedConditioningController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Infeed",
                            TableNames = new List<string> { "infeed" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNoInfeed", new List<string> { } } },
                            Selected = new List<string> { "BatchNoInfeed", "BlendInfeed", "SAPBatchInfeed", "SiloDestinationInfeed", "QuantityWeighterInfeed", "DurationInfeed" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNoInfeed", "BlendInfeed", "SAPBatchInfeed", "SiloDestinationInfeed", "QuantityWeighterInfeed", "DurationInfeed" }
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Conditioning",
                            TableNames = new List<string> { "conditioning" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNoDCC", new List<string> { } } },
                            Selected = new List<string> { "BatchNoDCC", "BlendDCC", "conditioningSapbatch", "CCS1DCC", "CCS2DCC", "TWA1DCC", "TWA2DCC" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNoDCC", "BlendDCC", "conditioningSapbatch", "CCS1DCC", "CCS2DCC", "TWA1DCC", "TWA2DCC" }
                        },
                    }
                } },
            {"LPHPrimaryCloveCutDryPackingController", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Cutting",
                            TableNames = new List<string> { "cutting" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNo", new List<string> { } } },
                            Selected = new List<string> { "BatchNo", "Blend", "TotalWeyconMaster0", "AddbackAngkup0", "MCFinal0" },
                            Displayed = new List<string> { "LPHID", "Date", "BatchNo", "Blend", "TotalWeyconMaster0", "AddbackAngkup0", "MCFinal0" }
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Packing Clove",
                            TableNames = new List<string> { "packingclove" },
                            JoinBy = new Dictionary<string, List<string>> { { "packingclovenoopspacking", new List<string> { } } },
                            Selected = new List<string> { "packingclovenoopspacking", "packingclovestart", "packingclovestop", "packingclovetotal", "packingclovetotalbox", "packingcloveqty" },
                            Displayed = new List<string> { "LPHID", "Date", "packingclovenoopspacking", "packingclovestart", "packingclovestop", "packingclovetotal", "packingclovetotalbox", "packingcloveqty" }
                        },
                    }
                } },
            {"Uptime", new Filter()
                {
                    SheetByExtras = new List<ExtraSheet>
                    {
                        new ExtraSheet()
                        {
                            SheetName = "Kretek Line Drying",
                            TableNames = new List<string> { "drying" },
                            JoinBy = new Dictionary<string, List<string>> { { "DryingBatchNo", new List<string> { } } },
                            Selected = new List<string> { "DryingBatchNo", "Dryingblend", "DryingFinalWeigherTotal" },
                            Displayed = new List<string> { "DryingBatchNo", "Dryingblend", "DryingFinalWeigherTotal" },
                            Accumulated = new List<string> { "DryingFinalWeigherTotal" },
                            Summarized = new List<string> { "DryingFinalWeigherTotal" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Clove Line Cutting",
                            TableNames = new List<string> { "cutting" },
                            JoinBy = new Dictionary<string, List<string>> { { "BatchNo", new List<string> { } } },
                            Selected = new List<string> { "BatchNo", "Blend", "TotalWeyconMaster0" },
                            Displayed = new List<string> { "BatchNo", "Blend", "TotalWeyconMaster0" },
                            Accumulated = new List<string> { "TotalWeyconMaster0" },
                            Summarized = new List<string> { "TotalWeyconMaster0" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "RTC Line Packing",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "PackingBatchNo", new List<string> { } } },
                            Selected = new List<string> { "PackingBatchNo", "PackingTotalKg" },
                            Displayed = new List<string> { "PackingBatchNo", "PackingTotalKg" },
                            Accumulated = new List<string> { "PackingTotalKg" },
                            Summarized = new List<string> { "PackingTotalKg" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "CSF Line Cutting",
                            TableNames = new List<string> { "cuttingCsf" },
                            JoinBy = new Dictionary<string, List<string>> { { "cuttingBatchNo", new List<string> { } } },
                            Selected = new List<string> { "cuttingBatchNo", "cuttingBlend", "cuttingTotalizermaterial" },
                            Displayed = new List<string> { "cuttingBatchNo", "cuttingBlend", "cuttingTotalizermaterial" },
                            Accumulated = new List<string> { "cuttingTotalizermaterial" },
                            Summarized = new List<string> { "cuttingTotalizermaterial" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "Diet Packing",
                            TableNames = new List<string> { "packing" },
                            JoinBy = new Dictionary<string, List<string>> { { "PackingBlendCode", new List<string> { } } },
                            Selected = new List<string> { "PackingBlendCode", "PackingQtyKg" },
                            Displayed = new List<string> { "PackingBlendCode", "PackingQtyKg" },
                            Accumulated = new List<string> { "PackingQtyKg" },
                            Summarized = new List<string> { "PackingQtyKg" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "White Line OTP",
                            TableNames = new List<string> { "sis" },
                            JoinBy = new Dictionary<string, List<string>> { { "piBatchID", new List<string> { } } },
                            Selected = new List<string> { "piBatchID", "piBlendCode", "bcoTotalOutput" },
                            Displayed = new List<string> { "piBatchID", "piBlendCode", "bcoTotalOutput" },
                            Accumulated = new List<string> { "bcoTotalOutput" },
                            Summarized = new List<string> { "bcoTotalOutput" },
                        },
                        new ExtraSheet()
                        {
                            SheetName = "CRES Line",
                            TableNames = new List<string> { "packingcres", "stem" },
                            JoinBy = new Dictionary<string, List<string>> { { "packingnoopspacking", new List<string> { "BatchNo" } } },
                            Selected = new List<string> { "packingnoopspacking", "packingqty", "Blend", "BatchNo" },
                            Displayed = new List<string> { "packingnoopspacking", "Blend", "packingqty" },
                            Accumulated = new List<string> { "packingqty" },
                            Summarized = new List<string> { "packingqty" },
                            BeforeDisplay = (extras, submissionService, extraService) =>
                            {
                                for(int i = 0; i < extras.Count(); i++)
                                {
                                    if(extras[i].ContainsKey("packingnoopspacking") && extras[i]["packingnoopspacking"] == "AA36D")
                                    {
                                        string date = extras[i]["Date"];
                                        string shift = extras[i]["Shift"];
                                        List<QueryFilter> submissionsFilter = new List<QueryFilter>
                                        {
                                            new QueryFilter("Date", extras[i]["Date"]),
                                            new QueryFilter("Shift", extras[i]["Shift"]),
                                            new QueryFilter("LocationID", extras[i]["LocationID"]),
                                            new QueryFilter("LPHHeader", "LPHPrimaryCresFeedingConditioningController")
                                        };
                                        PPLPHSubmissionsModel sub = submissionService.Get(submissionsFilter).DeserializeToPPLPHSubmissions();
                                        if(sub != null)
                                        {
                                            List<QueryFilter> extraFilter = new List<QueryFilter>
                                            {
                                                new QueryFilter("FieldName", "Totalizer"),
                                                new QueryFilter("FieldName", "Blend", Operator.Equals, Operation.OrElse),
                                                new QueryFilter("HeaderName", "stem"),
                                                new QueryFilter("LPHID", sub.LPHID),
                                            };
                                            List<PPLPHExtrasModel> extraSub = extraService.Find(extraFilter).DeserializeToPPLPHExtrasList();
                                            List<PPLPHExtrasModel> extraWithBlendAA36D = extraSub.Where(x => x.FieldName == "Blend" && x.Value == "AA36D").ToList();
                                            if(extraWithBlendAA36D.Count() > 0)
                                            {
                                                List<PPLPHExtrasModel> totalizers = extraSub.Where(x => x.FieldName == "Totalizer" 
                                                    && extraWithBlendAA36D.Any(y => x.RowNumber == y.RowNumber && x.LPHID == y.LPHID)).ToList();
                                                double totalizer = 0;
                                                double t = 0;
                                                totalizers.ForEach(x => totalizer += double.TryParse(x.Value, out t) ? t : 0);
                                                if(extras[i].ContainsKey("packingqty")) extras[i]["packingqty"] = totalizer.ToString();
                                                else extras[i].Add("packingqty", totalizer.ToString());
                                            }

                                        }
                                    }
                                }
                                return extras;
                            }
                        },
                    }
                } },
        };

        [HttpPost]
        public ActionResult GetData(string startDate, string endDate, long prodCenterID, string mode, string batch) // Filter submission with header
        {
            try
            {
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
                List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.Find(submissionsFilter).DeserializeToPPLPHSubmissionsList();
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

                Dictionary<string, List<List<string>>> sheets = new Dictionary<string, List<List<string>>>();
                Dictionary<string, List<string>> summarySheets = new Dictionary<string, List<string>>();
                if(SapMap[mode].SheetByComponents != null) SapMap[mode].SheetByComponents.ForEach(sheetByComponent =>
                {
                    List<string> headerComponent = new List<string>();
                    List<Dictionary<string, string>> comvals = new List<Dictionary<string, string>>();
                    submissions.ForEach(sub =>
                    {
                        List<PPLPHComponentsModel> header = componentList.Where(x => x.LPHID == sub.LPHID && sheetByComponent.Displayed.Any(y => y == x.ComponentName)).ToList();
                        List<string> head = header.Select(x => x.ComponentName).Distinct().ToList();
                        if (head.Count() > headerComponent.Count())
                        {
                            headerComponent = head;
                            headerComponent.Insert(0, "LPH ID");
                        }

                        Dictionary<string, string> comval = new Dictionary<string, string>() { { "LPH ID", sub.LPHID.ToString() } };
                        header.ForEach(h =>
                        {
                            string value = valueList.Where(v => h.ID == v.LPHComponentID).Select(x => x.Value).FirstOrDefault();
                            if (value != null)
                            {
                                if (!comval.ContainsKey(h.ComponentName)) comval.Add(h.ComponentName, value);
                                else comval[h.ComponentName] = value;
                            }
                        });
                        comvals.Add(comval);
                    });
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
                    componentRows.Insert(0, headerComponent);
                    sheets.Add(sheetByComponent.SheetName, componentRows);
                });

                if (SapMap[mode].SheetByExtras != null) SapMap[mode].SheetByExtras.ForEach(sheetByExtra =>
                {
                    List<PPLPHExtrasModel> extrasForThisSheet = extraList.Where(x => sheetByExtra.TableNames.Any(y => y == x.HeaderName) && sheetByExtra.Selected.Any(y => y == x.FieldName)).ToList();
                    List<string> joinFieldValue = new List<string>();
                    sheetByExtra.JoinBy.Keys.ToList().ForEach(key => joinFieldValue = joinFieldValue.Concat(extrasForThisSheet.Where(x => x.FieldName == key).Select(x => x.Value)).ToList());
                    List<Dictionary<string, string>> extraResult = new List<Dictionary<string, string>>();
                    List<string> headerRow = sheetByExtra.Selected;
                    joinFieldValue.Distinct().ToList().ForEach(fv =>
                    {
                        List<PPLPHExtrasModel> extraUnexpanded = extrasForThisSheet.Where(x => x.Value == fv && sheetByExtra.JoinBy.Keys.Any(y => x.FieldName == y || sheetByExtra.JoinBy[y].Any(z => z == x.FieldName))).ToList();
                        List<PPLPHExtrasModel> extraExpanded = extrasForThisSheet.Where(x => extraUnexpanded.Any(y => y.HeaderName == x.HeaderName && y.LPHID == x.LPHID && y.RowNumber == x.RowNumber)).OrderByDescending(x => x.Date).ToList();

                        if(sheetByExtra.Accumulated != null) sheetByExtra.Accumulated.ForEach(accumulated =>
                        {
                            List<PPLPHExtrasModel> extraToAccumulated = extraExpanded.Where(x => x.FieldName == accumulated).ToList();
                            if(extraToAccumulated.Count() > 0)
                            {
                                double accumulatedValue = 0;
                                double t = 0;
                                extraToAccumulated.ForEach(x => accumulatedValue += double.TryParse(x.Value, out t) ? t : 0);
                                extraExpanded = extraExpanded.Where(x => !extraToAccumulated.Any(y => x.ID == y.ID)).ToList();
                                PPLPHExtrasModel extraSample = extraToAccumulated.FirstOrDefault();
                                extraSample.Value = accumulatedValue.ToString();
                                extraExpanded.Add(extraSample);
                            }
                        });

                        int tableRowIndex = extraResult.Count();
                        int subTableRowMaxIndex = extraResult.Count();
                        extraResult.Add(new Dictionary<string, string>());
                        long lphId = extraExpanded.FirstOrDefault()?.LPHID ?? 0;
                        PPLPHSubmissionsModel sub = submissions.Where(x => x.LPHID == lphId).FirstOrDefault();
                        extraResult[tableRowIndex].Add("LPHID", lphId.ToString());
                        extraResult[tableRowIndex].Add("Date", sub?.Date.ToString("yyyy-MM-dd") ?? "");
                        extraResult[tableRowIndex].Add("Shift", sub?.Shift ?? "");
                        extraResult[tableRowIndex].Add("LocationID", (sub?.LocationID ?? 0).ToString());
                        extraExpanded.ForEach(extra =>
                        {
                            if (extra.ValueType.Trim() == "Json")
                            {
                                List<Dictionary<string, string>> jsonD =
                                    JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(extra.Value);
                                int hi = headerRow.FindIndex(x => x == extra.FieldName);
                                if (hi >= 0 && jsonD.Count > 0)
                                {
                                    headerRow.RemoveAt(hi);
                                    List<string> keys = jsonD[0].Keys.ToList();
                                    for (int k = keys.Count - 1; k >= 0; k--)
                                        if (keys[k] != "id") headerRow.Insert(hi, keys[k]);
                                }
                                int subTableRowIndex = tableRowIndex;
                                int subTableStartFrom = extraResult[tableRowIndex].Count();
                                foreach (var json in jsonD)
                                {
                                    if (subTableRowMaxIndex >= subTableRowIndex)
                                    {
                                        foreach (KeyValuePair<string, string> item in json.OrderBy(x => x.Key))
                                            if (item.Key != "id" && !extraResult[subTableRowIndex].ContainsKey(item.Key)) 
                                                extraResult[subTableRowIndex].Add(item.Key, item.Value);
                                    }
                                    else
                                    {
                                        Dictionary<string, string> newRow = new Dictionary<string, string>();
                                        foreach (KeyValuePair<string, string> item in json.OrderBy(x => x.Key))
                                            if (item.Key != "id") newRow.Add(item.Key, item.Value);
                                        extraResult.Add(newRow);
                                        subTableRowMaxIndex++;
                                    }
                                    subTableRowIndex++;
                                }
                            }
                            else if (!extraResult[tableRowIndex].ContainsKey(extra.FieldName)) 
                                extraResult[tableRowIndex].Add(extra.FieldName, extra.Value);
                        });
                    });
                    headerRow = sheetByExtra.Displayed;
                    List<List<string>> extraRows = new List<List<string>>();
                    Dictionary<string, double> summaryValues = new Dictionary<string, double>();
                    if (sheetByExtra.BeforeDisplay != null)
                        extraResult = sheetByExtra.BeforeDisplay(extraResult, 
                            _ppLphSubmissionsAppService,
                            _ppLphExtrasAppService);
                    extraResult.ForEach(e =>
                    {
                        List<string> row = new List<string>();
                        headerRow.ForEach(h =>
                        {
                            if (e.ContainsKey(h))
                            {
                                if (double.TryParse(e[h], out double v))
                                {
                                    if(sheetByExtra.Summarized != null 
                                        && sheetByExtra.Summarized.Contains(h))
                                    {
                                        if(summaryValues.ContainsKey(h))
                                            summaryValues[h] += v;
                                        else summaryValues.Add(h, v);
                                    }
                                    row.Add(v.ToString("0.###"));
                                }
                                else row.Add(e[h]);
                            }
                            else row.Add(null);
                        });
                        extraRows.Add(row);
                    });
                    List<string> summaryRows = new List<string>();
                    headerRow.ForEach(h =>
                    {
                        if (summaryValues.ContainsKey(h))
                            summaryRows.Add(summaryValues[h].ToString("0.###"));
                        else summaryRows.Add("0");
                    });
                    extraRows.Insert(0, headerRow);
                    sheets.Add(sheetByExtra.SheetName, extraRows);
                    summarySheets.Add(sheetByExtra.SheetName, summaryRows);
                });

                return Json(new { data = sheets, summaries = summarySheets }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                return Json(new { error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult ExtractExcel(string data, string summaries)
        {
            try
            {
                Dictionary<string, List<List<string>>> sheets = data.DeserializeJson<Dictionary<string, List<List<string>>>>() ?? new Dictionary<string, List<List<string>>>();
                Dictionary<string, List<string>> summariesSheets = summaries.DeserializeJson<Dictionary<string, List<string>>>() ?? new Dictionary<string, List<string>>();
                Session["ExtractDataPPForSAP"] = ExcelGenerator.PPRawDataExtract(sheets, summariesSheets);

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
            if (Session["ExtractDataPPForSAP"] != null)
            {
                byte[] data = Session["ExtractDataPPForSAP"] as byte[];
                Session["ExtractDataPPForSAP"] = null;
                return File(data, "application/octet-stream", "DataSAP.xlsx");
            }
            return new EmptyResult();
        }

    }
}
