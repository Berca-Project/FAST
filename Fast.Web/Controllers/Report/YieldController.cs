using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
	//[CustomAuthorize("YieldController")]
	public class YieldController : BaseController<PPLPHModel>
	{
		private readonly IPPReportYieldsAppService _ppReportYieldsAppService;
		private readonly IPPReportYieldWhitesAppService _ppReportYieldWhitesAppService;
		private readonly IPPReportYieldKreteksAppService _ppReportYieldKreteksAppService;
		private readonly IPPLPHAppService _ppLphAppService;
		private readonly IPPLPHApprovalsAppService _ppLphApprovalAppService;
		private readonly IPPLPHExtrasAppService _ppLphExtrasAppService;
		private readonly IPPLPHSubmissionsAppService _ppLphSubmissionsAppService;
		private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IPPReportYieldOvsAppService _ppReportYieldOvsAppService;
        private readonly IPPReportYieldTargetsAppService _ppReportYieldTargetsAppService;
        private readonly IPPReportYieldIMLsAppService _ppReportYieldIMLsAppService;
        private readonly IPPReportYieldMCDietsAppService _ppReportYieldMCDietsAppService;

        public YieldController(
			IPPReportYieldsAppService ppReportYieldsAppService,
            IPPReportYieldWhitesAppService ppReportYieldWhitesAppService,
            IPPReportYieldKreteksAppService ppReportYieldKreteksAppService,
            IPPLPHAppService ppLPHAppService,
			IPPLPHApprovalsAppService ppLPHApprovalsAppService,
			IPPLPHExtrasAppService ppLPHExtrasAppService,
			IPPLPHSubmissionsAppService ppLPHSubmissionsAppService,
			ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IPPReportYieldOvsAppService ppReportYieldOvsAppService,
            IPPReportYieldTargetsAppService ppReportYieldTargetsAppService,
            IPPReportYieldIMLsAppService ppReportYieldIMLsAppService,
            IPPReportYieldMCDietsAppService ppReportYieldMCDietsAppService)
		{
			_ppReportYieldsAppService = ppReportYieldsAppService;
            _ppReportYieldWhitesAppService = ppReportYieldWhitesAppService;
            _ppReportYieldKreteksAppService = ppReportYieldKreteksAppService;
            _ppLphAppService = ppLPHAppService;
			_ppLphSubmissionsAppService = ppLPHSubmissionsAppService;
			_ppLphApprovalAppService = ppLPHApprovalsAppService;
			_ppLphExtrasAppService = ppLPHExtrasAppService;
			_locationAppService = locationAppService;
            _referenceAppService = referenceAppService;
            _ppReportYieldOvsAppService = ppReportYieldOvsAppService;
            _ppReportYieldTargetsAppService = ppReportYieldTargetsAppService;
            _ppReportYieldIMLsAppService = ppReportYieldIMLsAppService;
            _ppReportYieldMCDietsAppService = ppReportYieldMCDietsAppService;
        }

		// GET: Yield
		public ActionResult Index()
        {
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            return View();
        }

        [HttpPost]
        public ActionResult SetDashboard(string yield)
        {
            try
            {
                string modifiedDate = DateTime.Now.ToString("yyyy-MM-dd");
                string query = "";
                foreach (List<object> yg in yield.DeserializeJson<List<List<object>>>() ?? new List<List<object>>())
                {
                    /*query += @"
UPDATE [dbo].[DashboardYieldGovernances] SET [Target] = '" + yg[5] + @"', [Value] = '" + yg[6] + @"', [ModifiedDate] = '" + modifiedDate + @"' WHERE [Year] = '" + yg[0] + @"' AND [Week] = '" + yg[1] + @"' AND [Location] = '" + yg[2] + @"' AND [Area] = '" + yg[3] + @"' AND [Type] = '" + yg[4] + @"'
IF @@ROWCOUNT=0 INSERT INTO [dbo].[DashboardYieldGovernances] ([Year],[Week],[Location],[Area],[Type],[Target],[Value],[ModifiedDate]) VALUES ('" + yg[0] + "','" + yg[1] + "','" + yg[2] + "','" + yg[3] + "','" + yg[4] + "','" + yg[5] + "','" + yg[6] + "', '" + modifiedDate + "')";
                */
                    query += string.Format("UPDATE [dbo].[DashboardYieldGovernances] SET [Target] = '{0}', [Value] = '{1}', [ModifiedDate] = '{2}' WHERE [Year] = '{3}' AND [Week] = '{4}' AND [Location] = '{5}' AND [Area] = '{6}' AND [Type] = '{7}'; IF @@ROWCOUNT = 0 INSERT INTO[dbo].[DashboardYieldGovernances] ([Year],[Week],[Location],[Area],[Type],[Target],[Value],[ModifiedDate]) VALUES('{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}'); ", yg[5], yg[6], modifiedDate, yg[0], yg[1], yg[2], yg[3], yg[4], yg[0], yg[1], yg[2], yg[3], yg[4], yg[5], yg[6], modifiedDate);
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

        private int GetWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        
        [HttpPost]
		public ActionResult GetData(string startDate, string endDate, long prodCenterID)
		{
            try
            {
                DateTime dtStartDate = DateTime.ParseExact(startDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(endDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                int weekStart = GetWeek(dtStartDate);
                int weekEnd = GetWeek(dtEndDate);

                List<PPReportYieldTargetModel> targetAll = _ppReportYieldTargetsAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("LocationID", AccountDepartmentID.ToString())
                }).DeserializeJson<List<PPReportYieldTargetModel>>() ?? new List<PPReportYieldTargetModel>();
                Dictionary<string, Dictionary<string, string>> TargetList = new Dictionary<string, Dictionary<string, string>>();
                new List<string> { "Clove", "White", "Kretek", "Diet" }.ForEach(f =>
                {
                    Dictionary<string, string> target = new Dictionary<string, string>();
                    targetAll.Where(x => x.Type == f).ToList().ForEach(t => target.Add(t.Name.Replace(" ", ""), t.Target.ToString()));
                    TargetList.Add(f, target);
                });

                List<PPReportYieldModel> yieldCloveAll = _ppReportYieldsAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual),
                    new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeToPPReportYieldList() ?? new List<PPReportYieldModel>();
                List<PPReportYieldModel> yieldCloveList = new List<PPReportYieldModel>();
                List<PPReportYieldKretekModel> yieldKretekAll = _ppReportYieldKreteksAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual),
                    new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeToPPReportYieldKretekList() ?? new List<PPReportYieldKretekModel>();
                List<PPReportYieldKretekModel> yieldKretekBeforeCalculation = new List<PPReportYieldKretekModel>();
                List<PPReportYieldWhiteModel> yieldWhiteAll = _ppReportYieldWhitesAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual),
                    new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeToPPReportYieldWhiteList() ?? new List<PPReportYieldWhiteModel>();
                List<PPReportYieldWhiteModel> yieldWhiteBeforeCalculation = new List<PPReportYieldWhiteModel>();
                List<PPReportYieldOvModel> ovAll = _ppReportYieldOvsAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual),
                    new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeJson<List<PPReportYieldOvModel>>() ?? new List<PPReportYieldOvModel>();
                List<PPReportYieldOvModel> ovList = new List<PPReportYieldOvModel>();
                List<PPReportYieldIMLModel> imlAll = _ppReportYieldIMLsAppService.Find(new List<QueryFilter>
                {
                    new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual),
                    new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeJson<List<PPReportYieldIMLModel>>() ?? new List<PPReportYieldIMLModel>();
                List<PPReportYieldIMLModel> imlList = new List<PPReportYieldIMLModel>();

                for (int i = dtStartDate.Year; i <= dtEndDate.Year; i++)
                {
                    var ywt = yieldWhiteAll.Where(x => x.Year == i).ToList();
                    var yct = yieldCloveAll.Where(x => x.Year == i).ToList();
                    var ykt = yieldKretekAll.Where(x => x.Year == i).ToList();
                    var ot = ovAll.Where(x => x.Year == i).ToList();
                    var imlt = imlAll.Where(x => x.Year == i).ToList();

                    if (i == dtStartDate.Year)
                    {
                        ywt = yieldWhiteAll.Where(x => x.Week >= weekStart).ToList();
                        yct = yieldCloveAll.Where(x => x.Week >= weekStart).ToList();
                        ykt = yieldKretekAll.Where(x => x.Week >= weekStart).ToList();
                        ot = ovAll.Where(x => x.Week >= weekStart).ToList();
                        imlt = imlAll.Where(x => x.Week >= weekStart).ToList();
                    }

                    if (i == dtEndDate.Year)
                    {
                        ywt = yieldWhiteAll.Where(x => x.Week <= weekEnd).ToList();
                        yct = yieldCloveAll.Where(x => x.Week <= weekEnd).ToList();
                        ykt = yieldKretekAll.Where(x => x.Week <= weekEnd).ToList();
                        ot = ovAll.Where(x => x.Week <= weekEnd).ToList();
                        imlt = imlAll.Where(x => x.Week <= weekEnd).ToList();
                    }

                    yieldWhiteBeforeCalculation = yieldWhiteBeforeCalculation.Concat(ywt).ToList();
                    yieldCloveList = yieldCloveList.Concat(yct).ToList();
                    yieldKretekBeforeCalculation = yieldKretekBeforeCalculation.Concat(ykt).ToList();
                    ovList = ovList.Concat(ot).ToList();
                    imlList = imlList.Concat(imlt).ToList();
                }

                List<PPReportYieldWhiteModel> yieldWhiteList = new List<PPReportYieldWhiteModel>();
                yieldWhiteBeforeCalculation.Select(x => x.Year).Distinct().ToList().ForEach(year =>
                {
                    List<PPReportYieldWhiteModel> yieldWhiteThisYear = yieldWhiteBeforeCalculation.Where(x => x.Year == year).ToList();
                    yieldWhiteThisYear.Select(x => x.Week).Distinct().ToList().ForEach(week =>
                    {
                        List<PPReportYieldWhiteModel> yieldWhiteThisWeek = yieldWhiteThisYear.Where(x => x.Week == week).ToList();
                        double cfDry = yieldWhiteThisWeek.Sum(x => x.CFDryExclude) ?? 0;
                        double sumInput = yieldWhiteThisWeek.Sum(x => x.SumInputMaterialDry) ?? 0;
                        double rsAddback = yieldWhiteThisWeek.Sum(x => x.RS_AddbackWet) ?? 0;
                        double infeedMaterial = yieldWhiteThisWeek.Sum(x => x.InfeedMaterialWet) ?? 0;
                        PPReportYieldWhiteModel yield = new PPReportYieldWhiteModel
                        {
                            Year = year,
                            Week = week,
                            DryYield = cfDry == 0 ? 0 : (cfDry / sumInput),
                            WetYield = rsAddback == 0 ? 0 : (rsAddback / infeedMaterial),
                            SumInputMaterialDry = sumInput,
                        };
                        yieldWhiteList.Add(yield);
                    });
                });

                List<PPReportYieldKretekModel> yieldKretekList = new List<PPReportYieldKretekModel>();
                yieldKretekBeforeCalculation.Select(x => x.Year).Distinct().ToList().ForEach(year =>
                {
                    List<PPReportYieldKretekModel> yieldKretekThisYear = yieldKretekBeforeCalculation.Where(x => x.Year == year).ToList();
                    yieldKretekThisYear.Select(x => x.Week).Distinct().ToList().ForEach(week =>
                    {
                        List<PPReportYieldKretekModel> yieldKretekThisWeek = yieldKretekThisYear.Where(x => x.Week == week).ToList();
                        double cFDryExclude = yieldKretekThisWeek.Sum(x => x.CFDryExclude) ?? 0;
                        double infeedMaterialDry = yieldKretekThisWeek.Sum(x => x.InfeedMaterialDry) ?? 0;
                        double bCDry = yieldKretekThisWeek.Sum(x => x.BCWet * x.BCDryMatter / 100) ?? 0;
                        double aCDry = yieldKretekThisWeek.Sum(x => x.ACWet * x.ACDryMatter / 100) ?? 0;

                        double cutFiller = yieldKretekThisWeek.Sum(x => x.CutFiller) ?? 0;
                        double addback = yieldKretekThisWeek.Sum(x => x.Addback) ?? 0;
                        double infeedMaterialWet = yieldKretekThisWeek.Sum(x => x.InfeedMaterialWet) ?? 0;
                        PPReportYieldKretekModel yield = new PPReportYieldKretekModel
                        {
                            Year = year,
                            Week = week,
                            DryYield = cutFiller > 0 ? cFDryExclude / (infeedMaterialDry + bCDry + aCDry) : 0,
                            WetYield = cutFiller > 0 ? (cutFiller - addback) / infeedMaterialWet : 0,
                            InfeedMaterialDry = infeedMaterialDry,
                        };
                        yieldKretekList.Add(yield);
                    });
                });

                List<string> lphClove = new List<string> {
                    "LPHPrimaryCloveInfeedConditioningController",
                    "LPHPrimaryCloveCutDryPackingController"
                };
                List<string> lphWhite = new List<string> {
                    "LPHPrimaryKretekLineFeedingController",
                    "LPHPrimaryKretekLineConditioningController",
                    "LPHPrimaryKretekLineCuttingDryingController",
                    "LPHPrimaryKretekLineAddbackController",
                    "LPHPrimaryKretekLinePackingController"
                };
                List<string> lphKretek = new List<string> {
                    "LPHPrimaryWhiteLineOTPController",
                };
                List<string> lphDiet = new List<string> {
                    "LPHPrimaryDietController",
                };

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenterID, "productioncenter");
                List<QueryFilter> submissionsFilter = new List<QueryFilter>();
                locationIdList.ForEach(loc => submissionsFilter.Add(new QueryFilter("LocationID", loc.ToString(), Operator.Equals, Operation.OrElse)));
                if (locationIdList.Count() % 2 == 1) submissionsFilter.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                submissionsFilter.Add(new QueryFilter("Date", dtStartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                submissionsFilter.Add(new QueryFilter("Date", dtEndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
                List<PPLPHSubmissionsModel> subList = _ppLphSubmissionsAppService.Find(submissionsFilter).DeserializeToPPLPHSubmissionsList()
                    .Where(x => lphClove.Any(y => y == x.LPHHeader) || lphWhite.Any(y => y == x.LPHHeader) || lphKretek.Any(y => y == x.LPHHeader)).ToList();
                subList = subList.Where(s => s == null ? false : _ppLphApprovalAppService.FindBy("LPHSubmissionID", s.ID, false).DeserializeToPPLPHApprovalList().Any(x =>
                {
                    string status = x.Status.Trim().ToLower();
                    return status == "approved" || status == "submitted";
                })).ToList();

                List<QueryFilter> filterExtra = new List<QueryFilter>();
                subList.ForEach(sub => filterExtra.Add(new QueryFilter("LPHID", sub.LPHID.ToString(), Operator.Equals, Operation.OrElse)));
                if (subList.Count() % 2 == 1) filterExtra.Add(new QueryFilter("ID", "0", Operator.Equals, Operation.OrElse));
                filterExtra.Add(new QueryFilter("HeaderName", "waste"));
                List<PPLPHExtrasModel> wasteAll = _ppLphExtrasAppService.Find(filterExtra).DeserializeToPPLPHExtrasList().ToList();

                List<PPReportYieldGovernanceModel> CloveYGList = new List<PPReportYieldGovernanceModel>();
                List<PPReportYieldGovernanceModel> WhiteYGList = new List<PPReportYieldGovernanceModel>();
                List<PPReportYieldGovernanceModel> KretekYGList = new List<PPReportYieldGovernanceModel>();
                List<PPReportYieldGovernanceModel> DietYGList = new List<PPReportYieldGovernanceModel>();
                for (int i = dtStartDate.Year; i <= dtEndDate.Year; i++)
                {
                    int ws = 1;
                    int we = weekEnd;
                    if (i == dtStartDate.Year) ws = weekStart;
                    if (i != dtEndDate.Year) we = GetWeek(new DateTime(i, 12, 31));

                    List<PPReportYieldModel> yieldCloveThisYear = yieldCloveList.Where(x => x.Year == i).ToList();
                    List<PPReportYieldWhiteModel> yieldWhiteThisYear = yieldWhiteList.Where(x => x.Year == i).ToList();
                    List<PPReportYieldKretekModel> yieldKretekThisYear = yieldKretekList.Where(x => x.Year == i).ToList();

                    List<PPReportYieldIMLModel> imlThisYear = imlList.Where(x => x.Year == i).ToList();
                    List<PPLPHSubmissionsModel> subThisYear = subList.Where(x => x.Date.Year == i).ToList();
                    List<PPLPHExtrasModel> wasteThisYear = wasteAll.Where(x => subThisYear.Any(y => y.LPHID == x.LPHID)).ToList();
                    List<PPReportYieldOvModel> ovThisYear = ovList.Where(x => x.Year == i).ToList();

                    for (int j = ws; j <= we; j++)
                    {
                        List<PPReportYieldIMLModel> imlThisWeek = imlThisYear.Where(x => x.Week == j).ToList();
                        List<PPLPHSubmissionsModel> subThisWeek = subThisYear.Where(x => GetWeek(x.Date) == j).ToList();
                        List<PPLPHExtrasModel> wasteThisWeek = wasteThisYear.Where(x => subThisWeek.Any(y => y.LPHID == x.LPHID)).ToList();
                        List<PPReportYieldOvModel> ovThisWeek = ovThisYear.Where(x => x.Week == j).ToList();

                        PPReportYieldModel cloveYield = yieldCloveThisYear.Where(x => x.Week == j).FirstOrDefault();
                        decimal cloveIml = imlThisWeek.Where(x => x.Type == "clove").FirstOrDefault()?.Value ?? 0;
                        List<PPReportYieldOvModel> cloveOv = ovThisWeek.Where(x => x.Type == "clove").ToList();
                        List<PPLPHSubmissionsModel> cloveWasteSub = subThisWeek.Where(x => lphClove.Any(y => y == x.LPHHeader)).ToList();
                        List<PPLPHExtrasModel> cloveWasteAll = wasteThisWeek.Where(x => cloveWasteSub.Any(y => y.LPHID == x.LPHID)).ToList();
                        List<PPLPHExtrasModel> cloveWasteCloveTypeOnly = cloveWasteAll.Where(x => x.FieldName == "CloveOrCsf" && x.Value == "Clove").ToList();
                        List<PPLPHExtrasModel> cloveWasteThisWeek = cloveWasteAll.Where(x => cloveWasteCloveTypeOnly.Any(y => x.LPHID == y.LPHID && x.RowNumber == y.RowNumber)).ToList();
                        CloveYGList.Add(MakeModel(i, j, cloveYield?.DryYield ?? 0, cloveYield?.DryYield ?? 0, cloveYield?.DryInput ?? 0, cloveIml, cloveWasteThisWeek, cloveOv));

                        PPReportYieldWhiteModel whiteYield = yieldWhiteThisYear.Where(x => x.Week == j).FirstOrDefault();
                        decimal whiteIml = imlThisWeek.Where(x => x.Type == "white").FirstOrDefault()?.Value ?? 0;
                        List<PPReportYieldOvModel> whiteOv = ovThisWeek.Where(x => x.Type == "white").ToList();
                        List<PPLPHSubmissionsModel> whiteWasteSub = subThisWeek.Where(x => lphWhite.Any(y => y == x.LPHHeader)).ToList();
                        List<PPLPHExtrasModel> whiteWasteThisWeek = wasteThisWeek.Where(x => whiteWasteSub.Any(y => y.LPHID == x.LPHID)).ToList();
                        WhiteYGList.Add(MakeModel(i, j, whiteYield?.DryYield ?? 0, whiteYield?.DryYield ?? 0, whiteYield?.SumInputMaterialDry ?? 0, whiteIml, whiteWasteThisWeek, whiteOv));

                        PPReportYieldKretekModel kretekYield = yieldKretekThisYear.Where(x => x.Week == j).FirstOrDefault();
                        decimal kretekIml = imlThisWeek.Where(x => x.Type == "kretek").FirstOrDefault()?.Value ?? 0;
                        List<PPReportYieldOvModel> kretekOv = ovThisWeek.Where(x => x.Type == "kretek").ToList();
                        List<PPLPHSubmissionsModel> kretekWasteSub = subThisWeek.Where(x => lphKretek.Any(y => y == x.LPHHeader)).ToList();
                        List<PPLPHExtrasModel> kretekWasteThisWeek = wasteThisWeek.Where(x => kretekWasteSub.Any(y => y.LPHID == x.LPHID)).ToList();
                        KretekYGList.Add(MakeModel(i, j, kretekYield?.DryYield ?? 0, kretekYield?.DryYield ?? 0, kretekYield?.InfeedMaterialDry ?? 0, kretekIml, kretekWasteThisWeek, kretekOv));

                        List<PPLPHSubmissionsModel> dietWasteSub = subThisWeek.Where(x => lphKretek.Any(y => y == x.LPHHeader)).ToList();
                        List<PPLPHExtrasModel> dietWasteThisWeek = wasteThisWeek.Where(x => kretekWasteSub.Any(y => y.LPHID == x.LPHID)).ToList();
                        List<PPLPHExtrasModel> dietOpsNums = dietWasteThisWeek.Where(x => x.FieldName == "WasteNoSKJ").ToList();
                        decimal dietIml = imlThisWeek.Where(x => x.Type == "diet").FirstOrDefault()?.Value ?? 0;
                        PPReportYieldMCDietModel mc = _ppReportYieldMCDietsAppService.Get(new List<QueryFilter>
                        {
                            new QueryFilter("Year", i),
                            new QueryFilter("Week", j),
                            new QueryFilter("LocationID", AccountDepartmentID.ToString()),
                            new QueryFilter("IsDeleted", "0")
                        }).DeserializeJson<PPReportYieldMCDietModel>();
                        DietYGList.Add(MakeDietModel(i, j, dietWasteThisWeek, dietOpsNums, dietIml, mc));
                    }
                }

                return Json(new { CloveYGList, WhiteYGList, KretekYGList, DietYGList, TargetList, ovList });
            }
            catch (Exception e)
            {
                return Json(new { Error = e.Message });
            }
		}

        private PPReportYieldGovernanceModel MakeDietModel(int i, int j, List<PPLPHExtrasModel> dietWasteThisWeek, List<PPLPHExtrasModel> dietOpsNums, decimal dietIml, PPReportYieldMCDietModel mc)
        {
            double mcPacking = mc?.MCPacking ?? 0;
            double mcFlake = mc?.Flake ?? 0;
            double mcKrosok = mc?.MCKrosok ?? 0;
            double mcDM = mc?.DM ?? 0;
            double mcCVIB0069 = mc?.CVIB0069 ?? 0;
            double mcCSFR0022 = mc?.CSFR0022 ?? 0;
            double mcDSCL0034 = mc?.DSCL0034 ?? 0;
            double mcCVIB0070 = mc?.CVIB0070 ?? 0;
            double mcRV0054 = mc?.RV0054 ?? 0;
            PPReportYieldGovernanceModel dietYg = new PPReportYieldGovernanceModel
            {
                Year = i,
                Week = j,
                Items = new Dictionary<string, double>()
                {
                    { "DryYield", 0 },
                    { "WetYield", 0 },
                    { "DryWaste", 0 },
                    { "HeaviestWaste", 0 },
                    { "DustWaste", 0 },
                    { "HotDustWaste", 0 },
                    { "InfeedMassLoss", decimal.ToDouble(dietIml) / 100 },
                },
            };
            double sumWetInput = 0;
            double sumDryInput = 0;
            double sumOutput = 0;
            double sumCasing = 0;
            foreach (PPLPHExtrasModel opsExtra in dietOpsNums)
            {
                List<PPLPHExtrasModel> opsExpanded = dietWasteThisWeek.Where(x => x.LPHID == opsExtra.LPHID && x.RowNumber == opsExtra.RowNumber).ToList();
                string blend = opsExpanded.Where(x => x.FieldName == "WasteBlend").FirstOrDefault()?.Value ?? "";
                string opsString = opsExpanded.Where(x => x.FieldName == "WasteNoSKJ").FirstOrDefault()?.Value ?? "";
                int ops = int.TryParse(opsString, out int o) ? o : 0;
                string ConditioningOPS = "1" + opsString.Substring(2);

                PPLPHExtrasModel casingThisOps = _ppLphExtrasAppService.Get(new List<QueryFilter>
                {
                    new QueryFilter("HeaderName", "krosok"),
                    new QueryFilter("FieldName", "krosokBatchNo"),
                    new QueryFilter("Value", ConditioningOPS, Operator.Contains),
                    new QueryFilter("IsDeleted", "0")
                }).DeserializeToPPLPHExtras();
                int casing = 0;
                if (casingThisOps != null)
                {
                    PPLPHExtrasModel casingVal = _ppLphExtrasAppService.Get(new List<QueryFilter>
                    {
                        new QueryFilter("HeaderName", "krosok"),
                        new QueryFilter("LPHID", casingThisOps.LPHID.ToString()),
                        new QueryFilter("RowNumber", casingThisOps.RowNumber.ToString()),
                        new QueryFilter("IsDeleted", "0"),
                        new QueryFilter("FieldName", "krosokTotalizer")
                    }).DeserializeToPPLPHExtras();
                    casing = int.TryParse(casingVal?.Value ?? "", out int c) ? c : 0;
                }
                sumCasing += casing;

                double CVIB0069 = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteCVIB0069").FirstOrDefault()?.Value ?? "", out double cvib0069) ? cvib0069 : 0;
                double CSFR0022 = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteCSFR0022").FirstOrDefault()?.Value ?? "", out double csfr0022) ? csfr0022 : 0;
                double DSCL0034 = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteCSFR0022").FirstOrDefault()?.Value ?? "", out double dscl0034) ? dscl0034 : 0;
                double CVIB0070 = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteCSFR0022").FirstOrDefault()?.Value ?? "", out double cvib0070) ? cvib0070 : 0;
                double RV0054 = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteCSFR0022").FirstOrDefault()?.Value ?? "", out double rv0054) ? rv0054 : 0;
                double Output = double.TryParse(opsExpanded.Where(x => x.FieldName == "WastePackingKg").FirstOrDefault()?.Value ?? "", out double output) ? output : 0;
                double RM = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteRM").FirstOrDefault()?.Value ?? "", out double rm) ? rm : 0;
                double Addback = double.TryParse(opsExpanded.Where(x => x.FieldName == "WasteAddback").FirstOrDefault()?.Value ?? "", out double addback) ? addback : 0;
                double Input = RM + Addback;

                if (mc != null)
                {
                    double dryInput = ((Input - mcFlake) * (1 - mcKrosok)) + (mcFlake * (1 - mcFlake));
                    double dryCasing = casing * mcDM;
                    sumWetInput = Input;
                    sumDryInput += dryInput;
                    sumOutput += Output;
                    dietYg.Items["DryWaste"] += CVIB0069 * (1 - mcCVIB0069);
                    dietYg.Items["HeaviestWaste"] += CSFR0022 * (1 - mcCSFR0022);
                    dietYg.Items["DustWaste"] += ((DSCL0034 * (1 - mcDSCL0034)) + (CVIB0070 * (1 - mcCVIB0070)));
                    dietYg.Items["HotDustWaste"] += RV0054 * (1 - mcRV0054);
                }
            }
            double dryInputCasing = sumDryInput + sumCasing;
            dryInputCasing = dryInputCasing == 0 ? 1 : dryInputCasing;
            dietYg.Items["DryYield"] = sumOutput * (1 - mcPacking) / dryInputCasing;
            dietYg.Items["WetYield"] = sumOutput / (sumWetInput == 0 ? 1 : sumWetInput);
            if(dryInputCasing == 0)
            {
                dietYg.Items["DryWaste"] = 0;
                dietYg.Items["HeaviestWaste"] = 0;
                dietYg.Items["DustWaste"] = 0;
                dietYg.Items["HotDustWaste"] = 0;
            }
            else
            {
                dietYg.Items["DryWaste"] = dietYg.Items["DryWaste"] / dryInputCasing;
                dietYg.Items["HeaviestWaste"] = dietYg.Items["HeaviestWaste"] / dryInputCasing;
                dietYg.Items["DustWaste"] = dietYg.Items["DustWaste"] / dryInputCasing;
                dietYg.Items["HotDustWaste"] = dietYg.Items["HotDustWaste"] / dryInputCasing;
            }
            dietYg.Items.Add("UnaccountablePrimary", 1 - (
                dietYg.Items["DryYield"] + 
                dietYg.Items["WetYield"] + 
                dietYg.Items["DryWaste"] + 
                dietYg.Items["HeaviestWaste"] + 
                dietYg.Items["DustWaste"] + 
                dietYg.Items["HotDustWaste"] + 
                dietYg.Items["InfeedMassLoss"]));
            return dietYg;
        }

        private PPReportYieldGovernanceModel MakeModel(int year, int week, double dry, double wet, double materialDry, decimal iml, List<PPLPHExtrasModel> wastes, List<PPReportYieldOvModel> ovs)
        {
            PPReportYieldGovernanceModel yg = new PPReportYieldGovernanceModel
            {
                Year = year,
                Week = week,
                Items = new Dictionary<string, double>()
                {
                    { "DryYield", dry },
                    { "WetYield", wet },
                    { "DryWaste", 0 },
                    { "WetWaste", 0 },
                    { "DustWaste", 0 },
                    { "InfeedMassLoss", decimal.ToDouble(iml) / 100 },
                },
            };
            if (materialDry != 0)
            {
                List<string> categories = wastes.Where(x => x.FieldName == "WasteType").Select(x => x.Value).Distinct().ToList();
                foreach (string category in categories)
                {
                    List<PPLPHExtrasModel> types = wastes.Where(x => x.FieldName == "WasteType" && x.Value.Contains(category)).ToList();
                    foreach (PPLPHExtrasModel type in types)
                    {
                        List<PPLPHExtrasModel> wasteThisType = wastes.Where(y => y.LPHID == type.LPHID && y.RowNumber == type.RowNumber).ToList();
                        string waste = wasteThisType.Where(y => y.FieldName == "Waste").FirstOrDefault()?.Value ?? "";
                        string area = wasteThisType.Where(y => y.FieldName == "Area").FirstOrDefault()?.Value ?? "";
                        double weight = wasteThisType.Where(y => y.FieldName == "Weight").Sum(x => double.TryParse(x.Value, out double y) ? y : 0);
                        if (!string.IsNullOrEmpty(waste) && !string.IsNullOrEmpty(area))
                        {
                            PPReportYieldOvModel ovObj = ovs.Where(y => y.Waste == waste && y.Area == area).FirstOrDefault();
                            if (ovObj != null)
                            {
                                double ov = double.TryParse(ovObj.OvValue.ToString(), out double o) ? o : 1;
                                weight = (100 - ov) / 100 * weight;
                            }
                            if (category.Contains("Dry")) yg.Items["DryWaste"] += weight;
                            else if (category.Contains("Wet")) yg.Items["WetWaste"] += weight;
                            else if (category.Contains("Dust")) yg.Items["DustWaste"] += weight;
                        }
                    }
                }
                yg.Items["DryWaste"] /= materialDry;
                yg.Items["WetWaste"] /= materialDry;
                yg.Items["DustWaste"] /= materialDry;
            }
            yg.Items.Add("UnaccountablePrimary", 1 - (
                yg.Items["DryYield"] + 
                yg.Items["WetYield"] + 
                yg.Items["DryWaste"] + 
                yg.Items["WetWaste"] + 
                yg.Items["DustWaste"] + 
                yg.Items["InfeedMassLoss"]));
            return yg;
        }

        [HttpPost]
        public ActionResult ExtractExcel(string data)
        {
            try
            {
                List<List<string>> sheets = string.IsNullOrEmpty(data) ?
                    new List<List<string>>() :
                    JsonConvert.DeserializeObject<List<List<string>>>(data);
                Session["ExtractYieldGovernance"] = ExcelGenerator.PPRawDataExtract(new Dictionary<string, List<List<string>>>
                {
                    {"Sheet1", sheets }
                });

                return Json(new { status = true }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                return Json(new { status = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult DownloadExcel()
        {
            if (Session["ExtractYieldGovernance"] != null)
            {
                byte[] data = Session["ExtractYieldGovernance"] as byte[];
                Session["ExtractYieldGovernance"] = null;
                return File(data, "application/octet-stream", "YieldGovernance.xlsx");
            }
            return new EmptyResult();
        }


        public ActionResult DownloadTemplate(string filename)
		{
			try
			{
				string filepath = Server.MapPath("..") + "\\Templates\\" + filename + ".xlsx";

				if (System.IO.File.Exists(filepath)) using (FileStream fs = System.IO.File.OpenRead(filepath))
					{
						byte[] data = new byte[fs.Length];
						int br = fs.Read(data, 0, data.Length);
						if (br != fs.Length) throw new IOException(filepath);

						return File(data, MediaTypeNames.Application.Octet, filename + ".xlsx");
					}
			}
			catch (Exception)
			{
			}
			return RedirectToAction("Index");
		}

        [HttpPost]
        public ActionResult Upload(PPReportYieldGovernantUploadModel model)
        {
            IExcelDataReader reader = null;
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }
                if (model.PostedFilename == null || model.PostedFilename.ContentLength <= 0)
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
                }

                List<PPReportYieldOvModel> ovs = new List<PPReportYieldOvModel>();
                using (Stream stream = model.PostedFilename.InputStream)
                {
                    string filename = model.PostedFilename.FileName.ToLower();
                    if (filename.EndsWith(".xls")) reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    else if (filename.EndsWith(".xlsx")) reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    else
                    {
                        SetFalseTempData(UIResources.FileNotSupported);
                        return RedirectToAction("Index");
                    }
                    using (DataTable dt = reader.AsDataSet().Tables[0])
                    {
                        for (int i = 1; i < dt.Rows.Count; i++)
                        {
                            int col = 0;
                            ovs.Add(new PPReportYieldOvModel
                            {
                                Type = model.Type,
                                Area = dt.Rows[i][col++].ToString(),
                                Waste = dt.Rows[i][col++].ToString(),
                                WasteType = dt.Rows[i][col++].ToString(),
                                OvValue = decimal.TryParse(dt.Rows[i][col++].ToString(), out decimal v) ? v : default
                            });

                        }
                    }
                }

                string[] splitWeekStart = model.WeekStart.Split('-');
                string[] splitWeekEnd = model.WeekEnd.Split('-');
                int yearEnd = int.TryParse(splitWeekEnd[0], out int ye) ? ye : 0;
                int weekEnd = int.TryParse(splitWeekEnd[1].Substring(1), out int we) ? we : 0;
                List<PPReportYieldOvModel> ovToAdd = new List<PPReportYieldOvModel>();

                for (int year = int.TryParse(splitWeekStart[0], out int ys) ? ys : 0; year <= yearEnd; year++)
                {
                    for (int week = int.TryParse(splitWeekStart[1].Substring(1), out int ws) ? ws : 0; week <= weekEnd; week++)
                    {
                        ICollection<QueryFilter> filters = new List<QueryFilter>
                        {
                            new QueryFilter("LocationID", AccountDepartmentID.ToString()),
                            new QueryFilter("Week", week),
                            new QueryFilter("Year", year),
                            new QueryFilter("Type", model.Type),
                            new QueryFilter("IsDeleted", "0")
                        };

                        List<PPReportYieldOvModel> vs = _ppReportYieldOvsAppService.FindNoTracking(filters).DeserializeJson<List<PPReportYieldOvModel>>();
                        foreach (PPReportYieldOvModel ov in ovs)
                        {
                            PPReportYieldOvModel v = vs?.Where(x => x.Area == ov.Area && x.Waste == ov.Waste && x.WasteType == ov.WasteType).FirstOrDefault();
                            if (v != null)
                            {
                                v.OvValue = ov.OvValue;
                                v.ModifiedBy = AccountName;
                                v.ModifiedDate = DateTime.Now;
                                _ppReportYieldOvsAppService.Update(JsonHelper<PPReportYieldOvModel>.Serialize(v) ?? "");
                            }
                            else ovToAdd.Add(new PPReportYieldOvModel
                            {
                                Type = ov.Type,
                                Area = ov.Area,
                                Waste = ov.Waste,
                                WasteType = ov.WasteType,
                                OvValue = ov.OvValue,
                                Week = week,
                                Year = year,
                                LocationID = AccountDepartmentID,
                                ModifiedBy = AccountName,
                                ModifiedDate = DateTime.Now,
                            });
                        }
                    }
                }
                if (ovToAdd.Count() > 0)
                    _ppReportYieldOvsAppService.AddRange(JsonHelper<List<PPReportYieldOvModel>>.Serialize(ovToAdd));

                SetTrueTempData(UIResources.UploadSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UploadFailed + " - " + ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }
            }

            return RedirectToAction("Index");

        }

        [HttpPost]
        public ActionResult SetTarget(HttpPostedFileBase postedFilename)
        {
            IExcelDataReader reader = null;
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }
                if (postedFilename == null || postedFilename.ContentLength <= 0)
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
                }

                List<PPReportYieldTargetModel> targets = new List<PPReportYieldTargetModel>();
                using (Stream stream = postedFilename.InputStream)
                {
                    string filename = postedFilename.FileName.ToLower();
                    if (filename.EndsWith(".xls")) reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    else if (filename.EndsWith(".xlsx")) reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    else
                    {
                        SetFalseTempData(UIResources.FileNotSupported);
                        return RedirectToAction("Index");
                    }
                    using (DataTable dt = reader.AsDataSet().Tables[0])
                    {
                        for (int i = 1; i < dt.Rows.Count; i++)
                        {
                            int col = 0;
                            targets.Add(new PPReportYieldTargetModel
                            {
                                Type = dt.Rows[i][col++].ToString(),
                                Name = dt.Rows[i][col++].ToString(),
                                Target = decimal.TryParse(dt.Rows[i][col++].ToString(), out decimal v) ? v : default
                            });
                        }
                    }
                }

                List<PPReportYieldTargetModel> targetToAdd = new List<PPReportYieldTargetModel>();
                foreach(PPReportYieldTargetModel target in targets)
                {
                    ICollection<QueryFilter> filters = new List<QueryFilter>
                    {
                        new QueryFilter("LocationID", AccountDepartmentID.ToString()),
                        new QueryFilter("Type", target.Type),
                        new QueryFilter("Name", target.Name),
                        new QueryFilter("IsDeleted", "0")
                    };

                    PPReportYieldTargetModel ts = _ppReportYieldTargetsAppService.Get(filters, true).DeserializeJson<PPReportYieldTargetModel>();
                    if (ts != null)
                    {
                        ts.Target = target.Target;
                        ts.ModifiedBy = AccountName;
                        ts.ModifiedDate = DateTime.Now;
                        _ppReportYieldTargetsAppService.Update(JsonHelper<PPReportYieldTargetModel>.Serialize(ts) ?? "");
                    }
                    else targetToAdd.Add(new PPReportYieldTargetModel
                    {
                        Type = target.Type,
                        Name = target.Name,
                        LocationID = AccountDepartmentID,
                        ModifiedBy = AccountName,
                        ModifiedDate = DateTime.Now,
                    });
                };
                if (targetToAdd.Count() > 0)
                    _ppReportYieldTargetsAppService.AddRange(JsonHelper<List<PPReportYieldTargetModel>>.Serialize(targetToAdd));

                SetTrueTempData(UIResources.UploadSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.UploadFailed + " - " + ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }
            }
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult InputMass(string type, string value, string weekStart, string weekEnd)
        {
            try
            {
                string[] splitWeekStart = weekStart.Split('-');
                string[] splitWeekEnd = weekEnd.Split('-');
                int yEnd = int.TryParse(splitWeekEnd[0], out int ye) ? ye : 0;
                int wEnd = int.TryParse(splitWeekEnd[1].Substring(1), out int we) ? we : 0;
                decimal imlInt = decimal.TryParse(value, out decimal t) ? t : 0;
                List<PPReportYieldIMLModel> imlToAdd = new List<PPReportYieldIMLModel>();

                for (int year = int.TryParse(splitWeekStart[0], out int ys) ? ys : 0; year <= yEnd; year++)
                {
                    for (int week = int.TryParse(splitWeekStart[1].Substring(1), out int ws) ? ws : 0; week <= wEnd; week++)
                    {
                        ICollection<QueryFilter> filters = new List<QueryFilter>
                        {
                            new QueryFilter("LocationID", AccountDepartmentID.ToString()),
                            new QueryFilter("Week", week),
                            new QueryFilter("Year", year),
                            new QueryFilter("Type", type),
                            new QueryFilter("IsDeleted", "0")
                        };

                        PPReportYieldIMLModel imlInDB = _ppReportYieldIMLsAppService.Get(filters, true).DeserializeJson<PPReportYieldIMLModel>();
                        if (imlInDB != null)
                        {
                            imlInDB.Value = imlInt;
                            imlInDB.ModifiedBy = AccountName;
                            imlInDB.ModifiedDate = DateTime.Now;
                            _ppReportYieldIMLsAppService.Update(JsonHelper<PPReportYieldIMLModel>.Serialize(imlInDB) ?? "");
                        }
                        else imlToAdd.Add(new PPReportYieldIMLModel
                        {
                            Week = week,
                            Year = year,
                            LocationID = AccountDepartmentID,
                            ModifiedBy = AccountName,
                            ModifiedDate = DateTime.Now,
                            Value = imlInt,
                            Type = type,
                        });
                    }
                }
                if (imlToAdd.Count > 0)
                    _ppReportYieldIMLsAppService.AddRange(JsonHelper<List<PPReportYieldIMLModel>>.Serialize(imlToAdd));

                SetTrueTempData("Input Success");
            }
            catch (Exception ex)
            {
                SetFalseTempData("Error : " + ex.Message);
            }

            return RedirectToAction("Index");
        }

        // GET: Yield/Details/5
        public ActionResult Details(int id)
		{
			return View();
		}

		// GET: Yield/Create
		public ActionResult Create()
		{
			return View();
		}

		// POST: Yield/Create
		[HttpPost]
		public ActionResult Create(FormCollection collection)
		{
			try
			{
				// TODO: Add insert logic here

				return RedirectToAction("Index");
			}
			catch
			{
				return View();
			}
		}

		// GET: Yield/Edit/5
		public ActionResult Edit(int id)
		{
			return View();
		}

		// POST: Yield/Edit/5
		[HttpPost]
		public ActionResult Edit(int id, FormCollection collection)
		{
			try
			{
				// TODO: Add update logic here

				return RedirectToAction("Index");
			}
			catch
			{
				return View();
			}
		}

		// GET: Yield/Delete/5
		public ActionResult Delete(int id)
		{
			return View();
		}

		// POST: Yield/Delete/5
		[HttpPost]
		public ActionResult Delete(int id, FormCollection collection)
		{
			try
			{
				// TODO: Add delete logic here

				return RedirectToAction("Index");
			}
			catch
			{
				return View();
			}
		}
	}
}
