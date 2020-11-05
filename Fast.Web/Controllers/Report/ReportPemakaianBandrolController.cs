using Fast.Application.Interfaces;
using Fast.Web.Models.LPH;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;
using Fast.Web.Models;
using Fast.Web.Models.Report;
using Newtonsoft.Json;
using Fast.Infra.CrossCutting.Common;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace Fast.Web.Controllers.Report
{
    public class ReportPemakaianBandrolController : BaseController<LPHModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBrandAppService _brandAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly IWeeksAppService _weeksAppService;
        private readonly ILPHComponentsAppService _lphComponentsAppService;
        private readonly ILPHValuesAppService _lphValuesAppService;
        private readonly ILPHExtrasAppService _lphExtrasAppService;
        private readonly IMachineAppService _machineAppService;

        public ReportPemakaianBandrolController(
          ILPHAppService lphAppService,
          IReferenceAppService referenceAppService,
          IReferenceDetailAppService referenceDetailAppService,
          ILocationAppService locationAppService,
          ILoggerAppService logger,
          IBrandAppService brandAppService,
          ILPHSubmissionsAppService lPHSubmissionsAppService,
          ILPHApprovalsAppService lPHApprovalsAppService,
          IWeeksAppService weeksAppService,
          ILPHComponentsAppService componentsAppService,
          ILPHValuesAppService valuesAppService,
          ILPHExtrasAppService lphExtrasAppService,
          IMachineAppService machineAppService)
        {
            _lphAppService = lphAppService;
            _lphAppService = lphAppService;
            _referenceAppService = referenceAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _locationAppService = locationAppService;
            _logger = logger;
            _brandAppService = brandAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _lphApprovalAppService = lPHApprovalsAppService;
            _weeksAppService = weeksAppService;
            _lphComponentsAppService = componentsAppService;
            _lphValuesAppService = valuesAppService;
            _lphExtrasAppService = lphExtrasAppService;
            _machineAppService = machineAppService;
        }
        public ActionResult Index()
        {
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            string brand = _brandAppService.GetAll();
            List<BrandModel> brandList = brand.DeserializeToBrandList();

            List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
            ViewBag.Brand = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
            return View();
        }

        [HttpPost]
        public ActionResult GetBandrolWeekly()
        {
            try
            {
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
                string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
                //DateTime?[] dateAxis = { };
                double[] bandrolSubmitted = { 0, 0, 0, 0, 0 };
                double[] bandrolApproved = { 0, 0, 0, 0, 0 };

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQueryPMID = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) >= '" + monday.Date.ToShortDateString() +"')";
                List<PMIDPemakaianBandrolModel> ReportListPMID = GetData<List<PMIDPemakaianBandrolModel>>(conString, myQueryPMID) ?? new List<PMIDPemakaianBandrolModel>();
                ReportListPMID = ReportListPMID.OrderBy(x => x.Date).ToList();
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) >= '" + monday.Date.ToShortDateString() + "')";
                List<OtherPemakaianBandrolModel> ReportList = GetData<List<OtherPemakaianBandrolModel>>(conString, myQuery) ?? new List<OtherPemakaianBandrolModel>();
                ReportList = ReportList.OrderBy(x => x.Date).ToList();

                foreach (var satuan in ReportList)
                {
                    if (satuan.Location.Contains("ID-PJ") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[1] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[1] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PK") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[2] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[2] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    } 
                    else if (satuan.Location.Contains("ID-PB") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[4] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[4] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    } 
                }
                foreach (var satuan in ReportListPMID)
                {
                    if (satuan.SubmissionStatus.Trim() == "Approved")
                    {
                        bandrolApproved[3] += satuan.Missing ?? 0;
                        bandrolApproved[0] += satuan.Missing ?? 0;
                    }
                    else
                    {
                        bandrolSubmitted[3] += satuan.Missing ?? 0;
                        bandrolSubmitted[0] += satuan.Missing ?? 0;
                    }
                }
                return Json(new { Header = Header, Submitted = bandrolSubmitted, Approved = bandrolApproved }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetBandrolDaily()
        {
            try
            {
                string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
                double[] bandrolSubmitted = { 0, 0, 0, 0, 0 };
                double[] bandrolApproved = { 0, 0, 0, 0, 0 };

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQueryPMID = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) = '" + DateTime.Now.Date.ToShortDateString() + "')";
                List<PMIDPemakaianBandrolModel> ReportListPMID = GetData<List<PMIDPemakaianBandrolModel>>(conString, myQueryPMID) ?? new List<PMIDPemakaianBandrolModel>();
                ReportListPMID = ReportListPMID.OrderBy(x => x.Date).ToList();
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) = '" + DateTime.Now.Date.ToShortDateString() + "')";
                List<OtherPemakaianBandrolModel> ReportList = GetData<List<OtherPemakaianBandrolModel>>(conString, myQuery) ?? new List<OtherPemakaianBandrolModel>();
                ReportList = ReportList.OrderBy(x => x.Date).ToList();

                foreach (var satuan in ReportList)
                {
                    if (satuan.Location.Contains("ID-PJ") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[1] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[1] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PK") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[2] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[2] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    } 
                    else if (satuan.Location.Contains("ID-PB") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[4] += satuan.BandrolHilang ?? 0;
                            bandrolApproved[0] += satuan.BandrolHilang ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[4] += satuan.BandrolHilang ?? 0;
                            bandrolSubmitted[0] += satuan.BandrolHilang ?? 0;
                        }
                    } 
                }
                foreach (var satuan in ReportListPMID)
                {
                    if (satuan.SubmissionStatus.Trim() == "Approved")
                    {
                        bandrolApproved[3] += satuan.Missing ?? 0;
                        bandrolApproved[0] += satuan.Missing ?? 0;
                    }
                    else
                    {
                        bandrolSubmitted[3] += satuan.Missing ?? 0;
                        bandrolSubmitted[0] += satuan.Missing ?? 0;
                    }
                }
                return Json(new { Header = Header, Submitted = bandrolSubmitted, Approved = bandrolApproved }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetClaimableWeekly()
        {
            try
            {
                var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
                string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
                //DateTime?[] dateAxis = { };
                double[] bandrolSubmitted = { 0, 0, 0, 0, 0 };
                double[] bandrolApproved = { 0, 0, 0, 0, 0 };

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQueryPMID = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) >= '" + monday.Date.ToShortDateString() + "')";
                List<PMIDPemakaianBandrolModel> ReportListPMID = GetData<List<PMIDPemakaianBandrolModel>>(conString, myQueryPMID) ?? new List<PMIDPemakaianBandrolModel>();
                ReportListPMID = ReportListPMID.OrderBy(x => x.Date).ToList();
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) >= '" + monday.Date.ToShortDateString() + "')";
                List<OtherPemakaianBandrolModel> ReportList = GetData<List<OtherPemakaianBandrolModel>>(conString, myQuery) ?? new List<OtherPemakaianBandrolModel>();
                ReportList = ReportList.OrderBy(x => x.Date).ToList();

                foreach (var satuan in ReportList)
                {
                    if (satuan.Location.Contains("ID-PJ") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[1] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[1] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PK") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[2] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[2] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PB") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[4] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[4] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                }
                foreach (var satuan in ReportListPMID)
                {
                    if (satuan.SubmissionStatus.Trim() == "Approved")
                    {
                        bandrolApproved[3] += satuan.ClaimablePack ?? 0;
                        bandrolApproved[0] += satuan.ClaimablePack ?? 0;
                    }
                    else
                    {
                        bandrolSubmitted[3] += satuan.ClaimablePack ?? 0;
                        bandrolSubmitted[0] += satuan.ClaimablePack ?? 0;
                    }
                }
                return Json(new { Header = Header, Submitted = bandrolSubmitted, Approved = bandrolApproved }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetClaimableDaily()
        {
            try
            {
                string[] Header = { "ID", "PJ", "PK", "PI", "PB" };
                double[] bandrolSubmitted = { 0, 0, 0, 0, 0 };
                double[] bandrolApproved = { 0, 0, 0, 0, 0 };

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQueryPMID = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) = '" + DateTime.Now.Date.ToShortDateString() + "')";
                List<PMIDPemakaianBandrolModel> ReportListPMID = GetData<List<PMIDPemakaianBandrolModel>>(conString, myQueryPMID) ?? new List<PMIDPemakaianBandrolModel>();
                ReportListPMID = ReportListPMID.OrderBy(x => x.Date).ToList();
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) = '" + DateTime.Now.Date.ToShortDateString() + "')";
                List<OtherPemakaianBandrolModel> ReportList = GetData<List<OtherPemakaianBandrolModel>>(conString, myQuery) ?? new List<OtherPemakaianBandrolModel>();
                ReportList = ReportList.OrderBy(x => x.Date).ToList();

                foreach (var satuan in ReportList)
                {
                    if (satuan.Location.Contains("ID-PJ") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[1] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[1] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PK") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[2] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[2] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                    else if (satuan.Location.Contains("ID-PB") == true)
                    {
                        if (satuan.SubmissionStatus.Trim() == "Approved")
                        {
                            bandrolApproved[4] += satuan.Claimable ?? 0;
                            bandrolApproved[0] += satuan.Claimable ?? 0;
                        }
                        else
                        {
                            bandrolSubmitted[4] += satuan.Claimable ?? 0;
                            bandrolSubmitted[0] += satuan.Claimable ?? 0;
                        }
                    }
                }
                foreach (var satuan in ReportListPMID)
                {
                    if (satuan.SubmissionStatus.Trim() == "Approved")
                    {
                        bandrolApproved[3] += satuan.ClaimablePack ?? 0;
                        bandrolApproved[0] += satuan.ClaimablePack ?? 0;
                    }
                    else
                    {
                        bandrolSubmitted[3] += satuan.ClaimablePack ?? 0;
                        bandrolSubmitted[0] += satuan.ClaimablePack ?? 0;
                    }
                }
                return Json(new { Header = Header, Submitted = bandrolSubmitted, Approved = bandrolApproved }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetReportPMID(string status, string startDate, string endDate, long prodcenter, string brand, bool fromOther = false) //param nya menyesuaikan
        {
            try
            {
                string[] startDateArray = startDate.Split(new char[] { '-' });
                string[] endDateArray = endDate.Split(new char[] { '-' });

                int endWeek = 0;
                int endYear = 0;
                int startWeek = 0;
                int startYear = 0;
                if (startDateArray[0] != null)
                {
                    if (!Int32.TryParse(startDateArray[0], out startYear))
                    {
                        startYear = 0;
                    }
                }
                if (startDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(startDateArray[1].Substring(1, 2), out startWeek))
                    {
                        startWeek = 0;
                    }
                }
                if (endDateArray[0] != null)
                {
                    if (!Int32.TryParse(endDateArray[0], out endYear))
                    {
                        endYear = 0;
                    }
                }
                if (endDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(endDateArray[1].Substring(1, 2), out endWeek))
                    {
                        endWeek = 0;
                    }
                }

                DateTime startDateReal = FirstDateOfWeekISO8601(startYear, startWeek);
                DateTime endDateReal = FirstDateOfWeekISO8601(endYear, endWeek).AddDays(6);

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) BETWEEN '" + startDateReal.Date.ToShortDateString() + "' AND '" + endDateReal.Date.ToShortDateString() + "')";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<PMIDPemakaianBandrolModel> ReportBaseList = jsondata.DeserializeToPMIDPemakaianBandrolList();

                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                myQuery = @"SELECT * FROM PMIDPemakaianBandrolSlof WHERE (convert(date,[Date]) BETWEEN '" + startDateReal.Date.ToShortDateString() + "' AND '" + endDateReal.Date.ToShortDateString() + "')";

                dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<PMIDPemakaianBandrolSlofModel> ReportBaseListSlof = jsondata.DeserializeToPMIDPemakaianBandrolSlofList();

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodcenter, "productioncenter");
                List<PMIDPemakaianBandrolModel> ReportList = new List<PMIDPemakaianBandrolModel>();
                List<PMIDPemakaianBandrolSlofModel> ReportListSlof = new List<PMIDPemakaianBandrolSlofModel>();
                if (prodcenter == 0)
                {
                    ReportList = ReportBaseList.OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                    ReportListSlof = ReportBaseListSlof.OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }
                else
                {
                    ReportList = ReportBaseList.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                    ReportListSlof = ReportBaseListSlof.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }

                if (status == "Approved")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Approved").ToList();
                    ReportListSlof = ReportListSlof.Where(x => x.SubmissionStatus.Trim() == "Approved").ToList();
                }
                else if (status == "Submitted")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Submitted").ToList();
                    ReportListSlof = ReportListSlof.Where(x => x.SubmissionStatus.Trim() == "Submitted").ToList();
                }

                if(brand != "")
                {
                    ReportList = ReportList.Where(x => x.FACode != null && x.FACode.Trim() == brand).ToList();
                    ReportListSlof = ReportListSlof.Where(x => x.FACode != null && x.FACode.Trim() == brand).ToList();
                }
                
                List<Dictionary<string, string>> laporanHarianBawah = new List<Dictionary<string, string>>();

                foreach (var satuan in ReportList)
                {
                    double OpeningPack = satuan.TSRequest ?? 0;
                    double WIPManual = satuan.ManualTaxStamp ?? 0;
                    double StockAwal = satuan.StockAwal ?? 0;
                    double packOnMachine = satuan.PackOnMachine ?? 0;
                    double PackOnSamplingCabinet = satuan.SamplingCabinet ?? 0;
                    double PackOutofMachine = satuan.PackOutofMachine ?? 0;
                    double WIPfromOtherMachine = satuan.OtherMachine ?? 0;
                    double TotalOpening = satuan.TotalOpening ?? 0;
                    double CasewithTS = satuan.CaseWTS ?? 0;
                    double CasewithOTS = satuan.CaseWoTS ?? 0;
                    double TaxStampReturn = satuan.TaxStampReturn ?? 0;
                    double SisaAkhirBandrol = satuan.SisaAkhirBandrol ?? 0;
                    double packOnMachineClosing = satuan.PackOnMachineClosing ?? 0;
                    double PackOnSamplingCabinetClosing = satuan.PackonSamplingCabinet ?? 0;
                    double PackOutofMachineClosing = satuan.PackOutofMachineClosing ?? 0;
                    double ClaimablePack = satuan.ClaimablePack ?? 0;
                    double IPCCounted = satuan.IPCCounted ?? 0;
                    double IPCOther = satuan.IPCOther ?? 0;
                    double QASamplewStamp = satuan.QASamplewStamp ?? 0;
                    double QASampleWOStamp = satuan.QASamplewoStamp ?? 0;
                    double WIPtoOtherMachine = satuan.WIPtoOtherMachine ?? 0;
                    double TotalClosing = satuan.TotalClosing ?? 0;
                    double TotalLoss = satuan.Missing ?? 0;
                    double PercentLoss = satuan.PercentMissing ?? 0;
                    double Usage = satuan.Usage ?? 0;
                    string Machine = satuan.Machine ?? "";
                    string Shift = satuan.Shift ?? "";
                    string Group = satuan.Group ?? "";
                    string Brand = satuan.FACode ?? "";
                    DateTime Week = satuan.Date ?? DateTime.Now;

                    Dictionary<string, string> detailLaporan = new Dictionary<string, string>();
                    detailLaporan.Add("Week", CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(Week, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString());
                    detailLaporan.Add("LPHID", satuan.LPHID.ToString());
                    detailLaporan.Add("Tanggal", satuan.Date?.ToString("dd-MMM-yy"));
                    detailLaporan.Add("Hari", satuan.Date?.ToString("dddd"));
                    detailLaporan.Add("Shift", Shift.Trim());
                    detailLaporan.Add("Group", Group.Trim());
                    detailLaporan.Add("Brand", Brand.Trim());
                    detailLaporan.Add("Machine", Machine.Trim());
                    detailLaporan.Add("OpeningPack", OpeningPack.ToString());
                    detailLaporan.Add("WIPManual", WIPManual.ToString());
                    detailLaporan.Add("StockAwal", StockAwal.ToString());
                    detailLaporan.Add("packOnMachine", packOnMachine.ToString());
                    detailLaporan.Add("PackOnSamplingCabinet", PackOnSamplingCabinet.ToString());
                    detailLaporan.Add("PackOutofMachine", PackOutofMachine.ToString());
                    detailLaporan.Add("WIPfromOtherMachine", WIPfromOtherMachine.ToString());
                    detailLaporan.Add("TotalOpening", TotalOpening.ToString());
                    detailLaporan.Add("CasewithTS", CasewithTS.ToString());
                    detailLaporan.Add("CasewithOTS", CasewithOTS.ToString());
                    detailLaporan.Add("TaxStampReturn", TaxStampReturn.ToString());
                    detailLaporan.Add("SisaAkhirBandrol", SisaAkhirBandrol.ToString());
                    detailLaporan.Add("packOnMachineClosing", packOnMachineClosing.ToString());
                    detailLaporan.Add("PackOnSamplingCabinetClosing", PackOnSamplingCabinetClosing.ToString());
                    detailLaporan.Add("PackOutofMachineClosing", PackOutofMachineClosing.ToString());
                    detailLaporan.Add("ClaimablePack", ClaimablePack.ToString());
                    detailLaporan.Add("IPCCounted", IPCCounted.ToString());
                    detailLaporan.Add("IPCOther", IPCOther.ToString());
                    detailLaporan.Add("QASamplewStamp", QASamplewStamp.ToString());
                    detailLaporan.Add("QASampleWOStamp", QASampleWOStamp.ToString());
                    detailLaporan.Add("WIPtoOtherMachine", WIPtoOtherMachine.ToString());
                    detailLaporan.Add("TotalClosing", TotalClosing.ToString());
                    detailLaporan.Add("TotalLoss", TotalLoss.ToString());
                    detailLaporan.Add("PercentLoss", PercentLoss.ToString());
                    detailLaporan.Add("Usage", Usage.ToString());
                    detailLaporan.Add("plusminusbandrol", satuan.PlusMinusBandrol);
                    detailLaporan.Add("TSCode", satuan.TaxStampCode);
                    laporanHarianBawah.Add(detailLaporan);
                }
                foreach (var satuan in ReportListSlof)
                {
                    double OpeningPack = satuan.TSRequest ?? 0;
                    double WIPManual = satuan.ManualTaxStamp ?? 0;
                    double StockAwal = satuan.StockAwal ?? 0;
                    double packOnMachine = satuan.PackOnMachine ?? 0;
                    double PackOnSamplingCabinet = satuan.SamplingCabinet ?? 0;
                    double PackOutofMachine = satuan.PackOutofMachine ?? 0;
                    double WIPfromOtherMachine = satuan.OtherMachine ?? 0;
                    double TotalOpening = satuan.TotalOpening ?? 0;
                    double CasewithTS = satuan.CaseWTS ?? 0;
                    double TaxStampReturn = satuan.TaxStampReturn ?? 0;
                    double SisaAkhirBandrol = satuan.SisaAkhirBandrol ?? 0;
                    double packOnMachineClosing = satuan.PackOnMachineClosing ?? 0;
                    double PackOnSamplingCabinetClosing = satuan.PackonSamplingCabinet ?? 0;
                    double PackOutofMachineClosing = satuan.PackOutofMachineClosing ?? 0;
                    double ClaimablePack = satuan.ClaimablePack ?? 0;
                    double IPCCounted = satuan.IPCCounted ?? 0;
                    double IPCOther = satuan.IPCOther ?? 0;
                    double QASamplewStamp = satuan.QASamplewStamp ?? 0;
                    double WIPtoOtherMachine = satuan.WIPtoOtherMachine ?? 0;
                    double TotalClosing = satuan.TotalClosing ?? 0;
                    double TotalLoss = satuan.Missing ?? 0;
                    double PercentLoss = satuan.PercentMissing ?? 0;
                    double Usage = satuan.Usage ?? 0;
                    DateTime Week = satuan.Date ?? DateTime.Now;

                    string Machine = satuan.Machine ?? "";
                    string Shift = satuan.Shift ?? "";
                    string Group = satuan.Group ?? "";
                    string Brand = satuan.FACode ?? "";

                    Dictionary<string, string> detailLaporan = new Dictionary<string, string>();
                    detailLaporan.Add("Week", CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(Week, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString());
                    detailLaporan.Add("LPHID", satuan.LPHID.ToString());
                    detailLaporan.Add("Tanggal", satuan.Date?.ToString("dd-MMM-yy"));
                    detailLaporan.Add("Hari", satuan.Date?.ToString("dddd"));
                    detailLaporan.Add("Shift", Shift.Trim());
                    detailLaporan.Add("Group", Group.Trim());
                    detailLaporan.Add("Brand", Brand.Trim());
                    detailLaporan.Add("Machine", Machine.Trim());
                    detailLaporan.Add("OpeningPack", OpeningPack.ToString());
                    detailLaporan.Add("WIPManual", WIPManual.ToString());
                    detailLaporan.Add("StockAwal", StockAwal.ToString());
                    detailLaporan.Add("packOnMachine", packOnMachine.ToString());
                    detailLaporan.Add("PackOnSamplingCabinet", PackOnSamplingCabinet.ToString());
                    detailLaporan.Add("PackOutofMachine", PackOutofMachine.ToString());
                    detailLaporan.Add("WIPfromOtherMachine", WIPfromOtherMachine.ToString());
                    detailLaporan.Add("TotalOpening", TotalOpening.ToString());
                    detailLaporan.Add("CasewithTS", CasewithTS.ToString());
                    if (CasewithTS != null && CasewithTS != 0)
                    {
                        detailLaporan.Add("CasewithOTS", CasewithTS.ToString());
                    }
                    else
                    {
                        detailLaporan.Add("CasewithOTS", "");
                    }
                    detailLaporan.Add("TaxStampReturn", TaxStampReturn.ToString());
                    detailLaporan.Add("SisaAkhirBandrol", SisaAkhirBandrol.ToString());
                    detailLaporan.Add("packOnMachineClosing", packOnMachineClosing.ToString());
                    detailLaporan.Add("PackOnSamplingCabinetClosing", PackOnSamplingCabinetClosing.ToString());
                    detailLaporan.Add("PackOutofMachineClosing", PackOutofMachineClosing.ToString());
                    detailLaporan.Add("ClaimablePack", ClaimablePack.ToString());
                    detailLaporan.Add("IPCCounted", IPCCounted.ToString());
                    detailLaporan.Add("IPCOther", IPCOther.ToString());
                    detailLaporan.Add("QASamplewStamp", QASamplewStamp.ToString());
                    if(QASamplewStamp != null && QASamplewStamp != 0)
                    {
                        detailLaporan.Add("QASampleWOStamp", QASamplewStamp.ToString());
                    }
                    else
                    {
                        detailLaporan.Add("QASampleWOStamp", "");
                    }
                    detailLaporan.Add("WIPtoOtherMachine", WIPtoOtherMachine.ToString());
                    detailLaporan.Add("TotalClosing", TotalClosing.ToString());
                    detailLaporan.Add("TotalLoss", TotalLoss.ToString());
                    detailLaporan.Add("PercentLoss", PercentLoss.ToString());
                    detailLaporan.Add("Usage", Usage.ToString());
                    detailLaporan.Add("plusminusbandrol", satuan.PlusMinusBandrol);
                    detailLaporan.Add("TSCode", satuan.TaxStampCode);
                    laporanHarianBawah.Add(detailLaporan);
                }


                laporanHarianBawah = laporanHarianBawah.OrderBy(x => x["LPHID"]).ToList();

                return Json(new { Status = "True", Atas = laporanHarianBawah, fromOther = fromOther }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        
        [HttpPost]
        public ActionResult GetReportOther(string status, string startDate, string endDate, long prodcenter, string brand) //param nya menyesuaikan
        {
            try
            {
                string[] startDateArray = startDate.Split(new char[] { '-' });
                string[] endDateArray = endDate.Split(new char[] { '-' });

                int endWeek = 0;
                int endYear = 0;
                int startWeek = 0;
                int startYear = 0;
                if (startDateArray[0] != null)
                {
                    if (!Int32.TryParse(startDateArray[0], out startYear))
                    {
                        startYear = 0;
                    }
                }
                if (startDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(startDateArray[1].Substring(1, 2), out startWeek))
                    {
                        startWeek = 0;
                    }
                }
                if (endDateArray[0] != null)
                {
                    if (!Int32.TryParse(endDateArray[0], out endYear))
                    {
                        endYear = 0;
                    }
                }
                if (endDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(endDateArray[1].Substring(1, 2), out endWeek))
                    {
                        endWeek = 0;
                    }
                }

                DateTime startDateReal = FirstDateOfWeekISO8601(startYear, startWeek);
                DateTime endDateReal = FirstDateOfWeekISO8601(endYear, endWeek).AddDays(6);

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) BETWEEN '" + startDateReal.Date.ToShortDateString() + "' AND '" + endDateReal.Date.ToShortDateString() + "')";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<OtherPemakaianBandrolModel> ReportBaseList = jsondata.DeserializeToOtherPemakaianBandrolList();

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodcenter, "productioncenter");
                List<OtherPemakaianBandrolModel> ReportList = new List<OtherPemakaianBandrolModel>();
                if (prodcenter == 0)
                {
                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM Brands WHERE Code = '" + brand + "'";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }
                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    BrandModel brandModel = jsondata.DeserializeToBrandList().FirstOrDefault();

                    if(brandModel != null && brandModel.PcID == 5)
                    {
                        return GetReportPMID(status, startDate, endDate, prodcenter, brand, true);
                    }
                    ReportList = ReportBaseList.OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }
                else
                {
                    ReportList = ReportBaseList.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }

                if (status == "Approved")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Approved").ToList();
                }
                else if (status == "Submitted")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Submitted").ToList();
                }
                if(brand != "")
                {
                    ReportList = ReportList.Where(x => x.Brand != null && x.Brand.Trim() == brand).ToList();
                }
                
                return Json(new { Status = "True", Atas = ReportList }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetBandrolUsage(string status, string startDate, string endDate, long prodcenter, string brand) //param nya menyesuaikan
        {
            try
            {
                string[] startDateArray = startDate.Split(new char[] { '-' });
                string[] endDateArray = endDate.Split(new char[] { '-' });

                int endWeek = 0;
                int endYear = 0;
                int startWeek = 0;
                int startYear = 0;
                if (startDateArray[0] != null)
                {
                    if (!Int32.TryParse(startDateArray[0], out startYear))
                    {
                        startYear = 0;
                    }
                }
                if (startDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(startDateArray[1].Substring(1, 2), out startWeek))
                    {
                        startWeek = 0;
                    }
                }
                if (endDateArray[0] != null)
                {
                    if (!Int32.TryParse(endDateArray[0], out endYear))
                    {
                        endYear = 0;
                    }
                }
                if (endDateArray[1].Substring(1, 2) != null)
                {
                    if (!Int32.TryParse(endDateArray[1].Substring(1, 2), out endWeek))
                    {
                        endWeek = 0;
                    }
                }

                DateTime startDateReal = FirstDateOfWeekISO8601(startYear, startWeek);
                DateTime endDateReal = FirstDateOfWeekISO8601(endYear, endWeek).AddDays(6);

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                var myQuery = @"SELECT * FROM OtherPemakaianBandrol WHERE (convert(date,[Date]) BETWEEN '" + startDateReal.Date.ToShortDateString() + "' AND '" + endDateReal.Date.ToShortDateString() + "')";

                DataSet dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                string jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<OtherPemakaianBandrolModel> ReportList = jsondata.DeserializeToOtherPemakaianBandrolList();
                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                myQuery = @"SELECT * FROM PMIDPemakaianBandrol WHERE (convert(date,[Date]) BETWEEN '" + startDateReal.Date.ToShortDateString() + "' AND '" + endDateReal.Date.ToShortDateString() + "')";

                dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                List<PMIDPemakaianBandrolModel> ReportListPMID = jsondata.DeserializeToPMIDPemakaianBandrolList();

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodcenter, "productioncenter");
                if (prodcenter == 0)
                {
                    ReportList = ReportList.OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                    ReportListPMID = ReportListPMID.OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }
                else
                {
                    ReportList = ReportList.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                    ReportListPMID = ReportListPMID.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();
                }

                if (status == "Approved")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Approved").ToList();
                    ReportListPMID = ReportListPMID.Where(x => x.SubmissionStatus.Trim() == "Approved").ToList();
                }
                else if (status == "Submitted")
                {
                    ReportList = ReportList.Where(x => x.SubmissionStatus.Trim() == "Submitted").ToList();
                    ReportListPMID = ReportListPMID.Where(x => x.SubmissionStatus.Trim() == "Submitted").ToList();
                }

                if(brand != "")
                {
                    ReportList = ReportList.Where(x => x.Brand != null && x.Brand.Trim() == brand).ToList();
                    ReportListPMID = ReportListPMID.Where(x => x.FACode != null && x.FACode.Trim() == brand).ToList();
                }
                

                List<Dictionary<string, string>> laporanHarianBawah = new List<Dictionary<string, string>>();
                List<string> dateBawah = new List<string>();
                List<string> dayBawah = new List<string>();
                List<string> machineBawah = new List<string>();
                List<string> machineLocationBawah = new List<string>();
                List<string> prodCenter = new List<string>();

                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                myQuery = @"SELECT * FROM [References] WHERE Name = 'LimitasiLostBandrol'";

                dset = new DataSet();
                using (SqlConnection con = new SqlConnection(strConString))
                {
                    SqlCommand cmd = new SqlCommand(myQuery, con);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dset);
                    }
                }

                jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                ReferenceModel referenceModel = jsondata.DeserializeToReferenceList().FirstOrDefault();
                List<ReferenceDetailModel> referenceDetailModels = null;
                if (referenceModel != null)
                {
                    //string referenceDetailString = _referenceDetailAppService.FindBy("ReferenceID", referenceModel.ID);
                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM ReferenceDetails WHERE ReferenceID = " + referenceModel.ID;

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }

                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    referenceDetailModels = jsondata.DeserializeToRefDetailList();
                }

                ReferenceDetailModel targetID = referenceDetailModels.Where(x => x.Code.Trim() == "ID").FirstOrDefault();
                string tr = targetID.Description ?? "0";

                string brandName = "";
                if(brand != "")
                {
                    strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    myQuery = @"SELECT * FROM Brands WHERE Code = '" + brand + "'";

                    dset = new DataSet();
                    using (SqlConnection con = new SqlConnection(strConString))
                    {
                        SqlCommand cmd = new SqlCommand(myQuery, con);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dset);
                        }
                    }
                    jsondata = JsonConvert.SerializeObject(dset.Tables[0]);
                    BrandModel brandModel = jsondata.DeserializeToBrandList().FirstOrDefault();
                    brandName = brandModel.Description;
                }
                

                foreach (var satuan in ReportList)
                {
                    Dictionary<string, string> detailLaporan = new Dictionary<string, string>();
                    string date = satuan.Date?.ToString("dd-MMM-yy");
                    if (dateBawah.Contains(date) == false)
                    {
                        dateBawah.Add(date);
                        dayBawah.Add(satuan.Date?.ToString("dddd"));
                    }
                    string Machine = satuan.Machine ?? "";
                    string Location = satuan.Location ?? "";

                    double bandrol = satuan.BandrolHilang ?? 0;
                    detailLaporan.Add("Tanggal", date);
                    detailLaporan.Add("Shift", satuan.Shift.Trim());
                    detailLaporan.Add("Day", satuan.Date?.ToString("dddd"));
                    detailLaporan.Add("Machine", Machine);
                    detailLaporan.Add("Location", Location.Substring(3, 2));
                    detailLaporan.Add("Pack", bandrol.ToString() ?? "0");
                    ReferenceDetailModel target = referenceDetailModels.Where(x => x.Code.Trim() == satuan.Location.Substring(3, 2)).FirstOrDefault();
                    detailLaporan.Add("Target", target.Description);
                    laporanHarianBawah.Add(detailLaporan);
                }
                foreach (var satuan in ReportListPMID)
                {
                    Dictionary<string, string> detailLaporan = new Dictionary<string, string>();
                    string date = satuan.Date?.ToString("dd-MMM-yy");
                    if (dateBawah.Contains(date) == false)
                    {
                        dateBawah.Add(date);
                        dayBawah.Add(satuan.Date?.ToString("dddd"));
                    }
                    string Machine = satuan.Machine ?? "";
                    string Location = satuan.Location ?? "";
                    double bandrol = satuan.Missing ?? 0;
                    detailLaporan.Add("Tanggal", date);
                    detailLaporan.Add("Shift", satuan.Shift.Trim());
                    detailLaporan.Add("Day", satuan.Date?.ToString("dddd"));
                    detailLaporan.Add("Machine", satuan.Machine);
                    detailLaporan.Add("Location", satuan.Location.Substring(3, 2));
                    detailLaporan.Add("Pack", bandrol.ToString() ?? "0");
                    ReferenceDetailModel target = referenceDetailModels.Where(x => x.Code.Trim() == satuan.Location.Substring(3, 2)).FirstOrDefault();
                    detailLaporan.Add("Target", target.Description);
                    laporanHarianBawah.Add(detailLaporan);
                }

                laporanHarianBawah = laporanHarianBawah.OrderBy(x => x["Machine"]).ToList().OrderBy(x => x["Location"]).ToList();
                for (int i = 0; i < laporanHarianBawah.Count(); i++)
                {
                    if (machineBawah.Contains(laporanHarianBawah[i]["Machine"]) == false)
                    {
                        machineBawah.Add(laporanHarianBawah[i]["Machine"]);
                        machineLocationBawah.Add(laporanHarianBawah[i]["Location"]);
                    }
                }
                return Json(new { Status = "True", Bawah = laporanHarianBawah, dateBawah = dateBawah, dayBawah = dayBawah, machineBawah = machineBawah, machineLocationBawah = machineLocationBawah, startDate = startDateReal, endDate = endDateReal, targetID = tr, Brand = brandName, referenceDetailModels = referenceDetailModels }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        
        public static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // Use first Thursday in January to get first week of the year as
            // it will never be in Week 52/53
            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            // As we're adding days to a date in Week 1,
            // we need to subtract 1 in order to get the right date for week #1
            if (firstWeek == 1)
            {
                weekNum -= 1;
            }

            // Using the first Thursday as starting week ensures that we are starting in the right year
            // then we add number of weeks multiplied with days
            var result = firstThursday.AddDays(weekNum * 7);

            // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
            return result.AddDays(-3);
        }

        [HttpPost]
        public ActionResult GenerateExcel(DateTime dtFilter, long prodCenterID)
        {
            try
            {
                //byte[] excelData = ExcelGenerator.ExportBandrol(AccountName);

                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("content-disposition", "attachment;filename=PemakaianBandrol.xlsx");
                //Response.BinaryWrite(excelData);
                //Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.GenerateFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
        #region Helper

        [HttpPost]
        public int GetCurrentWeekNumber(string date)
        {
            DateTime dt = DateTime.Parse(date);
            var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weeknum;
        }

        [HttpPost]
        public ActionResult GetDepartmentByProdCenterID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, id);

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetBrand(long prodCenterID)
        {
            string brand = _brandAppService.GetAll();
            List<BrandModel> brandList = brand.DeserializeToBrandList();
            if(prodCenterID != 0)
            {
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenterID, "productioncenter");
                brandList = brandList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
            }

            

            List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
            brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
            return Json(new { Status = true, brand = brandModelList }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        private static T GetData<T>(string conString, string query)
        {
            DataSet dset = new DataSet();
            using (SqlConnection con = new SqlConnection(conString))
            {
                SqlCommand cmd = new SqlCommand(query, con);
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dset);
                }
            }

            return JsonConvert.SerializeObject(dset.Tables[0]).DeserializeJson<T>();
        }
    }
}
