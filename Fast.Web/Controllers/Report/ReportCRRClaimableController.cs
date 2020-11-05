using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Resources;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Mvc;
using Fast.Web.Models;
using Newtonsoft.Json;
using Fast.Infra.CrossCutting.Common;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace Fast.Web.Controllers.Report
{
    public class ReportCRRClaimableController : BaseController<LPHModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBrandAppService _brandAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly ILPHComponentsAppService _lphComponentsAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly ILPHValuesAppService _lphValuesAppService;
        private readonly ILPHExtrasAppService _lphExtrasAppService;

        public ReportCRRClaimableController(
          ILPHAppService lphAppService,
          IReferenceAppService referenceAppService,
          ILocationAppService locationAppService,
          ILoggerAppService logger,
          IBrandAppService brandAppService,
          ILPHSubmissionsAppService lPHSubmissionsAppService,
          ILPHApprovalsAppService lPHApprovalsAppService,
          IReferenceDetailAppService referenceDetailAppService,
          ILPHComponentsAppService lphComponentsAppService,
          IMachineAppService machineAppService,
          ILPHValuesAppService valuesAppService,
          ILPHExtrasAppService lphExtrasAppService)
        {
            _lphAppService = lphAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _logger = logger;
            _brandAppService = brandAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _lphApprovalAppService = lPHApprovalsAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _lphComponentsAppService = lphComponentsAppService;
            _machineAppService = machineAppService;
            _lphExtrasAppService = lphExtrasAppService;
            _lphValuesAppService = valuesAppService;
        }
        public ActionResult Index()
        {
            GetTempData();
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
            ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
            ViewBag.CRRType = null;
            if (referenceModelCRRType != null)
            {
                string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
                ViewBag.CRRType = referenceDetailStringCRRType.DeserializeToRefDetailList().OrderBy(x => x.Code);
            }
            string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
            ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
            ViewBag.RIOutput = null;
            if (referenceModelRIOutput != null)
            {
                string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
                ViewBag.RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList().OrderBy(x => x.Code);
            }
            return View();
        }

        public class Crr
        {
            public DateTime Date;
            public string shift;
            public Dictionary<string, List<CrrMesin>> Machines;
        }
        public class CrrMesin
        {
            public Dictionary<string, List<CrrBrand>> Machine;

            //public string Brand;
        }
        public class CrrBrand
        {
            public double Output;
            public double CigReject;
            public double TobReject;
            public double Total;
        }

        [HttpPost]
        public ActionResult GetReportWithParam(string startDate, string endDate, long prodCenter, string machine, string status)
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

                string dayStart = dtStart.ToString("yyyy-MM-dd");
                string dayEnd = dtEnd.ToString("yyyy-MM-dd");

                // Chanif: rubah ke ADO saja, lebih cepat. ini serius kamu ambil semua data LPH yang header = Packer? bakal banyak banget lo....
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                //kalau startdate & enddate nya sudah pasti bener, mending pakai query gini:
                var myQuery = @"SELECT * FROM LPHSubmissions WHERE ( LPHHeader = 'PackerController' OR LPHHeader = 'GWGeneralController' OR LPHHeader = 'RipperController') AND (convert(Date,[Date]) BETWEEN  '" + dayStart + "' AND '" + dayEnd + "' AND IsDeleted = 0)";

                //2detik
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
                List<LPHSubmissionsModel> lphListBase = jsondata.DeserializeToLPHSubmissionsList();
                
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenter, "productioncenter");
                lphListBase = lphListBase.Where(x => locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.Shift).OrderBy(x => x.Date).ToList();

                //List<QueryFilter> subFilter = new List<QueryFilter>();
                //subFilter.Add(new QueryFilter("Date", dayEnd, Operator.LessThanOrEqualTo));
                //subFilter.Add(new QueryFilter("Date", dayStart, Operator.GreaterThanOrEqual));

                //List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenter, "productioncenter");
                //List<LPHSubmissionsModel> submissions = _lphSubmissionsAppService.Find(subFilter).DeserializeToLPHSubmissionsList()
                //    .Where(x => locationIdList.Any(y => y == x.LocationID)).OrderByDescending(x => x.Date).ToList();
                //List<LPHSubmissionsModel> submissions = _lphSubmissionsAppService.GetAll(true).DeserializeToLPHSubmissionsList()
                //    .Where(x => x.Date >= dtStart && x.Date <= dtEnd && locationIdList.Any(y => y == x.LocationID)).OrderByDescending(x => x.Date).ToList();
                //submissions = submissions.Where(s => s == null ? false : _lphApprovalAppService.FindBy("LPHSubmissionID", s.ID, false).DeserializeToPPLPHApprovalList().Any(x =>
                //{
                //    string statusSub = x.Status.Trim().ToLower();
                //    if (status == "Submitted-Approved")
                //    {
                //        return statusSub == "approved" || statusSub == "submitted";
                //    }
                //    return statusSub == status.ToLower();
                //})).ToList();
                if (lphListBase.Count == 0)
                {
                    return Json(new { Status = "False", data = "Data Not Found" }, JsonRequestBehavior.AllowGet);
                }
                var submissionsID = lphListBase.Select(x => x.ID).ToList();
                var LPHsID = lphListBase.Select(x => x.LPHID).ToList();
                
                strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                myQuery = @"SELECT * FROM LPHApprovals WHERE IsDeleted = 0 AND LPHSubmissionID IN (" + string.Join(",", submissionsID) + ")";

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
                List<LPHApprovalsModel> approval = jsondata.DeserializeToLPHApprovalList();
                lphListBase = lphListBase.Where(l =>
                {
                    //6 detik
                    LPHSubmissionsModel s = lphListBase.Where(x => x.LPHID == l.LPHID).ToList().FirstOrDefault();
                    if (s != null) return approval.Where(x =>
                    {
                        string statusA = x.Status.Trim().ToLower();
                        if (status == "Submitted-Approved")
                        {
                            return x.LPHSubmissionID == s.ID && (statusA == "approved" || statusA == "submitted");
                        }
                        else if (status == "Submitted")
                        {
                            return x.LPHSubmissionID == s.ID && (statusA == "submitted");
                        }
                        else
                        {
                            return x.LPHSubmissionID == s.ID && (statusA == "approved");
                        }
                    }).Count() > 0;
                    return false;
                }).ToList();

                List<Crr> crrList = new List<Crr>();
                //tanggal, shift, mesin, FA
                Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>> masterDataPacker = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>>();
                Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>> masterDataMaker = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>>();
                //tgl, shift, fa
                Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> masterDataCrr = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> masterDataOutput = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                Dictionary<string, List<ReferenceDetailModel>> headerGI = new Dictionary<string, List<ReferenceDetailModel>>();

                //get item dari reference
                string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
                ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
                string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
                List<ReferenceDetailModel> itemCrr = referenceDetailStringCRRType.DeserializeToRefDetailList();
                itemCrr = itemCrr.OrderBy(x => x.Code).ToList();

                //get type output dari reference
                string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
                ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
                string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
                List<ReferenceDetailModel> RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList();
                RIOutput = RIOutput.OrderBy(x => x.Code).ToList();

                for (DateTime processedDate = dtStart; processedDate <= dtEnd; processedDate = processedDate.AddDays(1))
                {
                    string DateProcessedDate = processedDate.ToString("dd-MMM-yy");
                    if (masterDataPacker.ContainsKey(DateProcessedDate) == false)
                    {
                        Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> masterTanggal = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                        masterDataPacker.Add(DateProcessedDate, masterTanggal);
                    }
                    if (masterDataMaker.ContainsKey(DateProcessedDate) == false)
                    {
                        Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> masterTanggalMaker = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                        masterDataMaker.Add(DateProcessedDate, masterTanggalMaker);
                    }
                    if (masterDataCrr.ContainsKey(DateProcessedDate) == false)
                    {
                        Dictionary<string, Dictionary<string, Dictionary<string, double>>> masterTanggalCrr = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                        masterDataCrr.Add(DateProcessedDate, masterTanggalCrr);
                    }
                    if (masterDataOutput.ContainsKey(DateProcessedDate) == false)
                    {
                        Dictionary<string, Dictionary<string, Dictionary<string, double>>> masterTanggalOutput = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                        masterDataOutput.Add(DateProcessedDate, masterTanggalOutput);
                    }

                    for (int shift = 1; shift < 4; shift++)
                    {
                        string Stingshift = shift.ToString();
                        Dictionary<string, Dictionary<string, Dictionary<string, double>>> masterShift = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                        Dictionary<string, Dictionary<string, Dictionary<string, double>>> masterShiftMaker = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                        Dictionary<string, Dictionary<string, double>> masterShiftCrr = new Dictionary<string, Dictionary<string, double>>();
                        Dictionary<string, Dictionary<string, double>> masterShiftOutput = new Dictionary<string, Dictionary<string, double>>();
                        masterDataPacker[DateProcessedDate].Add(Stingshift, masterShift);
                        masterDataMaker[DateProcessedDate].Add(Stingshift, masterShiftMaker);
                        masterDataCrr[DateProcessedDate].Add(Stingshift, masterShiftCrr);
                        masterDataOutput[DateProcessedDate].Add(Stingshift, masterShiftOutput);
                        List<LPHSubmissionsModel> subShiftIni = lphListBase.Where(x => x.Date == processedDate && x.Shift.Trim() == Stingshift).ToList();
                        List<LPHSubmissionsModel> subPacker = subShiftIni.Where(x => x.LPHHeader == "PackerController").ToList();
                        List<LPHSubmissionsModel> subGW = subShiftIni.Where(x => x.LPHHeader == "GWGeneralController").ToList();
                        List<LPHSubmissionsModel> subRipper = subShiftIni.Where(x => x.LPHHeader == "RipperController").ToList();
                        var LPHsIDShiftIni = subShiftIni.Select(x => x.LPHID).ToList();
                        var SubsIDShiftIni = subShiftIni.Select(x => x.ID).ToList();
                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        if (LPHsIDShiftIni.Count == 0)
                        {
                            continue;
                            //return Json(new { Status = "False", data = "Data Not Found" }, JsonRequestBehavior.AllowGet);
                        }
                        string myQueryGW = @"SELECT * FROM LPHExtras where HeaderName = 'crr' AND IsDeleted = '0' AND LPHID IN (" + string.Join(",", LPHsIDShiftIni) + ")";
                        string myQueryRipper = @"SELECT * FROM LPHExtras where (HeaderName = 'inputTable' OR HeaderName = 'output') AND IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHsIDShiftIni) + ")";
                        DataSet dsetGW = new DataSet();
                        DataSet dsetRipper = new DataSet();
                        using (SqlConnection con = new SqlConnection(strConString))
                        {
                            SqlCommand cmdGW = new SqlCommand(myQueryGW, con);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmdGW))
                            {
                                da.Fill(dsetGW);
                            }
                            SqlCommand cmdRipper = new SqlCommand(myQueryRipper, con);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmdRipper))
                            {
                                da.Fill(dsetRipper);
                            }
                        }
                        jsondata = JsonConvert.SerializeObject(dsetGW.Tables[0]);
                        List<LPHExtrasModel> gwExtraData = jsondata.DeserializeToLPHExtraList();

                        jsondata = JsonConvert.SerializeObject(dsetRipper.Tables[0]);
                        List<LPHExtrasModel> ripperExtraData = jsondata.DeserializeToLPHExtraList();

                        List<LPHExtrasModel> gwExtraItemValColumn = gwExtraData
                            .Where(x => x.FieldName == "ItemVal").ToList();
                        List<LPHExtrasModel> ripperExtraItemValColumnInput = ripperExtraData
                            .Where(x => x.FieldName == "TypeInput").ToList();

                        //List<QueryFilter> filterGw = new List<QueryFilter>
                        //{
                        //    new QueryFilter("HeaderName", "crr")
                        //};
                        //string gwED = _lphExtrasAppService.Find(filterGw);
                        //List<LPHExtrasModel> gwExtraData = gwED.DeserializeToLPHExtraList()
                        //    .Where(x => subGW.Any(y => y.LPHID == x.LPHID) && x.IsDeleted == false).ToList();
                        //List<LPHExtrasModel> gwExtraItemValColumn = gwExtraData
                        //    .Where(x => x.FieldName == "ItemVal").ToList();

                        //List<QueryFilter> filterRipper = new List<QueryFilter>
                        //{
                        //    new QueryFilter("HeaderName", "inputTable", Operator.Equals, Operation.OrElse),
                        //    new QueryFilter("HeaderName", "output", Operator.Equals, Operation.OrElse)
                        //};

                        //string ripperED = _lphExtrasAppService.Find(filterRipper);
                        //List<LPHExtrasModel> ripperExtraData = ripperED.DeserializeToLPHExtraList()
                        //    .Where(x => subRipper.Any(y => y.LPHID == x.LPHID) && x.IsDeleted == false).ToList();
                        //List<LPHExtrasModel> ripperExtraItemValColumnInput = ripperExtraData
                        //    .Where(x => x.FieldName == "TypeInput").ToList();
                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        string myQueryNameCompo = @"SELECT * FROM LPHComponents WHERE IsDeleted = 0 AND LPHID IN (" + string.Join(",", LPHsIDShiftIni) + ")";
                        string myQueryVcompo = @"SELECT * FROM LPHValues WHERE IsDeleted = 0 AND SubmissionID IN (" + string.Join(",", SubsIDShiftIni) + ")";


                        DataSet dsetNameCompo = new DataSet();
                        DataSet dsetVcompo = new DataSet();
                        using (SqlConnection con = new SqlConnection(strConString))
                        {
                            SqlCommand cmdNameCompo = new SqlCommand(myQueryNameCompo, con);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmdNameCompo))
                            {
                                da.Fill(dsetNameCompo);
                            }

                            SqlCommand cmdVcompo = new SqlCommand(myQueryVcompo, con);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmdVcompo))
                            {
                                da.Fill(dsetVcompo);
                            }
                        }

                        jsondata = JsonConvert.SerializeObject(dsetNameCompo.Tables[0]);
                        List<LPHComponentsModel> AllComponents = jsondata.DeserializeToLPHComponentList();
                        jsondata = JsonConvert.SerializeObject(dsetVcompo.Tables[0]);
                        List<LPHValuesModel> AllValues = jsondata.DeserializeToLPHValueList();


                        //untuk packer dan maker
                        foreach (var packerSubmission in subPacker)
                        {
                            List<LPHComponentsModel> packerComponents = AllComponents
                            .Where(x => x.LPHID == packerSubmission.LPHID).ToList();
                            List<LPHValuesModel> packerValues = AllValues
                                .Where(x => x.SubmissionID == packerSubmission.ID).ToList();
                            //string LPHComponentString = _lphComponentsAppService.FindBy("LPHID", packerSubmission.LPHID);
                            //List<LPHComponentsModel> packerComponents = LPHComponentString.DeserializeToLPHComponentList();
                            //string LPHValuesString = _lphValuesAppService.FindBy("SubmissionID", packerSubmission.ID);
                            //List<LPHValuesModel> packerValues = LPHValuesString.DeserializeToLPHValueList();
                            //strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            //string myQueryNameCompo = @"SELECT * FROM LPHComponents WHERE LPHID = " + packerSubmission.LPHID;
                            //string myQueryVcompo = @"SELECT * FROM LPHValues WHERE SubmissionID = " + packerSubmission.LPHID;


                            //DataSet dsetNameCompo = new DataSet();
                            //DataSet dsetVcompo = new DataSet();
                            //using (SqlConnection con = new SqlConnection(strConString))
                            //{
                            //    SqlCommand cmdNameCompo = new SqlCommand(myQueryNameCompo, con);
                            //    using (SqlDataAdapter da = new SqlDataAdapter(cmdNameCompo))
                            //    {
                            //        da.Fill(dsetNameCompo);
                            //    }

                            //    SqlCommand cmdVcompo = new SqlCommand(myQueryVcompo, con);
                            //    using (SqlDataAdapter da = new SqlDataAdapter(cmdVcompo))
                            //    {
                            //        da.Fill(dsetVcompo);
                            //    }
                            //}

                            //jsondata = JsonConvert.SerializeObject(dsetNameCompo.Tables[0]);
                            //List<LPHComponentsModel> packerComponents = jsondata.DeserializeToLPHComponentList();
                            //jsondata = JsonConvert.SerializeObject(dsetVcompo.Tables[0]);
                            //List<LPHValuesModel> packerValues = jsondata.DeserializeToLPHValueList();
                            LPHComponentsModel machModel = packerComponents.Where(x => x.ComponentName.Trim() == "generalInfo-MachInfo").FirstOrDefault();
                            LPHValuesModel machValue = packerValues.Where(x => x.LPHComponentID == machModel.ID).FirstOrDefault();

                            if (machValue.Value != null)
                            {

                                string machineFilter = "";
                                if (machine != "" && machine !="All")
                                {
                                    if (machine.Substring(2) == machValue.Value.Substring(2))
                                    {
                                        //jaga2 kalo inputannya bukan LU
                                        machineFilter = "LU" + machValue.Value.Substring(2);
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    machineFilter = machValue.Value;
                                    MachineModel mesinmodel = _machineAppService.GetBy("Code", machineFilter).DeserializeToMachine();
                                    machineFilter = mesinmodel.LinkUp;
                                    if (machineFilter == null)
                                    {
                                        machineFilter = "LU" + mesinmodel.Code.Substring(2);
                                    }
                                }
                                

                                if (masterDataPacker[DateProcessedDate][Stingshift].ContainsKey(machineFilter.Trim()) == false)
                                {
                                    Dictionary<string, Dictionary<string, double>> masterMachine = new Dictionary<string, Dictionary<string, double>>();
                                    masterDataPacker[DateProcessedDate][Stingshift].Add(machineFilter.Trim(), masterMachine);
                                }

                                if (masterDataMaker[DateProcessedDate][Stingshift].ContainsKey(machineFilter.Trim()) == false)
                                {
                                    Dictionary<string, Dictionary<string, double>> masterMachineMaker = new Dictionary<string, Dictionary<string, double>>();
                                    masterDataMaker[DateProcessedDate][Stingshift].Add(machineFilter.Trim(), masterMachineMaker);
                                }

                                LPHComponentsModel brandModel = packerComponents.Where(x => x.ComponentName.Trim() == "generalInfo-BrandInfo").FirstOrDefault();
                                LPHValuesModel brandValue = packerValues.Where(x => x.LPHComponentID == brandModel.ID).FirstOrDefault();
                                if (brandValue.Value != null)
                                {
                                    #region ::Packer::
                                    //untuk Packer

                                    BrandModel brand = _brandAppService.GetBy("Code", brandValue.Value.Trim()).DeserializeToBrand();
                                    
                                    if (masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()].ContainsKey(brandValue.Value.Trim()) == false)
                                    {
                                        Dictionary<string, double> masterBrand = new Dictionary<string, double>();
                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()].Add(brandValue.Value.Trim(), masterBrand);
                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("Output", 0);
                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("CigRodReject", 0);
                                        
                                        //weight GW
                                        foreach (var DataExtra in gwExtraItemValColumn)
                                        {
                                            var exDataGW = gwExtraData.Where(x => x.LPHID == DataExtra.LPHID && x.RowNumber == DataExtra.RowNumber).ToList();
                                            LPHExtrasModel extraSelectedRowBrand = exDataGW.Where(x => x.FieldName == "FACode" && x.Value.Trim() == brand.Code.Trim()).FirstOrDefault();
                                            LPHExtrasModel extraSelectedRowMesin = exDataGW.Where(x => x.FieldName == "AsalLU" && x.Value.Trim() == machineFilter.Trim()).FirstOrDefault();
                                            LPHExtrasModel extraSelectedRowLPH = exDataGW.Where(x => x.FieldName == "MachineVal" && x.Value.Trim() == "Packer").FirstOrDefault();
                                            //get weight
                                            if (extraSelectedRowBrand != null && extraSelectedRowMesin != null && extraSelectedRowLPH != null)
                                            {
                                                LPHExtrasModel extraSelectedRowBerat = exDataGW.Where(x => x.FieldName == "WeightVal").FirstOrDefault();
                                                if (DataExtra.Value.Trim() == "Cig Reject" || DataExtra.Value.Trim() == "CGI REJECT")
                                                {
                                                    double cigreject = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()]["CigRodReject"] += cigreject;
                                                }
                                                //if (DataExtra.Value == "CTF")
                                                //{
                                                //    double beratCtf = double.Parse(extraSelectedRowBerat.Value);
                                                //    totalctf += beratCtf;
                                                //}
                                            }
                                            if (extraSelectedRowBrand == null && extraSelectedRowMesin != null && extraSelectedRowLPH != null)
                                            {
                                                LPHExtrasModel extraSelectedRowBerat = exDataGW.Where(x => x.FieldName == "WeightVal").FirstOrDefault();
                                                LPHExtrasModel extraSelectedRowBrandLain = exDataGW.Where(x => x.FieldName == "FACode").FirstOrDefault();
                                                if (DataExtra.Value.Trim() == "Cig Reject" || DataExtra.Value.Trim() == "CGI REJECT")
                                                {
                                                    if (masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()].ContainsKey(extraSelectedRowBrandLain.Value.Trim()) == false)
                                                    {
                                                        Dictionary<string, double> masterBrand2 = new Dictionary<string, double>();
                                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()].Add(extraSelectedRowBrandLain.Value.Trim(), masterBrand2);
                                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("Output", 0);
                                                        masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("CigRodReject", 0);
                                                    }
                                                    double cigreject = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()]["CigRodReject"] += cigreject;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region ::Maker::
                                    //untuk Maker
                                    if (masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()].ContainsKey(brandValue.Value.Trim()) == false)
                                    {
                                        //BrandModel brand = _brandAppService.GetBy("Code", brandValue.Value.Trim()).DeserializeToBrand();
                                        double? BeratCigarette = brand.BeratCigarette;
                                        double CTF = brand.PackToStick;//sementara pakai ini dulu
                                        Dictionary<string, double> masterBrandMaker = new Dictionary<string, double>();
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()].Add(brandValue.Value.Trim(), masterBrandMaker);
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("Output", 0);
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("CigRodReject", 0);
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("TobRodReject", 0);
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("BeratCigarette", Convert.ToDouble(BeratCigarette));
                                        masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()].Add("CTF", Convert.ToDouble(brand.CTF));


                                        //weight GW
                                        foreach (var DataExtra in gwExtraItemValColumn)
                                        {
                                            var exDataGW = gwExtraData.Where(x => x.LPHID == DataExtra.LPHID && x.RowNumber == DataExtra.RowNumber).ToList();

                                            LPHExtrasModel extraSelectedRowBrand = exDataGW.Where(x => x.FieldName == "FACode" && x.Value.Trim() == brand.Code.Trim()).FirstOrDefault();
                                            LPHExtrasModel extraSelectedRowMesin = exDataGW.Where(x => x.FieldName == "AsalLU" && x.Value.Trim() == machineFilter.Trim()).FirstOrDefault();
                                            LPHExtrasModel extraSelectedRowLPH = exDataGW.Where(x => x.FieldName == "MachineVal" && x.Value.Trim() == "Maker").FirstOrDefault();
                                            //get weight
                                            if (extraSelectedRowBrand != null && extraSelectedRowMesin != null && extraSelectedRowLPH != null)
                                            {
                                                LPHExtrasModel extraSelectedRowBerat = exDataGW.Where(x => x.FieldName == "WeightVal").FirstOrDefault();
                                                if (DataExtra.Value.Trim() == "Cig Reject" || DataExtra.Value.Trim() == "CGI REJECT")
                                                {
                                                    double cigreject = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()]["CigRodReject"] += cigreject;
                                                }
                                                if (DataExtra.Value == "CTF")
                                                {
                                                    double beratCtf = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()]["TobRodReject"] += beratCtf;
                                                }
                                            }
                                            //get weight selain brand packer
                                            if (extraSelectedRowBrand == null && extraSelectedRowMesin != null && extraSelectedRowLPH != null)
                                            {
                                                LPHExtrasModel extraSelectedRowBerat = exDataGW.Where(x => x.FieldName == "WeightVal").FirstOrDefault();
                                                LPHExtrasModel extraSelectedRowBrandLain = exDataGW.Where(x => x.FieldName == "FACode").FirstOrDefault();

                                                if (masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()].ContainsKey(extraSelectedRowBrandLain.Value.Trim()) == false)
                                                {
                                                    BrandModel brand2 = _brandAppService.GetBy("Code", extraSelectedRowBrandLain.Value.Trim()).DeserializeToBrand();
                                                    double? BeratCigarette2 = brand2.BeratCigarette;
                                                    double CTF2 = brand2.PackToStick;//sementara pakai ini dulu
                                                    Dictionary<string, double> masterBrand2 = new Dictionary<string, double>();
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()].Add(extraSelectedRowBrandLain.Value.Trim(), masterBrand2);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("Output", 0);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("CigRodReject", 0);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("TobRodReject", 0);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("BeratCigarette", Convert.ToDouble(BeratCigarette2));
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()].Add("CTF", Convert.ToDouble(brand2.CTF));
                                                }
                                                if (DataExtra.Value.Trim() == "Cig Reject" || DataExtra.Value.Trim() == "CGI REJECT")
                                                {
                                                    double cigreject = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()]["CigRodReject"] += cigreject;
                                                }
                                                if (DataExtra.Value == "CTF")
                                                {
                                                    double beratCtf = double.Parse(extraSelectedRowBerat.Value);
                                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][extraSelectedRowBrandLain.Value.Trim()]["TobRodReject"] += beratCtf;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    LPHComponentsModel resultProductC = null;
                                    //cek PMID atau Bukan
                                    if (prodCenter == 5)
                                    {
                                        resultProductC = packerComponents.Where(x => x.ComponentName.Trim() == "Closingcol3-TotalClosing").FirstOrDefault();
                                    }
                                    else
                                    {
                                        resultProductC = packerComponents.Where(x => x.ComponentName.Trim() == "productInformation-ResultProduct").FirstOrDefault();
                                    }
                                    LPHValuesModel resultProductV = packerValues.Where(x => x.LPHComponentID == resultProductC.ID).FirstOrDefault() ?? null;

                                    int resultProduct = 0;
                                    if (resultProductV != null)
                                    {
                                        if (!Int32.TryParse(resultProductV.Value, out resultProduct))
                                        {
                                            resultProduct = 0;
                                        }
                                    }

                                    //olah output
                                    //BrandModel brandPackerOutput = _brandAppService.GetBy("Code", brandValue.Value.Trim()).DeserializeToBrand();
                                    int boxtoslof = brand.BoxToSlof;
                                    int sloftopack = brand.SlofToPack;
                                    int packtostick = brand.PackToStick;
                                    int resultProductMaker = resultProduct * packtostick;
                                    int resultProductPacker = resultProduct / (boxtoslof * sloftopack);
                                    

                                    //packer output
                                    masterDataPacker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()]["Output"] += resultProductPacker;
                                    //maker output
                                    masterDataMaker[DateProcessedDate][Stingshift][machineFilter.Trim()][brandValue.Value.Trim()]["Output"] += resultProductMaker;
                                }
                            }
                        }
                        #region ::STOCK CIGARETTE REJECT REGULER::



                        //Sum semua summary weight pada tabel crr GW GENERAL
                        List<LPHExtrasModel> gwExtraItemValColumnWeight = gwExtraData
                            .Where(x => x.FieldName == "WeightVal").ToList();
                        double SumWeight = gwExtraItemValColumnWeight.Sum(x => double.Parse(x.Value));

                        foreach (var DataRipperExtraData in ripperExtraItemValColumnInput)
                        {
                            foreach (var dataItemCrr in itemCrr)
                            {
                                var exDataRipper = ripperExtraData.Where(x => x.LPHID == DataRipperExtraData.LPHID && x.RowNumber == DataRipperExtraData.RowNumber).ToList();
                                LPHExtrasModel extraSelectedRowBrand = exDataRipper.Where(x => x.FieldName == "FACodeInput").FirstOrDefault();
                                if (masterDataCrr[DateProcessedDate][Stingshift].ContainsKey(extraSelectedRowBrand.Value.Trim()) == false)
                                {
                                    Dictionary<string, double> masterBrandCrr = new Dictionary<string, double>();
                                    masterDataCrr[DateProcessedDate][Stingshift].Add(extraSelectedRowBrand.Value.Trim(), masterBrandCrr);

                                    //memberi index untuk semua item pada 1 brand
                                    foreach (var dataItemCrr2 in itemCrr)
                                    {
                                        //rename item hard code mengikuti reference 
                                        if (dataItemCrr2.Code == "CGI REJECT")
                                        {
                                            masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add("Cig Reject", 0);
                                        }
                                        else
                                        {
                                            masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add(dataItemCrr2.Code, 0);
                                        }

                                    }
                                    BrandModel brand = _brandAppService.GetBy("Code", extraSelectedRowBrand.Value.Trim()).DeserializeToBrand();

                                    masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add("In", SumWeight);
                                    masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add("CTW", Convert.ToDouble(brand.CTW));
                                    masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add("BeratCigarette", Convert.ToDouble(brand.BeratCigarette));
                                    masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()].Add("BrandCTF", Convert.ToDouble(brand.CTF));//sementara

                                }
                                //rename
                                string renameCGIREJECT = "";
                                if (DataRipperExtraData.Value == "CGI REJECT")
                                {
                                    renameCGIREJECT = "Cig Reject";
                                }
                                if (DataRipperExtraData.Value == dataItemCrr.Code || renameCGIREJECT == dataItemCrr.Code)
                                {
                                    LPHExtrasModel extraSelectedRowBerat = exDataRipper.Where(x => x.FieldName == "WeightInput").FirstOrDefault();
                                    double WeightItem = double.Parse(extraSelectedRowBerat.Value);
                                    masterDataCrr[DateProcessedDate][Stingshift][extraSelectedRowBrand.Value.Trim()][dataItemCrr.Code] += WeightItem;
                                }
                            }
                        }
                        #endregion

                        #region ::OUTPUT RIPPER SHORT::


                        foreach (var DatasubRipper in subRipper)
                        {
                            List<LPHComponentsModel> ripperComponents = AllComponents
                            .Where(x => x.LPHID == DatasubRipper.LPHID).ToList();
                            List<LPHValuesModel> ripperValues = AllValues
                                .Where(x => x.SubmissionID == DatasubRipper.ID).ToList();
                            //string LPHComponentString = _lphComponentsAppService.FindBy("LPHID", DatasubRipper.LPHID);
                            //List<LPHComponentsModel> ripperComponents = LPHComponentString.DeserializeToLPHComponentList();
                            //string LPHValuesString = _lphValuesAppService.FindBy("SubmissionID", DatasubRipper.ID);
                            //List<LPHValuesModel> ripperValues = LPHValuesString.DeserializeToLPHValueList();
                            LPHComponentsModel brandModel = ripperComponents.Where(x => x.ComponentName.Trim() == "Brand").FirstOrDefault();
                            LPHValuesModel brandValue = ripperValues.Where(x => x.LPHComponentID == brandModel.ID).FirstOrDefault();

                            
                            
                            if (brandValue.Value != null)
                            {
                                BrandModel brand = _brandAppService.GetBy("Code", brandValue.Value.Trim()).DeserializeToBrand();
                                if (masterDataOutput[DateProcessedDate][Stingshift].ContainsKey(brandValue.Value.Trim()) == false)
                                {
                                    List<LPHExtrasModel> ripperExtraItemValColumnOutput = ripperExtraData
                                        .Where(x => x.FieldName == "TypeOutput" && x.LPHID == DatasubRipper.LPHID).ToList();
                                    if (ripperExtraItemValColumnOutput.Count > 0)
                                    {
                                        Dictionary<string, double> masterBrandRipper = new Dictionary<string, double>();
                                        masterDataOutput[DateProcessedDate][Stingshift].Add(brandValue.Value.Trim(), masterBrandRipper);
                                        masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].Add("BeratCigarette", Convert.ToDouble(brand.BeratCigarette));
                                        masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].Add("CTF", Convert.ToDouble(brand.CTF));
                                        masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].Add("CTW", Convert.ToDouble(brand.CTW));
                                        masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].Add("yield", 0);

                                    
                                        double TotalReclaim = 0;
                                    
                                        foreach (var dataripperextradata in ripperExtraItemValColumnOutput)
                                        {
                                            foreach (var dataRIOutput in RIOutput)
                                            {
                                                if (masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].ContainsKey(dataRIOutput.Code) == false)
                                                {
                                                    masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()].Add(dataRIOutput.Code, 0);

                                                }
                                                //mengisi value index yang ada di dalam brand
                                                if (dataripperextradata.Value == dataRIOutput.Code)
                                                {
                                                    string masterDataOutput1 = DateProcessedDate;
                                                    string shift1 = Stingshift;
                                                    string brand3 = brandValue.Value.Trim();
                                                    string type = dataRIOutput.Code;
                                                    LPHExtrasModel extraselectedrowberat = ripperExtraData.Where(x => x.LPHID == dataripperextradata.LPHID && x.RowNumber == dataripperextradata.RowNumber && x.FieldName == "WeightOutput").FirstOrDefault();
                                                    double weightitem = double.Parse(extraselectedrowberat.Value);
                                                    masterDataOutput[DateProcessedDate][Stingshift][brandValue.Value.Trim()][dataRIOutput.Code] += weightitem;

                                                    if (dataripperextradata.Value == "Reclaim")
                                                    {
                                                        TotalReclaim += weightitem;
                                                    }
                                                }
                                            }
                                        }
                                        //cek code brand, jika code brand ada yang sama dengan yang di tabel stock CRR lansung hitung yield berdasarkan isi tabel stock crr yang brandnya sama
                                        //jika brand code tidak ada yang sama dengan yang di tabel stock CRR, cari bran lain yang RSCODE nya sama.
                                        //cek brand lain tersebut apakah ada di tabel stock CRR, jika ada hitung yield, jika tidak yield =0;
                                        //pencarian brand dilakukan sesuai date, dan shift yang sama.
                                        string brandCode = brandValue.Value;
                                        if (masterDataCrr[DateProcessedDate][Stingshift].ContainsKey(brandCode) == true)
                                        {
                                            double yield = 0;

                                            Dictionary<string, double> dic = masterDataCrr[DateProcessedDate][Stingshift][brandCode];

                                            double inValue = dic["In"];
                                            double cigreject = dic["Cig Reject"];
                                            double ctw = dic["CTW"];
                                            double beratcig = dic["BeratCigarette"];
                                            double ctfbrand = dic["BrandCTF"];
                                            double DustKasar = dic["Dust Kasar"];
                                            double CTF = 0;
                                            if (dic.ContainsKey("CTF") == false)
                                            {
                                                CTF = 0;
                                            }
                                            else
                                            {
                                                CTF = dic["CTF"];
                                            }
                                            yield = (TotalReclaim / ctw) / ((cigreject / beratcig) + (CTF / ctfbrand) + (DustKasar / ctw));

                                            yield = double.IsNaN(yield) ? 0 : yield;
                                            masterDataOutput[DateProcessedDate][Stingshift][brandCode.Trim()]["yield"] = yield;
                                        }
                                        else
                                        {
                                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                            myQuery = @"SELECT * FROM Brands WHERE code = '" + brandCode + "'";

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
                                            BrandModel tempBrand = jsondata.DeserializeToBrandList().FirstOrDefault();

                                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                                            myQuery = @"SELECT * FROM Brands WHERE RSCode = '" + tempBrand.RSCode + "'";

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
                                            List<BrandModel> brandRsCode = jsondata.DeserializeToBrandList();

                                            double TotalYield = 0;
                                            foreach (var DataBrandRSC in brandRsCode)
                                            {
                                                if (masterDataCrr[DateProcessedDate][Stingshift].ContainsKey(DataBrandRSC.Code) == true)
                                                {
                                                    Dictionary<string, double> dic = masterDataCrr[DateProcessedDate][Stingshift][DataBrandRSC.Code];
                                                    double inValue = dic["In"];
                                                    double cigreject = dic["Cig Reject"];
                                                    double ctw = dic["CTW"];
                                                    double beratcig = dic["BeratCigarette"];
                                                    double ctfbrand = dic["BrandCTF"];
                                                    double DustKasar = dic["Dust Kasar"];
                                                    double CTF = 0;
                                                    if (dic.ContainsKey("CTF") == false)
                                                    {
                                                        CTF = 0;
                                                    }
                                                    else
                                                    {
                                                        CTF = dic["CTF"];
                                                    }
                                                    TotalYield += (TotalReclaim / ctw) / ((cigreject / beratcig) + (CTF / ctfbrand) + (DustKasar / ctw));
                                                }
                                            }
                                            TotalYield = double.IsNaN(TotalYield) ? 0 : TotalYield;
                                            masterDataOutput[DateProcessedDate][Stingshift][brandCode.Trim()]["yield"] = TotalYield;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        string date = DateProcessedDate;
                        string shiftString = Stingshift;

                        foreach (string brandNow in masterDataCrr[date][shiftString].Keys)
                        {
                            double stockAwal = 0;
                            if (masterDataCrr.Count() != 1 || masterDataCrr[date].Count() != 1)
                            {
                                bool shouldBreak = false;
                                foreach (string dateKey in masterDataCrr.Keys.Reverse())
                                {
                                    foreach (string shiftKey in masterDataCrr[dateKey].Keys.Reverse())
                                    {
                                        if (masterDataCrr[dateKey][shiftKey].ContainsKey(brandNow) && masterDataCrr[dateKey][shiftKey][brandNow].ContainsKey("StockAkhir"))
                                        {
                                            stockAwal = masterDataCrr[dateKey][shiftKey][brandNow]["StockAkhir"];
                                            shouldBreak = true;
                                            break;
                                        }
                                    }
                                    if (shouldBreak) break;
                                }
                            }
                            Dictionary<string, double> dic = masterDataCrr[date][shiftString][brandNow];
                            double inValue = dic["In"];
                            double cigreject = dic["Cig Reject"];
                            double ctw = dic["CTW"];
                            double beratcig = dic["BeratCigarette"];
                            double ctfbrand = dic["BrandCTF"];
                            double DustKasar = dic["Dust Kasar"];
                            double CTF = 0;
                            if (dic.ContainsKey("CTF")== false)
                            {
                                CTF = 0;
                            }
                            else
                            {
                                CTF = dic["CTF"];
                            }

                            double SAP = ((cigreject / beratcig) + (CTF / ctfbrand) + (DustKasar / ctw)) * beratcig;
                            SAP = double.IsNaN(SAP) ? 0 : SAP;
                            double stockAkhir = stockAwal + inValue - SAP;

                            masterDataCrr[date][shiftString][brandNow].Add("StockAwal", stockAwal);
                            masterDataCrr[date][shiftString][brandNow].Add("StockAkhir", stockAkhir);
                            masterDataCrr[date][shiftString][brandNow].Add("SAP", SAP);
                        }

                    }

                }


                return Json(new { Status = "True", ResultMaker = masterDataMaker, ResultPacker = masterDataPacker, ResultGI = masterDataCrr, ResultGR = masterDataOutput, headerGI = itemCrr, headerGR = RIOutput }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new {Status ="False", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        #region ::tidak di pakai::
        [HttpPost]
        public ActionResult GetMakerSegmentWithParam(DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                return PartialView("_MakerSegment");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult GetPackerSegmentWithParam(DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                return PartialView("_PackerSegment");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetGISegmentWithParam(DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                return PartialView("_GISegment");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetGRSegmentWithParam(DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                return PartialView("_GRSegment");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetClaimbleSegmentWithParam(DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                return PartialView("_ClaimableSegment");
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public ActionResult GenerateExcel()//DateTime startDate, DateTime endDate, string prodCenter, string machine)
        {
            try
            {
                // do process here, then send param/result to excel generator
                Session["DownloadExcel_CRR"] = ExcelGenerator.ExportCRRClaimable(AccountName);
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

            if (Session["DownloadExcel_CRR"] != null)
            {
                byte[] data = Session["DownloadExcel_CRR"] as byte[];
                return File(data, "application/octet-stream", "CRR Claimable.xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }
        #endregion

        #region ::Download Excel::
        public ActionResult ExtractExcel(string maker, string packer, string GI, string GR)
        {
            try
            {
                List<List<string>> sheetsMaker = string.IsNullOrEmpty(maker) ?
                    new List<List<string>>() : JsonConvert.DeserializeObject<List<List<string>>>(maker);
                List<List<string>> sheetsPacker = string.IsNullOrEmpty(packer) ?
                    new List<List<string>>() : JsonConvert.DeserializeObject<List<List<string>>>(packer);
                List<List<string>> sheetsGI = string.IsNullOrEmpty(GI) ?
                    new List<List<string>>() : JsonConvert.DeserializeObject<List<List<string>>>(GI);
                List<List<string>> sheetsGR = string.IsNullOrEmpty(GI) ?
                    new List<List<string>>() : JsonConvert.DeserializeObject<List<List<string>>>(GR);
                Session["ReportCrr"] = ExcelGenerator.PPRawDataExtract(new Dictionary<string, List<List<string>>>
                {
                    {"Maker", sheetsMaker },
                    {"Packer", sheetsPacker },
                    {"Stock CIgarette Reject Reguler", sheetsGI },
                    {"Output Ripper Short", sheetsGR }
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
            if (Session["ReportCrr"] != null)
            {
                byte[] data = Session["ReportCrr"] as byte[];
                Session["ReportCrr"] = null;
                return File(data, "application/octet-stream", "ReportCRR.xlsx");
            }
            return new EmptyResult();
        }
        #endregion

        [HttpPost]
        public ActionResult GetMachine(long prodCenterID)
        {
            List<MachineModel> machineList = _machineAppService.GetAll(true).DeserializeToMachineList().Where(x => x.LinkUp != null && x.LinkUp != "").ToList();

            List<long> locationIdList = _locationAppService.GetLocIDListByLocType(prodCenterID, "productioncenter");
            List<string> listLinkUp = machineList.Where(x =>
                    locationIdList.Any(y => y == x.LocationID)
                ).Select(x => x.LinkUp).Distinct().ToList();
            if(listLinkUp.Count > 0)
            {
                return Json(new { Status = false, machine = listLinkUp }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                string machine = _machineAppService.GetAll(true);
                List<MachineModel> machineList2 = machine.DeserializeToMachineList();
                //machineList = machineList.Where(x => x.Code != null).OrderBy(x => x.Code).ToList();
                machineList2 = machineList2.Where(x => x.Location.Contains("PC") || x.Location.Contains("MK")).OrderBy(x => x.Code).ToList();
                machineList2 = machineList2.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
                machineList2 = machineList2.GroupBy(x => x.Code).Select(x => x.FirstOrDefault()).ToList();
                return Json(new { Status = true, machine = machineList2 }, JsonRequestBehavior.AllowGet);
            }
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
        private long GetRefID(string value, List<ReferenceDetailModel> refModelList)
        {
            var result = refModelList.Where(x => x.Code == value).FirstOrDefault();
            return result.ID;
        }
        #endregion

    }
}
