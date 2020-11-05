using ExcelDataReader;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Fast.Web.Controllers.Report
{
    [CustomAuthorize("ReportPPYield")]
    public class ReportPPYieldController : BaseController<PPLPHModel>
    {
        //private readonly IPPReportYieldAppService _ppReportYieldAppService;
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
        private readonly IPPReportYieldsAppService _pPReportYieldsAppService;
        private readonly IPPReportYieldWhitesAppService _pPReportYieldWhitesAppService;
        private readonly IPPReportYieldKreteksAppService _pPReportYieldKreteksAppService;
        private readonly IPPReportYieldMCDietsAppService _pPReportYieldMCDietsAppService;

        public ReportPPYieldController(
        //IPPReportYieldAppService ppReportYieldAppService,
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
        IReferenceDetailAppService referenceDetailAppService,
        IPPReportYieldsAppService pPReportYieldsAppService,
        IPPReportYieldWhitesAppService pPReportYieldWhitesAppService,
        IPPReportYieldKreteksAppService pPReportYieldKreteksAppService,
        IPPReportYieldMCDietsAppService pPReportYieldMCDietsAppService)
        {
            //_ppReportYieldAppService = ppReportYieldAppService;
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
            _pPReportYieldsAppService = pPReportYieldsAppService;
            _pPReportYieldWhitesAppService = pPReportYieldWhitesAppService;
            _pPReportYieldKreteksAppService = pPReportYieldKreteksAppService;
            _pPReportYieldMCDietsAppService = pPReportYieldMCDietsAppService;
        }
        

        // GET: ShiftDaily
        public ActionResult Index()
        {
            GetTempData();
            ViewBag.isWest = 0;
            if (AccountProdCenterID == 4 || AccountProdCenterID == 5)
            {
                ViewBag.isWest = 1;
            }
            ViewBag.YieldState = TempData["YieldState"] == null ? 0 : TempData["YieldState"];
            return View();
        }
        public ActionResult DownloadYieldDietPPTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateMCDiet.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateMCDiet.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
        public ActionResult DownloadYieldKretekPPTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateYieldKretekPP.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateYieldKretekPP.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
        public ActionResult DownloadYieldWhitePPTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateYieldWhitePP.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateYieldWhitePP.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
        public ActionResult DownloadYieldPPTemplate()
        {
            try
            {
                string filepath = Server.MapPath("..") + "\\Templates\\TemplateYieldPP.xlsx";

                if (System.IO.File.Exists(filepath))
                {
                    byte[] fileBytes = GetFile(filepath);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "TemplateYieldPP.xlsx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
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
        public ActionResult Upload(PPReportYieldUploadModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                TempData["YieldState"] = 3;
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

                    int countRows = dt_.Rows.Count;
                    int countColumn = dt_.Columns.Count;

                    List<PPReportYieldModel> listData = new List<PPReportYieldModel>();
                    for(int i = 1; i < countRows; i++)
                    {
                        int innerCell = 0;
                        PPReportYieldModel data = new PPReportYieldModel();
                        //Year
                        int year;
                        string yearS = dt_.Rows[i][innerCell].ToString();
                        if (int.TryParse(dt_.Rows[i][innerCell++].ToString(), out year))
                        {
                            data.Year = year;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Week
                        int week;
                        if (int.TryParse(dt_.Rows[i][innerCell++].ToString(), out week))
                        {
                            data.Week = week;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //Input
                        double input;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out input))
                        {
                            data.Input = input;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //MCInput
                        double MCinput;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out MCinput))
                        {
                            data.MCInput = MCinput;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //NonBom
                        double NonBom;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out NonBom))
                        {
                            data.NonBom = NonBom;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //MCNonBom
                        double MCNonBom;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out MCNonBom))
                        {
                            data.MCNonBom = MCNonBom;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        
                        //DryNonBom
                        double DryNonBom;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out DryNonBom))
                        {
                            data.DryNonBom = DryNonBom;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //DryMatter
                        double DryMatter;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out DryMatter))
                        {
                            data.DryMatter = DryMatter;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        data.LocationID = AccountDepartmentID;
                        data.Location = AccountLocation.Substring(0, 5);//AccountLocation;
                        data.ModifiedBy = AccountName;
                        data.ModifiedDate = DateTime.Now;

                        listData.Add(data);
                    }
                    foreach(PPReportYieldModel prym in listData)
                    {
                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                        filters.Add(new QueryFilter("Week", prym.Week.ToString()));
                        filters.Add(new QueryFilter("Year", prym.Year.ToString()));
                        filters.Add(new QueryFilter("IsDeleted", "0"));

                        string app = _pPReportYieldsAppService.Find(filters);
                        List<PPReportYieldModel> yield_list = app.DeserializeToPPReportYieldList();
                        var yieldModel = yield_list.LastOrDefault();
                        if (yieldModel != null)
                        {
                            yieldModel = _pPReportYieldsAppService.GetById(yieldModel.ID, true).DeserializeToPPReportYield();

                            yieldModel.Input = prym.Input;
                            yieldModel.MCInput = prym.MCInput;
                            yieldModel.NonBom = prym.NonBom;
                            yieldModel.MCNonBom = prym.MCNonBom;
                            yieldModel.DryNonBom = prym.DryNonBom;
                            yieldModel.DryMatter = prym.DryMatter;

                            yieldModel.ModifiedBy = AccountName;
                            yieldModel.ModifiedDate = DateTime.Now;

                            string dataApp = JsonHelper<PPReportYieldModel>.Serialize(yieldModel);
                            _pPReportYieldsAppService.Update(dataApp);
                        }
                        else
                        {
                            string input = JsonHelper<PPReportYieldModel>.Serialize(prym);
                            _pPReportYieldsAppService.Add(input);
                        }
                    }

                    reader.Close();
                    reader.Dispose();
                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
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
        [HttpPost]
        public ActionResult UploadWhite(PPReportYieldUploadModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                TempData["YieldState"] = 3;
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

                    int countRows = dt_.Rows.Count;
                    int countColumn = dt_.Columns.Count;

                    List<PPReportYieldWhiteModel> listData = new List<PPReportYieldWhiteModel>();
                    for (int i = 1; i < countRows; i++)
                    {
                        int innerCell = 0;
                        PPReportYieldWhiteModel data = new PPReportYieldWhiteModel();
                        //Year
                        int year;
                        string yearS = dt_.Rows[i][innerCell].ToString();
                        if (int.TryParse(dt_.Rows[i][innerCell++].ToString(), out year))
                        {
                            data.Year = year;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Week
                        int week;
                        if (int.TryParse(dt_.Rows[i][innerCell++].ToString(), out week))
                        {
                            data.Week = week;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Blend
                        if (dt_.Rows[i][innerCell] != null )
                        {
                            data.Blend = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //InfeedMaterialWet
                        double InfeedMaterialWet;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out InfeedMaterialWet))
                        {
                            data.InfeedMaterialWet = InfeedMaterialWet;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //InfeedMaterialDry
                        double InfeedMaterialDry;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out InfeedMaterialDry))
                        {
                            data.InfeedMaterialDry = InfeedMaterialDry;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //SumInputMaterialDry
                        double SumInputMaterialDry;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out SumInputMaterialDry))
                        {
                            data.SumInputMaterialDry = SumInputMaterialDry;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //RS_AddbackPercen
                        double RS_AddbackPercen;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out RS_AddbackPercen))
                        {
                            data.RS_AddbackPercen = RS_AddbackPercen;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }


                        data.LocationID = AccountDepartmentID;
                        data.Location = AccountLocation.Substring(0, 5);//AccountLocation;
                        data.ModifiedBy = AccountName;
                        data.ModifiedDate = DateTime.Now;

                        listData.Add(data);
                    }
                    foreach (PPReportYieldWhiteModel prym in listData)
                    {
                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                        filters.Add(new QueryFilter("Week", prym.Week.ToString()));
                        filters.Add(new QueryFilter("Year", prym.Year.ToString()));
                        filters.Add(new QueryFilter("Blend", prym.Blend.ToString()));
                        filters.Add(new QueryFilter("IsDeleted", "0"));

                        string app = _pPReportYieldWhitesAppService.Find(filters);
                        List<PPReportYieldWhiteModel> yield_list = app.DeserializeToPPReportYieldWhiteList();
                        var yieldModel = yield_list.LastOrDefault();
                        if (yieldModel != null)
                        {
                            yieldModel = _pPReportYieldWhitesAppService.GetById(yieldModel.ID, true).DeserializeToPPReportYieldWhite();

                            yieldModel.InfeedMaterialWet = prym.InfeedMaterialWet;
                            yieldModel.InfeedMaterialDry = prym.InfeedMaterialDry;
                            yieldModel.SumInputMaterialDry = prym.SumInputMaterialDry;

                            yieldModel.RS_AddbackPercen = prym.RS_AddbackPercen;

                            yieldModel.ModifiedBy = AccountName;
                            yieldModel.ModifiedDate = DateTime.Now;

                            string dataApp = JsonHelper<PPReportYieldWhiteModel>.Serialize(yieldModel);
                            _pPReportYieldWhitesAppService.Update(dataApp);
                        }
                        else
                        {
                            string input = JsonHelper<PPReportYieldWhiteModel>.Serialize(prym);
                            _pPReportYieldWhitesAppService.Add(input);
                        }
                    }

                    reader.Close();
                    reader.Dispose();
                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
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
        [HttpPost]
        public ActionResult UploadKretek(PPReportYieldUploadModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                TempData["YieldState"] = 2;
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
                    #region OV Excel Sheet 2
                    DataTable dt_2 = reader.AsDataSet().Tables[1];

                    int countRows2 = dt_2.Rows.Count;
                    int countColumn2 = dt_2.Columns.Count;

                    List<PPReportYieldKretekModel> listDataOV = new List<PPReportYieldKretekModel>();
                    for (int i = 1; i < countRows2; i++)
                    {
                        int innerCell = 0;
                        PPReportYieldKretekModel dataOV = new PPReportYieldKretekModel();
                        //Blend
                        if (dt_2.Rows[i][innerCell] != null)
                        {
                            dataOV.Blend = dt_2.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Final OV
                        double finalOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out finalOV))
                        {
                            dataOV.AvgFinalOV = finalOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Leaf OV
                        double leafOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out leafOV))
                        {
                            dataOV.LeafOV = leafOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Clove OV
                        double cloveOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out cloveOV))
                        {
                            dataOV.CloveOV = cloveOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Cres OV
                        double cresOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out cresOV))
                        {
                            dataOV.CRESOV = cresOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Diet OV
                        double dietOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out dietOV))
                        {
                            dataOV.DIETOV = dietOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //RTC OV
                        double rtcOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out rtcOV))
                        {
                            dataOV.RTCOV = rtcOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //SmallLamina OV
                        double smallLaminaOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out smallLaminaOV))
                        {
                            dataOV.SmallLaminaOV = smallLaminaOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //CSF OV
                        double csfOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out csfOV))
                        {
                            dataOV.CloveSteamFlakeOV = csfOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //BC Dry Matter
                        double BCDryOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out BCDryOV))
                        {
                            dataOV.BCDryMatter = BCDryOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //AC Dry Matter
                        double ACDryOV;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out ACDryOV))
                        {
                            dataOV.ACDryMatter = ACDryOV;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //MCNonBom
                        double mcNonBom;
                        if (double.TryParse(dt_2.Rows[i][innerCell++].ToString(), out mcNonBom))
                        {
                            dataOV.AddbackMC = mcNonBom;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        dataOV.Year = int.Parse(model.WeekYear.Substring(0, 4));
                        dataOV.Week = int.Parse(model.WeekYear.Substring(6, 2));

                        dataOV.LocationID = AccountDepartmentID;
                        dataOV.Location = AccountLocation.Substring(0, 5);//AccountLocation;

                        listDataOV.Add(dataOV);
                    }
                    //foreach (PPReportYieldKretekModel prym in listDataOV)
                    //{
                    //    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    //    filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                    //    filters.Add(new QueryFilter("Week", prym.Week.ToString()));
                    //    filters.Add(new QueryFilter("Year", prym.Year.ToString()));
                    //    filters.Add(new QueryFilter("Blend", prym.Blend.ToString()));
                    //    filters.Add(new QueryFilter("IsDeleted", "0"));

                    //    string app = _pPReportYieldKreteksAppService.Find(filters);
                    //    List<PPReportYieldKretekModel> yield_list = app.DeserializeToPPReportYieldKretekList();
                    //    var yieldModel = yield_list.LastOrDefault();
                    //    if (yieldModel != null)
                    //    {
                    //        yieldModel = _pPReportYieldKreteksAppService.GetById(yieldModel.ID, true).DeserializeToPPReportYieldKretek();

                    //        yieldModel.AvgFinalOV = prym.AvgFinalOV;
                    //        yieldModel.LeafOV = prym.LeafOV;
                    //        yieldModel.CloveOV = prym.CloveOV;
                    //        yieldModel.CRESOV = prym.CRESOV;
                    //        yieldModel.DIETOV = prym.DIETOV;
                    //        yieldModel.RTCOV = prym.RTCOV;
                    //        yieldModel.SmallLaminaOV = prym.SmallLaminaOV;
                    //        yieldModel.CloveSteamFlakeOV = prym.CloveSteamFlakeOV;
                    //        yieldModel.BCDryMatter = prym.BCDryMatter;
                    //        yieldModel.ACDryMatter = prym.ACDryMatter;
                    //        yieldModel.AddbackMC = prym.AddbackMC;

                    //        yieldModel.ModifiedBy = AccountName;
                    //        yieldModel.ModifiedDate = DateTime.Now;

                    //        string dataApp = JsonHelper<PPReportYieldKretekModel>.Serialize(yieldModel);
                    //        _pPReportYieldKreteksAppService.Update(dataApp);
                    //    }
                    //    else
                    //    {
                    //        string input = JsonHelper<PPReportYieldKretekModel>.Serialize(prym);
                    //        _pPReportYieldKreteksAppService.Add(input);
                    //    }
                    //}
                    #endregion
                    #region SAP Excel Sheet 1
                    DataTable dt_ = reader.AsDataSet().Tables[0];

                    int countRows = dt_.Rows.Count;
                    int countColumn = dt_.Columns.Count;

                    List<PPReportYieldKretekSAPModel> listData = new List<PPReportYieldKretekSAPModel>();
                    for (int i = 1; i < countRows; i++)
                    {
                        int innerCell = 0;
                        PPReportYieldKretekSAPModel data = new PPReportYieldKretekSAPModel();

                        //Material
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.Material = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //MaterialCode
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.MaterialCode = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            data.MaterialCode = ""; innerCell++;
                            //SetFalseTempData(UIResources.FileCorrupted);
                            //return RedirectToAction("Index");
                        }
                        //MaterialDescription
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.MaterialDescription = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //BaseUnitofMeasure
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.BaseUnitofMeasure = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //Category
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.Category = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        //MaterialGroupDesc
                        if (dt_.Rows[i][innerCell] != null)
                        {
                            data.MaterialGroupDesc = dt_.Rows[i][innerCell++].ToString();
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        string xxx = dt_.Rows[i][innerCell].ToString();
                        //GoodsReceiptQty
                        double GoodsReceiptQty;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out GoodsReceiptQty))
                        {
                            data.GoodsReceiptQty = GoodsReceiptQty;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //GoodIssueQty
                        double GoodIssueQty;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out GoodIssueQty))
                        {
                            data.GoodIssueQty = GoodIssueQty;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //WIPQuantity
                        double WIPQuantity;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out WIPQuantity))
                        {
                            data.WIPQuantity = WIPQuantity;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //StockTakeQuantity
                        double StockTakeQuantity;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out StockTakeQuantity))
                        {
                            data.StockTakeQuantity = StockTakeQuantity;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //ScrapQty
                        double ScrapQty;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out ScrapQty))
                        {
                            data.ScrapQty = ScrapQty;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //NonBOMQty
                        double NonBOMQty;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out NonBOMQty))
                        {
                            data.NonBOMQty = NonBOMQty;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }

                        //TotTKGWIP
                        double TotTKGWIP;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out TotTKGWIP))
                        {
                            data.TotTKGWIP = TotTKGWIP;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        /*
                        //YieldPercent
                        double YieldPercent;
                        if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out YieldPercent))
                        {
                            data.YieldPercent = YieldPercent;
                        }
                        else
                        {
                            SetFalseTempData(UIResources.FileCorrupted);
                            return RedirectToAction("Index");
                        }
                        */
                        listData.Add(data);
                    }
                    List<String> listBlend = listData.Select(x => x.Material).Distinct().ToList();
                    List<PPReportYieldKretekModel> listYeldKretek = new List<PPReportYieldKretekModel>();
                    foreach(String blend in listBlend)
                    {
                        PPReportYieldKretekModel pprykm = new PPReportYieldKretekModel();
                        
                        pprykm.Leaf = listData.Where(x => x.MaterialGroupDesc == "Leaf Tobacco LOT" && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.Clove = listData.Where(x => (x.MaterialGroupDesc.Length >= 10 ? x.MaterialGroupDesc.Substring(0, 10) == "Cut Cloves" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.CRES = listData.Where(x => (x.MaterialGroupDesc.Length >= 4 ? x.MaterialDescription.Substring(0, 4) == "Stem" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.DIET = listData.Where(x => (x.MaterialGroupDesc.Length >= 10 ? x.MaterialGroupDesc.Substring(0, 10) == "Expanded T" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.RTC = listData.Where(x => (x.MaterialGroupDesc.Length >= 10 ? x.MaterialGroupDesc.Substring(0, 10) == "Hom Cloves" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.SmallLamina = listData.Where(x => (x.MaterialGroupDesc.Length >= 10 ? x.MaterialGroupDesc.Substring(0, 10) == "Small Lami" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.CloveSteamFlake = listData.Where(x => (x.MaterialGroupDesc.Length >= 3 ? x.MaterialDescription.Substring(0, 3) == "RWK" : false) && x.Material == blend).Sum(x => x.NonBOMQty);
                        pprykm.BCWet = listData.Where(x => x.MaterialDescription.ToUpper().Contains("CASING") && x.Material == blend).Sum(x => x.GoodIssueQty);
                        pprykm.Addback = listData.Where(x => (x.Category.Length >= 7 ? x.Category.Substring(0, 7).ToUpper() == "NON BOM" : false) && x.Material == blend).Sum(x => x.NonBOMQty) - pprykm.CloveSteamFlake;
                        pprykm.CutFiller = listData.Where(x => (x.MaterialGroupDesc.Length >= 10 ? x.MaterialGroupDesc.Substring(0, 10) == "Cut Filler" : false) && x.Material == blend).Sum(x => x.GoodsReceiptQty);
                        pprykm.InfeedMaterialWet = (pprykm.Leaf + pprykm.Clove + pprykm.CRES + pprykm.DIET + pprykm.RTC + pprykm.SmallLamina + pprykm.CloveSteamFlake);
                        pprykm.ACWet = listData.Where(x => (x.MaterialDescription.Length >= 4 ? x.MaterialDescription.Substring(0, 4).ToUpper() == "AFTE" : false) && x.Material == blend).Sum(x => x.GoodIssueQty);


                        //OV 
                        PPReportYieldKretekModel ovData = listDataOV.Where(x => x.Blend == blend).FirstOrDefault();
                        if (ovData != null)
                        {
                            pprykm.AvgFinalOV = ovData.AvgFinalOV;
                            pprykm.LeafOV = ovData.LeafOV;
                            pprykm.CloveOV = ovData.CloveOV;
                            pprykm.CRESOV = ovData.CRESOV;
                            pprykm.DIETOV = ovData.DIETOV;
                            pprykm.RTCOV = ovData.RTCOV;
                            pprykm.SmallLaminaOV = ovData.SmallLaminaOV;
                            pprykm.CloveSteamFlakeOV = ovData.CloveSteamFlakeOV;
                            pprykm.BCDryMatter = ovData.BCDryMatter;
                            pprykm.ACDryMatter = ovData.ACDryMatter;
                            pprykm.AddbackMC = ovData.AddbackMC;
                        }

                        pprykm.InfeedMaterialDry = (pprykm.Leaf * (1 - pprykm.LeafOV / 100)) + (pprykm.Clove * (1 - pprykm.CloveOV / 100)) + (pprykm.CRES * (1 - pprykm.CRESOV / 100)) + (pprykm.DIET * (1 - pprykm.DIETOV / 100)) + (pprykm.RTC * (1 - pprykm.RTCOV / 100)) + (pprykm.SmallLamina * (1 - pprykm.SmallLaminaOV / 100)) + (pprykm.CloveSteamFlake * (1 - pprykm.CloveSteamFlakeOV / 100));
                        pprykm.CFDry = pprykm.CutFiller != 0 ? (100 - pprykm.AvgFinalOV) / 100 * pprykm.CutFiller : 0;
                        pprykm.AddbackDry = pprykm.AddbackMC != null ? pprykm.Addback * (1 - pprykm.AddbackMC) : 0;
                        pprykm.CFDryExclude = pprykm.CFDry - pprykm.AddbackDry;
                        pprykm.DryYield = pprykm.CutFiller > 0 ? pprykm.CFDryExclude / (pprykm.InfeedMaterialDry + (pprykm.BCWet * (pprykm.BCDryMatter != null ? pprykm.BCDryMatter : 0) / 100) + (pprykm.ACWet * (pprykm.ACDryMatter != null ? pprykm.ACDryMatter : 0) / 100)) : 0;
                        pprykm.WetYield = pprykm.CutFiller > 0 ? (pprykm.CutFiller - pprykm.Addback) / pprykm.InfeedMaterialWet : 0;

                        pprykm.Blend = blend;
                        pprykm.Year = int.Parse(model.WeekYear.Substring(0, 4));
                        pprykm.Week = int.Parse(model.WeekYear.Substring(6, 2));

                        pprykm.LocationID = AccountDepartmentID;
                        pprykm.Location = AccountLocation.Substring(0, 5);//AccountLocation;
                        pprykm.ModifiedBy = AccountName;
                        pprykm.ModifiedDate = DateTime.Now;
                        listYeldKretek.Add(pprykm);
                    }
                    foreach (PPReportYieldKretekModel prym in listYeldKretek)
                    {
                        ICollection<QueryFilter> filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                        filters.Add(new QueryFilter("Week", prym.Week.ToString()));
                        filters.Add(new QueryFilter("Year", prym.Year.ToString()));
                        filters.Add(new QueryFilter("Blend", prym.Blend.ToString()));
                        filters.Add(new QueryFilter("IsDeleted", "0"));

                        string app = _pPReportYieldKreteksAppService.Find(filters);
                        List<PPReportYieldKretekModel> yield_list = app.DeserializeToPPReportYieldKretekList();
                        var yieldModel = yield_list.LastOrDefault();
                        if (yieldModel != null)
                        {
                            yieldModel = _pPReportYieldKreteksAppService.GetById(yieldModel.ID, true).DeserializeToPPReportYieldKretek();


                            yieldModel.AvgFinalOV = prym.AvgFinalOV;
                            yieldModel.LeafOV = prym.LeafOV;
                            yieldModel.CloveOV = prym.CloveOV;
                            yieldModel.CRESOV = prym.CRESOV;
                            yieldModel.DIETOV = prym.DIETOV;
                            yieldModel.RTCOV = prym.RTCOV;
                            yieldModel.SmallLaminaOV = prym.SmallLaminaOV;
                            yieldModel.CloveSteamFlakeOV = prym.CloveSteamFlakeOV;
                            yieldModel.BCDryMatter = prym.BCDryMatter;
                            yieldModel.ACDryMatter = prym.ACDryMatter;
                            yieldModel.AddbackMC = prym.AddbackMC;

                            yieldModel.Leaf = prym.Leaf ;
                            yieldModel.Clove = prym.Clove;
                            yieldModel.CRES = prym.CRES;
                            yieldModel.DIET = prym.DIET;
                            yieldModel.RTC = prym.RTC;
                            yieldModel.SmallLamina = prym.SmallLamina;
                            yieldModel.CloveSteamFlake = prym.CloveSteamFlake;
                            yieldModel.InfeedMaterialWet = prym.InfeedMaterialWet;
                            yieldModel.InfeedMaterialDry = prym.InfeedMaterialDry;
                            yieldModel.BCWet = prym.BCWet;
                            yieldModel.ACWet = prym.ACWet;
                            yieldModel.Addback = prym.Addback;
                            yieldModel.AddbackDry = prym.AddbackDry;
                            yieldModel.CutFiller = prym.CutFiller;
                            yieldModel.CFDry = prym.CFDry;
                            yieldModel.CFDryExclude = prym.CFDryExclude;
                            yieldModel.WetYield = prym.WetYield;
                            yieldModel.DryYield = prym.DryYield;

                            yieldModel.ModifiedBy = AccountName;
                            yieldModel.ModifiedDate = DateTime.Now;

                            string dataApp = JsonHelper<PPReportYieldKretekModel>.Serialize(yieldModel);
                            _pPReportYieldKreteksAppService.Update(dataApp);
                        }
                        else
                        {
                            string input = JsonHelper<PPReportYieldKretekModel>.Serialize(prym);
                            _pPReportYieldKreteksAppService.Add(input);
                        }
                    }
                    #endregion
                   
                    reader.Close();
                    reader.Dispose();
                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
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
        [HttpPost]
        public ActionResult UploadDiet(PPReportYieldUploadModel model)
        {
            IExcelDataReader reader = null;

            try
            {
                TempData["YieldState"] = 4;
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

                    #region SAP Excel Sheet 1
                    DataTable dt_ = reader.AsDataSet().Tables[0];

                    int countRows = dt_.Rows.Count;
                    int countColumn = dt_.Columns.Count;

                    int i = 2;
                    int innerCell = 0;
                    PPReportYieldMCDietModel data = new PPReportYieldMCDietModel();

                    //MC Flake
                    double MCFlake;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out MCFlake))
                    {
                        data.MCFlake = MCFlake;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //MCKrosok
                    double MCKrosok;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out MCKrosok))
                    {
                        data.MCKrosok = MCKrosok;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //CVIB0069
                    double CVIB0069;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out CVIB0069))
                    {
                        data.CVIB0069 = CVIB0069;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //CSFR0022
                    double CSFR0022;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out CSFR0022))
                    {
                        data.CSFR0022 = CSFR0022;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //DSCL0034
                    double DSCL0034;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out DSCL0034))
                    {
                        data.DSCL0034 = DSCL0034;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //CVIB0070
                    double CVIB0070;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out CVIB0070))
                    {
                        data.CVIB0070 = CVIB0070;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //RV0054
                    double RV0054;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out RV0054))
                    {
                        data.RV0054 = RV0054;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //DM
                    double DM;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out DM))
                    {
                        data.DM = DM;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //Flake
                    double Flake;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out Flake))
                    {
                        data.Flake = Flake;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }

                    //MCPacking
                    double MCPacking;
                    if (double.TryParse(dt_.Rows[i][innerCell++].ToString(), out MCPacking))
                    {
                        data.MCPacking = MCPacking;
                    }
                    else
                    {
                        SetFalseTempData(UIResources.FileCorrupted);
                        return RedirectToAction("Index");
                    }
                    data.Year = int.Parse(model.WeekYear.Substring(0, 4));
                    data.Week = int.Parse(model.WeekYear.Substring(6, 2));

                    data.LocationID = AccountDepartmentID;
                    data.Location = AccountLocation.Substring(0, 5);//AccountLocation;

                    data.ModifiedBy = AccountName;
                    data.ModifiedDate = DateTime.Now;

                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                    filters.Add(new QueryFilter("Week", data.Week.ToString(),Operator.GreaterThanOrEqual));
                    filters.Add(new QueryFilter("Year", data.Year.ToString()));
                    filters.Add(new QueryFilter("IsDeleted", "0"));

                    string app = _pPReportYieldMCDietsAppService.Find(filters);
                    List<PPReportYieldMCDietModel> yield_list = app.DeserializeToPPReportYieldMCDietList();
                    foreach (PPReportYieldMCDietModel pprymcd in yield_list) {
                        _pPReportYieldMCDietsAppService.Remove(pprymcd.ID);
                    }
                    for (int iAdd = int.Parse(model.WeekYear.Substring(6, 2)); iAdd <= GetWeeksInYear(data.Year); iAdd++)
                    {
                        PPReportYieldMCDietModel newData = data;
                        newData.Week = iAdd;
                        string dataApp = JsonHelper<PPReportYieldMCDietModel>.Serialize(newData);
                        _pPReportYieldMCDietsAppService.Add(dataApp);
                    }
                
                    #endregion

                    reader.Close();
                    reader.Dispose();
                    SetTrueTempData(UIResources.UploadSucceed);
                }
                else
                {
                    SetFalseTempData(UIResources.FileCorrupted);
                    return RedirectToAction("Index");
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



        [HttpPost]
        public ActionResult GetReportWithParam(string ppType) //param nya menyusul
        {
            try
            {
                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", AccountLocationID.ToString()));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string ppList = _ppLphApprovalAppService.Find(filters);
                List<PPLPHApprovalsModel> pps = ppList.DeserializeToPPLPHApprovalList();
                pps = pps.Where(x => x.Status.Trim().ToLower() == "approved").ToList();

                foreach (var item in pps)
                {
                    string subs = _ppLphSubmissionsAppService.FindBy("ID", item.LPHSubmissionID, true);
                    PPLPHSubmissionsModel subsModel = subs.DeserializeToPPLPHSubmissions();
                    item.LPHType = subsModel.LPHHeader.Replace("Controller", "");
                }

                pps = pps.Where(x => x.LPHType == ppType).ToList();


                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<PPLPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #region View Data Clove
        public ActionResult GetData(string StartDate, string EndDate)
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
                /*
                // Getting all data submissions   			
                string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
                List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();
                */
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                int startWeek = getWeek(dtStartDate);
                int endWeek = getWeek(dtEndDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                filters.Add(new QueryFilter("Week", startWeek.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Week", endWeek.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string ppYieldList = _pPReportYieldsAppService.Find(filters);
                List<PPReportYieldModel> yields = ppYieldList.DeserializeToPPReportYieldList().OrderBy(x => x.Week).ToList();
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountDepartmentID, "productioncenter");
                List<String> LPH = new List<String>()
                {
                    "LPHPrimaryCloveInfeedConditioningController",
                    "LPHPrimaryCloveCutDryPackingController"
                };

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Header", "LPHPrimaryCloveInfeedConditioningController"));
                filters.Add(new QueryFilter("Header", "LPHPrimaryCloveCutDryPackingController", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                //List<LPHModel> lphList = _ppLphAppService.Find(filters).DeserializeToLPHList()
                List<LPHModel> lphList = _ppLphAppService.GetAll().DeserializeToLPHList()
                    .Where(x =>
                        locationIdList.Any(y => y == x.LocationID)
                        && LPH.Any(z => x.Header == z) //chanif: sudah ganti filter di atas
                    ).ToList();


                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", dtStartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", dtEndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.Find(filters).DeserializeToPPLPHSubmissionsList()
                    .Where(x => lphList.Any(y => y.ID == x.LPHID)).ToList();

                var minSubmissionID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                var maxSubmissionID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("LPHSubmissionID", minSubmissionID.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("LPHSubmissionID", maxSubmissionID.ToString(), Operator.LessThanOrEqualTo));

                List<PPLPHApprovalsModel> approval = _ppLphApprovalAppService.Find(filters).DeserializeToPPLPHApprovalList();
                lphList = lphList.Where(l =>
                {
                    PPLPHSubmissionsModel s = submissions.Where(x => x.LPHID == l.ID).ToList().FirstOrDefault();
                    if (s != null) return approval.Where(x => x.LPHSubmissionID == s.ID && x.Status.Trim().ToLower() == "approved").Count() > 0;
                    return false;
                }).ToList();
                List<String> compoList = new List<String>()
                {
                    "CCS1DCC",
                    "CCS2DCC",
                    "MCFinal0",
                    "WeighconP11"
                };

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("FieldName", "CCS1DCC"));
                filters.Add(new QueryFilter("FieldName", "CCS2DCC", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "MCFinal0", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "WeighconP11", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<LPHExtrasModel> compoResultList = _ppLphExtrasAppService.Find(filters).DeserializeToLPHExtraList()
                    .Where(x => 
                        lphList.Any(z => z.ID == x.LPHID)
                        && compoList.Any(y => y == x.FieldName)
                    ).ToList();
                foreach (PPReportYieldModel prym in yields)
                {
                    var prymM = _pPReportYieldsAppService.GetById(prym.ID, true).DeserializeToPPReportYield();

                    prymM.DryInput = prymM.Input * (1 - (prymM.MCInput / 100));

                    List<LPHModel> lphThisWeek = lphList.Where(l =>
                    {
                        PPLPHSubmissionsModel s = submissions.Where(x => x.LPHID == l.ID).ToList().FirstOrDefault();
                        if (getWeek(s.Date) == prymM.Week && s.Date.Year == prymM.Year) return true;
                        return false;
                    }).ToList();
                    List<String> ccsList = new List<String>()
                    {
                    "CCS1DCC",
                    "CCS2DCC",
                    };

                    List<LPHExtrasModel> casingResult = compoResultList.Where(x => lphThisWeek.Any(y => y.ID == x.LPHID) && ccsList.Any(z => z == x.FieldName)).ToList();
                    double CasingValue = casingResult.Count == 0 ? 0 : casingResult.Sum(x => double.TryParse(x.Value, out double r) ? r : 0);
                    prymM.Casing = CasingValue;
                    prymM.DryCasing = (prymM.DryMatter / 100) * CasingValue;

                    List<LPHExtrasModel> outputResult = compoResultList.Where(x => lphThisWeek.Any(y => y.ID == x.LPHID) && x.FieldName == "WeighconP11").ToList();
                    double outputValue = outputResult.Count == 0 ? 0 : outputResult.Sum(x => double.TryParse(x.Value, out double r) ? r : 0);
                    prymM.Output = outputValue;

                    List<LPHExtrasModel> mcResult = compoResultList.Where(x => lphThisWeek.Any(y => y.ID == x.LPHID) && x.FieldName == "MCFinal0").ToList();
                    double mcOutputValue = mcResult.Count == 0 ? 0 : mcResult.Average(x => double.TryParse(x.Value, out double r) ? r : 0);
                    prymM.MCOutput = mcOutputValue;

                    prymM.DryOutput = prymM.Output * (1 - (mcOutputValue / 100));

                    prymM.DryYield = (prymM.DryOutput - prymM.DryNonBom) / (prymM.DryInput + prymM.DryCasing);
                    prymM.WetYield = (prymM.Output - prymM.NonBom) / prymM.Input;

                    string dataApp = JsonHelper<PPReportYieldModel>.Serialize(prymM);
                    _pPReportYieldsAppService.Update(dataApp);
                }

                yields = ppYieldList.DeserializeToPPReportYieldList().Where(x => x.Week >= startWeek && x.Week <= endWeek).OrderBy(x => x.Week).ToList();

                int recordsTotal = yields.Count();

                // total number of rows count     
                int recordsFiltered = yields.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = yields.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region View Data White
        public ActionResult GetDataWhite(string StartDate, string EndDate)
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
                /*
                // Getting all data submissions   			
                string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
                List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();
                */
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                int startWeek = getWeek(dtStartDate);
                int endWeek = getWeek(dtEndDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                filters.Add(new QueryFilter("Week", startWeek.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Week", endWeek.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string ppYieldList = _pPReportYieldWhitesAppService.Find(filters);
                List<PPReportYieldWhiteModel> yields = ppYieldList.DeserializeToPPReportYieldWhiteList().OrderBy(x => x.Week).ToList();
                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountDepartmentID, "productioncenter");
                

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Header", "LPHPrimaryWhiteLineOTPController"));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<LPHModel> lphList = _ppLphAppService.Find(filters).DeserializeToLPHList()
                    .Where(x =>
                        locationIdList.Any(y => y == x.LocationID)
                    ).ToList();


                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("Date", dtStartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", dtEndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.Find(filters).DeserializeToPPLPHSubmissionsList()
                    .Where(x => lphList.Any(y => y.ID == x.LPHID)).ToList();

                var minSubmissionID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                var maxSubmissionID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("LPHSubmissionID", minSubmissionID.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("LPHSubmissionID", maxSubmissionID.ToString(), Operator.LessThanOrEqualTo));

                List<PPLPHApprovalsModel> approval = _ppLphApprovalAppService.Find(filters).DeserializeToPPLPHApprovalList();
                lphList = lphList.Where(l =>
                {
                    PPLPHSubmissionsModel s = submissions.Where(x => x.LPHID == l.ID).ToList().FirstOrDefault();
                    if (s != null) return approval.Where(x => x.LPHSubmissionID == s.ID && x.Status.Trim().ToLower() == "approved").Count() > 0;
                    return false;
                }).ToList();
               

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("FieldName", "piBlendCode"));
                filters.Add(new QueryFilter("FieldName", "bcoPackingTotal", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "bcoPackingBeratSatuan", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "bcoPackingPlusKilo", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "bcoPackingMCPacking", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("FieldName", "addback", Operator.Equals, Operation.Or));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<LPHExtrasModel> compoResultList = _ppLphExtrasAppService.Find(filters).DeserializeToLPHExtraList()
                    .Where(x =>
                        lphList.Any(z => z.ID == x.LPHID)
                    ).ToList();
                
                foreach (PPReportYieldWhiteModel prym in yields)
                {
                    var prymM = _pPReportYieldWhitesAppService.GetById(prym.ID, true).DeserializeToPPReportYieldWhite();

                    List<LPHModel> lphThisWeek = lphList.Where(l =>
                    {
                        PPLPHSubmissionsModel s = submissions.Where(x => x.LPHID == l.ID).ToList().FirstOrDefault();
                        if (getWeek(s.Date) == prymM.Week && s.Date.Year == prymM.Year) return true;
                        return false;
                    }).ToList();
                    double cutFillerTotal = 0;
                    double addbackQty = 0;
                    double countMC = 0;
                    double totalMC = 0;
                    foreach (LPHModel lph in lphThisWeek)
                    {
                        List<LPHExtrasModel> listRow = compoResultList.Where(x => x.LPHID == lph.ID && x.FieldName == "piBlendCode" && x.Value == prym.Blend).ToList();
                        foreach (LPHExtrasModel lem in listRow)
                        {
                            //LPHExtrasModel blendCodeResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "piBlendCode").FirstOrDefault();
                            LPHExtrasModel packingResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "bcoPackingTotal").FirstOrDefault();
                            LPHExtrasModel satuanResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "bcoPackingBeratSatuan").FirstOrDefault();
                            LPHExtrasModel recehResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "bcoPackingPlusKilo").FirstOrDefault();
                            LPHExtrasModel addbackResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "addback").FirstOrDefault();
                            LPHExtrasModel mcResult = compoResultList.Where(x => x.LPHID == lph.ID && x.RowNumber == lem.RowNumber && x.FieldName == "bcoPackingMCPacking").FirstOrDefault();

                            totalMC += mcResult == null ? 0 : (double.TryParse(mcResult.Value, out double mc) ? double.Parse(mcResult.Value) : 0);
                            countMC++;
                            List<Dictionary<string, string>> jsonD = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(addbackResult.Value);
                            foreach (var json in jsonD)
                            {
                               addbackQty += (double.TryParse(json["AddbackQty"], out double aq) ? double.Parse(json["AddbackQty"]) : 0);
                            }
                            double packingVal = packingResult == null ? 0 : (double.TryParse(packingResult.Value, out double a) ? double.Parse(packingResult.Value) : 0);
                            double satuanVal = satuanResult == null ? 0 : (double.TryParse(satuanResult.Value, out double b) ? double.Parse(satuanResult.Value) : 0);
                            double recehVal = recehResult == null ? 0 : (double.TryParse(recehResult.Value, out double c) ? double.Parse(recehResult.Value) : 0);
                            cutFillerTotal += (packingVal * satuanVal) + recehVal;
                        }
                    }
                    prymM.CutFiller = cutFillerTotal;

                    prymM.RS_Addback = addbackQty;

                    prymM.AvgFinalOV = countMC == 0 ? 0 : totalMC / countMC;

                    prymM.CFDry = (100 - prymM.AvgFinalOV) / 100 * cutFillerTotal;
                    prymM.RS_AddDry = (100 - prymM.RS_AddbackPercen) / 100 * prymM.RS_Addback;

                    prymM.CFDryExclude = (prymM.CFDry - prymM.RS_AddDry);
                    prymM.RS_AddbackWet = (prymM.CutFiller - prymM.RS_Addback);

                    prymM.DryYield = prymM.CFDryExclude == 0 ? 0 : (prymM.CFDryExclude / prymM.SumInputMaterialDry);
                    prymM.WetYield = prymM.RS_AddbackWet == 0 ? 0 : (prymM.RS_AddbackWet / prymM.InfeedMaterialWet);

                    string dataApp = JsonHelper<PPReportYieldWhiteModel>.Serialize(prymM);
                    _pPReportYieldWhitesAppService.Update(dataApp);
                }
                
                yields = ppYieldList.DeserializeToPPReportYieldWhiteList().Where(x => x.Week >= startWeek && x.Week <= endWeek).OrderBy(x => x.Week).ToList();

                if(yields.Count > 0)
                {
                    PPReportYieldWhiteModel sumYield = new PPReportYieldWhiteModel();
                    double? sumSumInputMaterialDry = yields.Sum(x => x.SumInputMaterialDry);
                    double? sumCFDryExclude = yields.Sum(x => x.CFDryExclude);

                    double? sumRS_AddbackWet = yields.Sum(x => x.RS_AddbackWet);
                    double? sumInfeedMaterialWet = yields.Sum(x => x.InfeedMaterialWet);

                    sumYield.DryYield = sumCFDryExclude == 0 ? 0 : (sumCFDryExclude / sumSumInputMaterialDry);
                    sumYield.WetYield = sumRS_AddbackWet == 0 ? 0 : (sumRS_AddbackWet / sumInfeedMaterialWet);

                    yields.Insert(0, sumYield);
                }
                int recordsTotal = yields.Count();

                // total number of rows count     
                int recordsFiltered = yields.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = yields.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region View Data Kretek
        public ActionResult GetDataKretek(string StartDate, string EndDate)
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
                
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                int startWeek = getWeek(dtStartDate);
                int endWeek = getWeek(dtEndDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                filters.Add(new QueryFilter("Week", startWeek.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Week", endWeek.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string ppYieldList = _pPReportYieldKreteksAppService.Find(filters);
                List<PPReportYieldKretekModel> yields = ppYieldList.DeserializeToPPReportYieldKretekList().OrderBy(x => x.Week).ToList();

                if (yields.Count > 0)
                {
                    PPReportYieldKretekModel summaryYield = new PPReportYieldKretekModel();

                    double? sumBCWet = 0;
                    double? sumACWet = 0;
                    foreach (PPReportYieldKretekModel yil in yields)
                    {
                        sumBCWet += yil.BCDryMatter != null && yil.BCWet != null ? (yil.BCWet * yil.BCDryMatter / 100) : 0;
                        sumACWet += yil.ACDryMatter != null && yil.ACWet != null ? (yil.ACWet * yil.ACDryMatter / 100) : 0;
                    }

                    double? sumCutFiller = yields.Sum(x => x.CutFiller);
                    double? sumAddback = yields.Sum(x => x.Addback);
                    double? sumInfeedMaterialWet = yields.Sum(x => x.InfeedMaterialWet);

                    double? sumInfeedMaterialDry = yields.Sum(x => x.InfeedMaterialDry);
                    double? sumCFDryExclude = yields.Sum(x => x.CFDryExclude);


                    summaryYield.WetYield = sumCutFiller > 0 ? (sumCutFiller - sumAddback) / sumInfeedMaterialWet : 0;
                    summaryYield.DryYield = sumCutFiller > 0 ? sumCFDryExclude / (sumInfeedMaterialDry + sumBCWet + sumACWet) : 0;
                    yields.Insert(0, summaryYield);
                }
                int recordsTotal = yields.Count();

                // total number of rows count     
                int recordsFiltered = yields.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = yields.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        private string DataTableToJSONWithJavaScriptSerializer(DataTable table)
        {
            JavaScriptSerializer jsSerializer = new JavaScriptSerializer();
            List<Dictionary<string, object>> parentRow = new List<Dictionary<string, object>>();
            Dictionary<string, object> childRow;
            foreach (DataRow row in table.Rows)
            {
                childRow = new Dictionary<string, object>();
                foreach (DataColumn col in table.Columns)
                {
                    childRow.Add(col.ColumnName, row[col]);
                }
                parentRow.Add(childRow);
            }
            return jsSerializer.Serialize(parentRow);
        }
        #region View Data White West
        [HttpPost]
        public ActionResult GetDataWhiteWest(string StartDate, string EndDate)
        {
            try
            {

                string strConString = ConfigurationManager.ConnectionStrings["AdoPSS4Conn2"].ConnectionString;
                string strConString2 = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                DataSet dset2 = new DataSet();
                string queryAle = "";
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                string startD = dtStartDate.ToString("MM/dd/yyyy");
                string endD = dtEndDate.ToString("MM/dd/yyyy");
                //int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                //int startWeek = getWeek(dtStartDate);
                //int endWeek = getWeek(dtEndDate);

                //Get Wet Yield
                queryAle = @"SELECT NpssReportBatch.EndTime,NpssReportBatch.SAPID, NpssReportBatch.BatchIdent, NpssReportBatch.BlendCode, NpssReportBatch.ProducedQTY, CAST(NpssReportBatchData.Value As DECIMAL(9, 2)) as Tobacco,
                            NpssReportBatch.TotalStems, NpssReportBatch.TotalExpandedTobacco, NpssReportBatch.TotalSmallLamina,
                            NpssReportBatch.TotalRipperShorts
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000'
                            AND '" + endD + @" 23:59:59.000'
                            AND NpssReportBatchData.DataName = 'totaltobacco' AND NpssReportBatchData.Category = 'End_Batch'
                            ORDER BY NpssReportBatch.BatchIdent
                        
                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) As DMBC
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'BC ISS DryMatterRate' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DMBT
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'BT ISS DryMatterRate' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DMAC
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'FlavourISS DryMatterRate' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as  FinalOV 
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'FinalMM625.056Average' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DryISCRES
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'ISWB620.055TotDryTob' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DryET
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'ETWB620.070TotDryTob' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DrySL
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'SLWB620.175TotDryTob' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DryRS
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'RSWB620.085TotDryTob' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, NpssReportBatch.TotalAfterCut as AC
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'FlavourISS DryMatterRate' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST(NpssReportBatchData.Value As DECIMAL (9,2)) BC
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'Bright.BCPump340.019.TotalBC' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) as DMBS
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'BS ISS DryMatterRate' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST(NpssReportBatchData.Value As DECIMAL (9,2)) as BS
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'Burley.BSPump360.020.TotalBS' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent 

                            SELECT NpssReportBatch.BatchIdent, CAST(NpssReportBatchData.Value As DECIMAL (9,2)) as BT
                            FROM NpssReportBatch INNER JOIN
                            NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                            WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                            AND '" + endD + @" 23:59:59.000' 
                            AND NpssReportBatchData.DataName = 'Burley.BTPump360.041.TotalBT' AND NpssReportBatchData.Category = 'End_Batch' 
                            ORDER BY NpssReportBatch.BatchIdent ";
                using (SqlConnection con = new SqlConnection(strConString))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }
                List<PPReportYieldWhiteWestModel> data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToPPReportYieldWhiteWestList();
                //DMBC
                List<PPReportYieldWhiteWestModel> data2 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[1]).DeserializeToPPReportYieldWhiteWestList();
                //DMBT
                List<PPReportYieldWhiteWestModel> data3 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[2]).DeserializeToPPReportYieldWhiteWestList();
                //DMAC
                List<PPReportYieldWhiteWestModel> data4 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[3]).DeserializeToPPReportYieldWhiteWestList();
                //FinalOV
                List<PPReportYieldWhiteWestModel> data5 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[4]).DeserializeToPPReportYieldWhiteWestList();
                //DryISCRES
                List<PPReportYieldWhiteWestModel> data6 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[5]).DeserializeToPPReportYieldWhiteWestList();
                //DryET
                List<PPReportYieldWhiteWestModel> data7 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[6]).DeserializeToPPReportYieldWhiteWestList();
                //DrySL
                List<PPReportYieldWhiteWestModel> data8 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[7]).DeserializeToPPReportYieldWhiteWestList();
                //DryRS
                List<PPReportYieldWhiteWestModel> data9 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[8]).DeserializeToPPReportYieldWhiteWestList();
                //AC 
                List<PPReportYieldWhiteWestModel> data10 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[9]).DeserializeToPPReportYieldWhiteWestList();
                //BC 
                List<PPReportYieldWhiteWestModel> data11 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[10]).DeserializeToPPReportYieldWhiteWestList();
                //DMBS 
                List<PPReportYieldWhiteWestModel> data12 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[11]).DeserializeToPPReportYieldWhiteWestList();
                //BS 
                List<PPReportYieldWhiteWestModel> data13 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[12]).DeserializeToPPReportYieldWhiteWestList();
                //BT
                List<PPReportYieldWhiteWestModel> data14 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[13]).DeserializeToPPReportYieldWhiteWestList();

                queryAle = @"delete from PPReportYieldWhiteWest where 
                                [EndTime] between '" + startD + @" 00:00:00.000' AND  '" + endD + @" 23:59:59.000'";
                using (SqlConnection con = new SqlConnection(strConString2))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }
                foreach (PPReportYieldWhiteWestModel dt in data)
                {
                    dt.DMBC = data2.Count > 0 ? data2.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DMBC : 0;
                    dt.DMBT = data3.Count > 0 ? data3.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DMBT : 0;
                    dt.DMAC = data4.Count > 0 ? data4.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DMAC : 0;
                    dt.FinalOV = data5.Count > 0 ? data5.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().FinalOV : 0;
                    dt.DryISCRES = data6.Count > 0 ? data6.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryISCRES : 0;
                    dt.DryET = data7.Count > 0 ? data7.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryET : 0;
                    dt.DrySL = data8.Count > 0 ? data8.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DrySL : 0;
                    dt.DryRS = data9.Count > 0 ? data9.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryRS : 0;
                    dt.AC = data10.Count > 0 ? data10.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().AC : 0;
                    dt.BC = data11.Count > 0 ? data11.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().BC : 0;
                    dt.DMBS = data12.Count > 0 ? data12.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DMBS : 0;
                    dt.BS = data13.Count > 0 ? data13.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().BS : 0;
                    dt.BT = data14.Count > 0 ? data14.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().BT : 0;


                    dt.WetYield = ((dt.ProducedQTY - dt.TotalRipperShorts) / (dt.Tobacco + dt.TotalStems + dt.TotalExpandedTobacco + dt.TotalSmallLamina)) * 100;
                    dt.Packing = dt.ProducedQTY * (1 - (dt.FinalOV / 100));
                    dset2 = new DataSet();
                    queryAle = @"SELECT NpssReportBatchData.Value 
                                FROM NpssReportBatch INNER JOIN 
                                NpssReportBatchData 
                                ON NpssReportBatch.BatchID = 
                                NpssReportBatchData.EventID 
                                WHERE NpssReportBatch.BatchIdent =  '"+dt.BatchIdent+@"' AND
                                NpssReportBatchData.DataName = 'InfeedBoxBarcode' AND 
                                NpssReportBatchData.Category = 'End_Batch' ";
                    using (SqlConnection con = new SqlConnection(strConString))
                    {

                        SqlCommand cmd = new SqlCommand(queryAle, con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dset2);

                        Helper.LogErrorMessage("dset count :" + dset2.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                    }

                    double DryWeight = 0;
                    foreach (DataRow rowBarcode in dset2.Tables[0].Rows)
                    {
                        string barcode = rowBarcode[dset2.Tables[0].Columns[0]].ToString();
                        double invoiceMC = 0;
                        if (barcode.Length == 45)
                        {
                            invoiceMC = int.Parse(barcode.Substring(barcode.Length - 4, 4))/100;
                            DryWeight += int.Parse(barcode.Substring(33, 3)) * (1 - invoiceMC / 100);
                        }
                        else
                        {
                            invoiceMC = int.Parse(barcode.Substring(22, 4))/100;
                            DryWeight += int.Parse(barcode.Substring(18, 3)) * (1 - invoiceMC / 100);
                        }
                    }
                    dt.WetTarget = 0;
                    dt.DryTobacco = DryWeight;
                    dt.InvoiceOV = (1 - (dt.Tobacco / dt.ProducedQTY)) * 100;
                    dt.DryYield = (dt.Packing - dt.DryRS) / ((dt.DryTobacco + dt.DryISCRES + dt.DryET + dt.DrySL) + (dt.BC * dt.DMBC / 100) + (dt.BS * dt.DMBS / 100) + (dt.BT * dt.DMBT / 100) + (dt.AC * dt.DMAC / 100)) * 100;
                    
                    queryAle = @"insert into PPReportYieldWhiteWest ([EndTime],[SapID],[BatchIdent],[BlendCode],[ProducedQTY],[Tobacco],[TotalStems],[TotalExpandedTobacco],[TotalSmallLamina],[TotalRipperShorts],[WetYield],[WetTarget],[DryTobacco],[DryISCRES],[DryET],[DrySL],[DryRS],[FinalOV],[InvoiceOV],[DMBS],[DMBT],[DMBC],[DMAC],[BS],[BT],[BC],[AC],[Packing],[DryYield]) 
                        values ('" + dt.EndTime + "','" + dt.SAPID + "','" + dt.BatchIdent + "','" +
                        dt.BlendCode + "'," + dt.ProducedQTY + "," + dt.Tobacco + "," + dt.TotalStems + "," + dt.TotalExpandedTobacco + "," + dt.TotalSmallLamina
                        + "," + dt.TotalRipperShorts + "," + dt.WetYield + "," + dt.WetTarget + "," + dt.DryTobacco
                        + "," + dt.DryISCRES + "," + dt.DryET + "," + dt.DrySL + "," + dt.DryRS + "," + dt.FinalOV + "," + dt.InvoiceOV
                        + "," + dt.DMBS + "," + dt.DMBT + "," + dt.DMBC + ","  + dt.DMAC + "," + dt.BS + "," + dt.BT + "," + dt.BC+"," + dt.AC+ "," + dt.Packing + "," + dt.DryYield + ")";
                    using (SqlConnection con = new SqlConnection(strConString2))
                    {

                        SqlCommand cmd = new SqlCommand(queryAle, con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dset);

                        Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                    }
                }


                //----------------------------------------------

                return Json(new { Status = "True", Data = data }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult GetViewDataWhiteWest(string StartDate, string EndDate)
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

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                string queryAle = "";
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                string startD = dtStartDate.ToString("MM/dd/yyyy");
                string endD = dtEndDate.ToString("MM/dd/yyyy");
                queryAle = @"Select * from PPReportYieldWhiteWest where 
                                [EndTime] between '" + startD + @" 00:00:00.000' AND  '" + endD + @" 23:59:59.000'";
                using (SqlConnection con = new SqlConnection(strConString))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }

                List<PPReportYieldWhiteWestModel> yields = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToPPReportYieldWhiteWestList();
                if (!string.IsNullOrEmpty(searchValue))
                {
                    yields = yields.Where(m => (m.SAPID != null ? m.SAPID.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BatchIdent != null ? m.BatchIdent.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BlendCode != null ? m.BlendCode.ToLower().Contains(searchValue.ToLower()) : false)).ToList();

                }
                int recordsTotal = yields.Count();

                // total number of rows count     
                int recordsFiltered = yields.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = yields.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region View Data Kretek West
        [HttpPost]
        public ActionResult GetDataKretekWest(string StartDate, string EndDate)
        {
            try
            {

                string strConString = ConfigurationManager.ConnectionStrings["AdoPSS4Conn2"].ConnectionString;
                string strConString2 = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                string queryAle = "";
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                string startD = dtStartDate.ToString("MM/dd/yyyy");
                string endD = dtEndDate.ToString("MM/dd/yyyy");
                //int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                //int startWeek = getWeek(dtStartDate);
                //int endWeek = getWeek(dtEndDate);

                //Get Wet Yield
                queryAle = @"SELECT NpssReportBatch.EndTime,NpssReportBatch.SAPID, NpssReportBatch.BatchIdent, NpssReportBatch.BlendCode, NpssReportBatch.ProducedQTY, CAST(NpssReportBatchData.Value As DECIMAL (9,2)) as Tobacco,
                                    NpssReportBatch.TotalStems, NpssReportBatch.TotalExpandedTobacco, NpssReportBatch.TotalSmallLamina,
                                    NpssReportBatch.TotalCloves, NpssReportBatch.TotalOffspec
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000'
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skip1
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skip1:
                                    @"AND NpssReportBatchData.DataName = 'END_Packing_TotalRJKR' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent
                                    
                                    SELECT NpssReportBatch.BatchIdent,CAST (NpssReportBatchData.Value As Decimal (9,0)) as CSF
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe8a
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe8a:
                                    @"AND NpssReportBatchData.DataName = 'Total CSF' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0'AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) AS  DryISCRES
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '" + startD + @"  00:00:00.000'
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe22
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe22:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600125TotDryTob_CRES' AND NpssReportBatchData.Category = 'End_Batch'
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) AS DryRTC
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000'
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe23
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe23:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600130TotDryTob_RCS' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) AS DryET
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000'
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe24
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe24:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600120TotDryTob_ET' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) As FinalOV
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between  '"+ startD + @" 00:00:00.000' 
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe16
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe16:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_MM600230Average_Final' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) As DryOfSpec
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between  '"+ startD + @" 00:00:00.000' 
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe28
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe28:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600160TotDryTob_OS' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) DryCLOVE
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between  '"+ startD + @" 00:00:00.000' 
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe25
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe25:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600150TotDryTob_CLOVE' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent, CAST (NpssReportBatchData.Value As Decimal (9,2)) AS DryCSF
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between  '"+ startD + @" 00:00:00.000' 
                                    AND '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe26
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe26:
                                    @"AND NpssReportBatchData.DataName = 'END_AB_WB600155TotDryTob_CSF' AND NpssReportBatchData.Category = 'End_Batch' 
                                    AND NpssReportBatch.BatchIdent <> '0' AND NpssReportBatch.LastEventName like '%Pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    ORDER BY NpssReportBatch.BatchIdent 

                                    SELECT NpssReportBatch.BatchIdent,NpssReportBatch.TotalBrightCasing,NpssReportBatch.TotalBurleySpray, NpssReportBatch.TotalAfterCut
                                    FROM NpssReportBatch INNER JOIN
                                    NpssReportBatchData ON NpssReportBatch.BatchID = NpssReportBatchData.EventID
                                    WHERE NpssReportBatch.EndTime between '"+ startD + @" 00:00:00.000' 
                                    AND  '" + endD + @" 23:59:59.000'" +
                                    //--If range(C5).Value = ALL Then GoTo Skipe15
                                    //--AND NpssReportBatch.BlendCode =  & ' & range(C5).Text & ' 
                                    //--Skipe15:
                                    @"AND NpssReportBatchData.DataName = 'Batch ID' 
                                    AND NpssReportBatchData.EventType =  'END AddBack' 
                                    AND NpssReportBatch.LastEventName like '%pack%'
                                    AND NpssReportBatch.LastEventName <> 'Lamina ST Pack'
                                    AND NpssReportBatch.BatchIdent <> '0' 
                                    ORDER BY NpssReportBatch.BatchIdent ";
                using (SqlConnection con = new SqlConnection(strConString))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }
                List<PPReportYieldKretekWestModel> data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToPPReportYieldKretekWestList();
                //CSF
                List<PPReportYieldKretekWestModel> data2 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[1]).DeserializeToPPReportYieldKretekWestList();
                //ISCRES
                List<PPReportYieldKretekWestModel> data3 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[2]).DeserializeToPPReportYieldKretekWestList();
                //RTC
                List<PPReportYieldKretekWestModel> data4 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[3]).DeserializeToPPReportYieldKretekWestList();
                //ET
                List<PPReportYieldKretekWestModel> data5 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[4]).DeserializeToPPReportYieldKretekWestList();
                //FinalOV
                List<PPReportYieldKretekWestModel> data6 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[5]).DeserializeToPPReportYieldKretekWestList();
                //OffSpec
                List<PPReportYieldKretekWestModel> data7 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[6]).DeserializeToPPReportYieldKretekWestList();
                //Clove
                List<PPReportYieldKretekWestModel> data8 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[7]).DeserializeToPPReportYieldKretekWestList();
                //CSF Dry
                List<PPReportYieldKretekWestModel> data9 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[8]).DeserializeToPPReportYieldKretekWestList();
                //BC BS AC 
                List<PPReportYieldKretekWestModel> data10 = DataTableToJSONWithJavaScriptSerializer(dset.Tables[9]).DeserializeToPPReportYieldKretekWestList();
                queryAle = @"delete from PPReportYieldKretekWest where 
                                [EndTime] between '" + startD + @" 00:00:00.000' AND  '" + endD + @" 23:59:59.000'";
                using (SqlConnection con = new SqlConnection(strConString2))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }
                foreach (PPReportYieldKretekWestModel dt in data)
                {
                    dt.CSF = data2.Count > 0 ? data2.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().CSF : 0;
                    dt.DryISCRES = data3.Count > 0 ? data3.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryISCRES : 0;
                    dt.DryRTC = data4.Count > 0 ? data4.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryRTC : 0;
                    dt.DryET = data5.Count > 0 ? data5.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryET : 0;
                    dt.FinalOV = data6.Count > 0 ? data6.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().FinalOV : 0;
                    dt.DryOfSpec = data7.Count > 0 ? data7.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryOfSpec : 0;
                    dt.DryCLOVE = data8.Count > 0 ? data8.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryCLOVE : 0;
                    dt.DryCSF = data9.Count > 0 ? data9.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().DryCSF : 0;
                    dt.TotalBrightCasing = data10.Count > 0 ? data10.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().TotalBrightCasing : 0;
                    dt.TotalBurleySpray = data10.Count > 0 ? data10.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().TotalBurleySpray : 0;
                    dt.TotalAfterCut = data10.Count > 0 ? data10.Where(x => x.BatchIdent == dt.BatchIdent).FirstOrDefault().TotalAfterCut : 0;

                    //Define
                    switch (dt.BlendCode)
                    {
                        case "D0A1T":
                            dt.WetTarget = 106.63;
                            dt.InvoiceOV = 12.86;
                            dt.DMBC = 75.6;
                            dt.DMAC = 59.9;
                            break;
                        case "V0CGJ":
                            dt.WetTarget = 106.46;
                            dt.InvoiceOV = 12.93;
                            dt.DMBC = 77.3;
                            dt.DMAC = 80.2;
                            break;
                        case "V0N5A":
                            dt.WetTarget = 106.52;
                            dt.InvoiceOV = 12.67;
                            dt.DMBC = 77.3;
                            dt.DMAC = 80.2;
                            break;
                        case "V0C3Z":
                            dt.WetTarget = 109.61;
                            dt.InvoiceOV = 13.1;
                            dt.DMBC = 62.6;
                            dt.DMAC = 75.8;
                            break;
                        case "V0Q3W":
                            dt.WetTarget = 103.92;
                            dt.InvoiceOV = 12.99;
                            dt.DMBC = 75.6;
                            dt.DMAC = 59.9;
                            break;
                        case "V0CBS":
                            dt.WetTarget = 108.21;
                            dt.InvoiceOV = 13.01;
                            dt.DMBC = 62.7;
                            dt.DMAC = 71.8;
                            break;
                        case "V0CL7":
                            dt.WetTarget = 106.8;
                            dt.InvoiceOV = 12.47;
                            dt.DMBC = 57.1;
                            dt.DMAC = 73.1;
                            break;
                        case "V0CL4":
                            dt.WetTarget = 103.92;
                            dt.InvoiceOV = 12.99;
                            dt.DMBC = 75.6;
                            dt.DMAC = 59.9;
                            break;
                        case "ZJ0CP":
                            dt.WetTarget = 106.2;
                            dt.InvoiceOV = 12.5;
                            dt.DMBC = 77.3;
                            dt.DMAC = 80;
                            break;
                    }

                    //Calculation
                    dt.WetYield = ((dt.ProducedQTY - dt.TotalOffspec) / (dt.Tobacco + dt.TotalStems + dt.TotalExpandedTobacco + dt.TotalSmallLamina + dt.TotalCloves + dt.CSF)) * 100;
                    dt.DryTobacco = dt.Tobacco * (1 - dt.InvoiceOV / 100);
                    dt.Packing = dt.ProducedQTY * (1 - dt.FinalOV / 100);
                    dt.DryCasing = (dt.TotalBrightCasing + dt.TotalBurleySpray) * dt.DMBC / 100;
                    dt.DryAC = dt.TotalAfterCut * dt.DMAC / 100;
                    dt.DryYield = (dt.Packing - dt.DryOfSpec) / (dt.DryTobacco + dt.DryISCRES + dt.DryRTC + dt.DryET + dt.DryCLOVE + dt.DryCSF + dt.DryCasing + dt.DryAC) * 100;

                    queryAle = @"insert into PPReportYieldKretekWest ([EndTime],[SapID],[BatchIdent],[BlendCode],[ProducedQTY],[Tobacco],[TotalStems],[TotalExpandedTobacco],[TotalSmallLamina],[TotalCloves],[TotalOffspec],[CSF]
                        ,[WetYield],[WetTarget],[DryTobacco],[DryISCRES],[DryRTC],[DryET],[DryCLOVE],[DryCSF],[DryOfSpec],[FinalOV],[InvoiceOV],[DMBC],[TotalBrightCasing],[TotalBurleySpray],[TotalAfterCut],[Packing],[DryYield],[DMAC],[DryCasing],[DryAC]) 
                        values ('" + dt.EndTime + "','" + dt.SAPID + "','" + dt.BatchIdent + "','" +
                        dt.BlendCode + "'," + dt.ProducedQTY + "," + dt.Tobacco + "," + dt.TotalStems + "," + dt.TotalExpandedTobacco + "," + dt.TotalSmallLamina
                        + "," + dt.TotalCloves + "," + dt.TotalOffspec + "," + dt.CSF + "," + dt.WetYield + "," + dt.WetTarget + "," + dt.DryTobacco
                        + "," + dt.DryISCRES + "," + dt.DryRTC + "," + dt.DryET + "," + dt.DryCLOVE + "," + dt.DryCSF + "," + dt.DryOfSpec + "," + dt.FinalOV
                        + "," + dt.InvoiceOV + "," + dt.DMBC + "," + dt.TotalBrightCasing + "," + dt.TotalBurleySpray + "," + dt.TotalAfterCut + "," + dt.Packing + "," + dt.DryYield
                        + "," + dt.DMAC + "," + dt.DryCasing + "," + dt.DryAC + ")";
                    using (SqlConnection con = new SqlConnection(strConString2))
                    {

                        SqlCommand cmd = new SqlCommand(queryAle, con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dset);

                        Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                    }
                }


                //----------------------------------------------

                return Json(new { Status = "True", Data = data }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult GetViewDataKretekWest(string StartDate, string EndDate)
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

                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                DataSet dset = new DataSet();
                string queryAle = "";
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                string startD = dtStartDate.ToString("MM/dd/yyyy");
                string endD = dtEndDate.ToString("MM/dd/yyyy");
                queryAle = @"Select * from PPReportYieldKretekWest where 
                                [EndTime] between '" + startD + @" 00:00:00.000' AND  '" + endD + @" 23:59:59.000'";
                using (SqlConnection con = new SqlConnection(strConString))
                {

                    SqlCommand cmd = new SqlCommand(queryAle, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dset);

                    Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
                }

                List<PPReportYieldKretekWestModel> yields = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]).DeserializeToPPReportYieldKretekWestList();
                if (!string.IsNullOrEmpty(searchValue))
                {
                    yields = yields.Where(m => (m.SAPID != null ? m.SAPID.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BatchIdent != null ? m.BatchIdent.ToLower().Contains(searchValue.ToLower()) : false) ||
                                            (m.BlendCode != null ? m.BlendCode.ToLower().Contains(searchValue.ToLower()) : false)).ToList();

                }
                int recordsTotal = yields.Count();

                // total number of rows count     
                int recordsFiltered = yields.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = yields.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region View Data Diet
        public ActionResult GetDataMCDiet(string StartDate, string EndDate)
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
                /*
                // Getting all data submissions   			
                string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
                List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();
                */
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                int startWeek = getWeek(dtStartDate);
                int endWeek = getWeek(dtEndDate);

                List<QueryFilter> filterMCs = new List<QueryFilter>();
                filterMCs.Add(new QueryFilter("Year", dtStartDate.Year.ToString(), Operator.GreaterThanOrEqual));
                filterMCs.Add(new QueryFilter("Year", dtEndDate.Year.ToString(), Operator.LessThanOrEqualTo, Operation.AndAlso));
                //filterMCs.Add(new QueryFilter("Week", startWeek.ToString(), Operator.GreaterThanOrEqual,Operation.Or));
                filterMCs.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                filterMCs.Add(new QueryFilter("IsDeleted", "0"));
                List<PPReportYieldMCDietModel> prymdm = _pPReportYieldMCDietsAppService.Find(filterMCs).DeserializeToPPReportYieldMCDietList();
                prymdm = prymdm.Where(x => x.Week >= startWeek && x.Week <= endWeek).ToList();

                int recordsTotal = prymdm.Count();
                
                // total number of rows count     
                int recordsFiltered = prymdm.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = prymdm.Skip(skip).Take(pageSize).ToList();
                
                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
                
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult GetDataDiet(string StartDate, string EndDate)
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
                /*
                // Getting all data submissions   			
                string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
                List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();
                */
                DateTime dtStartDate = DateTime.ParseExact(StartDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
                DateTime dtEndDate = DateTime.ParseExact(EndDate, "dd-MMM-yy", CultureInfo.InvariantCulture);

                int getWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                int startWeek = getWeek(dtStartDate);
                int endWeek = getWeek(dtEndDate);

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                filters.Add(new QueryFilter("Week", startWeek.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Week", endWeek.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountDepartmentID, "productioncenter");


                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("LPHHeader", "LPHPrimaryDietController"));
                filters.Add(new QueryFilter("Date", dtStartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("Date", dtEndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<PPLPHSubmissionsModel> submissions = _ppLphSubmissionsAppService.Find(filters).DeserializeToPPLPHSubmissionsList();

                var minSubmissionID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                var maxSubmissionID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("LPHSubmissionID", minSubmissionID.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("LPHSubmissionID", maxSubmissionID.ToString(), Operator.LessThanOrEqualTo));

                List<PPLPHApprovalsModel> approval = _ppLphApprovalAppService.Find(filters).DeserializeToPPLPHApprovalList();
                submissions = submissions.Where(l =>
                { 
                    return approval.Where(x => x.LPHSubmissionID == l.ID && x.Status.Trim().ToLower() == "approved").Count() > 0;
                }).ToList();

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("HeaderName", "waste"));
                filters.Add(new QueryFilter("FieldName", "WasteNoSKJ"));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                List<LPHExtrasModel> OpsResultList = _ppLphExtrasAppService.Find(filters).DeserializeToLPHExtraList()
                    .Where(x =>
                        submissions.Any(z => z.LPHID == x.LPHID)
                    ).ToList();

                List<PPReportYieldDietModel> ResultYield = new List<PPReportYieldDietModel>();
                PPReportYieldMCDietModel prymdm = new PPReportYieldMCDietModel();
                foreach (PPLPHSubmissionsModel plsm in submissions)
                {
                    List<LPHExtrasModel> OpsThisSubmission = OpsResultList.Where(x => x.LPHID == plsm.LPHID).ToList();
                    foreach (LPHExtrasModel lem in OpsThisSubmission)
                    {
                        PPReportYieldDietModel prydm = new PPReportYieldDietModel();
                        prydm.Year = plsm.Date.Year;
                        prydm.Week = getWeek(dtStartDate);
                        prydm.Date = plsm.Date;

                        //Get Table Waste Per Row
                        filters = new List<QueryFilter>();
                        filters.Add(new QueryFilter("LPHID", plsm.LPHID));
                        filters.Add(new QueryFilter("HeaderName", "waste"));
                        filters.Add(new QueryFilter("IsDeleted", "0"));

                        List<LPHExtrasModel> wasteThisSubmission = _ppLphExtrasAppService.Find(filters).DeserializeToLPHExtraList();

                        prydm.Blend = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteBlend").FirstOrDefault().Value;
                        string Ops = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteNoSKJ").FirstOrDefault().Value;
                        prydm.Ops = Ops != null ? int.Parse(Ops) : 0;

                        //Get Casing Data

                        //Get Table Waste Per Row
                        string ConditioningOPS = "1" + Ops.Substring(2);
                        List<QueryFilter> filterCasings = new List<QueryFilter>();
                        filterCasings.Add(new QueryFilter("HeaderName", "krosok"));
                        filterCasings.Add(new QueryFilter("FieldName", "krosokBatchNo"));
                        filterCasings.Add(new QueryFilter("Value", ConditioningOPS, Operator.Contains));
                        filterCasings.Add(new QueryFilter("IsDeleted", "0"));

                        LPHExtrasModel casingThisOps = _ppLphExtrasAppService.Find(filterCasings).DeserializeToLPHExtraList().FirstOrDefault();
                        prydm.Casing = 0;
                        if (casingThisOps != null) {
                            List<QueryFilter> filterCasingVals = new List<QueryFilter>();
                            filterCasingVals.Add(new QueryFilter("HeaderName", "krosok"));
                            filterCasingVals.Add(new QueryFilter("LPHID", casingThisOps.LPHID.ToString()));
                            filterCasingVals.Add(new QueryFilter("RowNumber", casingThisOps.RowNumber.ToString()));
                            filterCasingVals.Add(new QueryFilter("IsDeleted", "0"));
                            List<LPHExtrasModel> casingVal = _ppLphExtrasAppService.Find(filterCasingVals).DeserializeToLPHExtraList();

                            string Casing = casingVal.Where(x => x.FieldName == "krosokTotalizer").FirstOrDefault().Value;
                            prydm.Casing = Casing != null ? double.Parse(Casing) : 0;
                        }
                        string CVIB0069 =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCVIB0069").FirstOrDefault().Value;
                        prydm.CVIB0069 = CVIB0069 != null ? double.Parse(CVIB0069) : 0;

                        string CSFR0022 =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCSFR0022").FirstOrDefault().Value;
                        prydm.CSFR0022 = CSFR0022 != null ? double.Parse(CSFR0022) : 0;

                        string DSCL0034 =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteDSCL0034").FirstOrDefault().Value;
                        prydm.DSCL0034 = DSCL0034 != null ? double.Parse(DSCL0034) : 0;

                        string CVIB0070 =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCVIB0070").FirstOrDefault().Value;
                        prydm.CVIB0070 = CVIB0070 != null ? double.Parse(CVIB0070) : 0;

                        string RV0054 =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteRV0054").FirstOrDefault().Value;
                        prydm.RV0054 = RV0054 != null ? double.Parse(RV0054) : 0;

                        string RM =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteRM").FirstOrDefault().Value;
                        string Addback =  wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteAddback").FirstOrDefault().Value;
                        prydm.Input = (RM != null ? double.Parse(RM) : 0) + (Addback != null ? double.Parse(Addback) : 0);

                        string Output = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WastePackingKg").FirstOrDefault().Value;
                        prydm.Output = Output != null ? double.Parse(Output) : 0;

                        string CGNSolar = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteTotalGAS").FirstOrDefault().Value;
                        prydm.CGNSolar = CGNSolar != null ? double.Parse(CGNSolar) : 0;

                        string STEAM = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteTotalSTEAM").FirstOrDefault().Value;
                        prydm.STEAM = STEAM != null ? double.Parse(STEAM) : 0;

                        string Awal = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCO2Awal").FirstOrDefault().Value;
                        prydm.Awal = Awal != null ? double.Parse(Awal) : 0;

                        string Terima = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCO2Terima").FirstOrDefault().Value;
                        prydm.Terima = Terima != null ? double.Parse(Terima) : 0;

                        string Akhir = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCO2Akhir").FirstOrDefault().Value;
                        prydm.Akhir = Akhir != null ? double.Parse(Akhir) : 0;

                        string Pakai = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteCO2Pakai").FirstOrDefault().Value;
                        prydm.Pakai = Pakai != null ? double.Parse(Pakai) : 0;

                        string RateCO2 = wasteThisSubmission.Where(x => x.LPHID == plsm.LPHID && x.RowNumber == lem.RowNumber && x.FieldName == "WasteConsumptionCO2").FirstOrDefault().Value;
                        prydm.RateCO2 = RateCO2 != null ? double.Parse(RateCO2) : 0;

                        prydm.WetYield = prydm.Output / prydm.Input;

                        List<QueryFilter>  filterMCs = new List<QueryFilter>();
                        filterMCs.Add(new QueryFilter("Year", prydm.Year));
                        filterMCs.Add(new QueryFilter("Week", prydm.Week));
                        filterMCs.Add(new QueryFilter("LocationID", AccountDepartmentID.ToString()));
                        filterMCs.Add(new QueryFilter("IsDeleted", "0"));
                        prymdm = _pPReportYieldMCDietsAppService.Find(filterMCs).DeserializeToPPReportYieldMCDietList().FirstOrDefault();
                        if(prymdm != null)
                        {
                            prydm.DryInput = ((prydm.Input - prymdm.Flake) * (1 - prymdm.MCKrosok)) + (prymdm.Flake * (1 - prymdm.MCFlake));
                            prydm.DryCasing = prydm.Casing * prymdm.DM;

                            prydm.DryYield = (prydm.Output * (1 - prymdm.MCPacking)) / (prydm.DryInput + prydm.DryCasing);
                            prydm.DryWaste = (prydm.CVIB0069 * (1 - prymdm.CVIB0069)) / (prydm.DryInput + prydm.DryCasing);
                            prydm.WetWaste = (prydm.CSFR0022 * (1 - prymdm.CSFR0022)) / (prydm.DryInput + prydm.DryCasing);
                            prydm.DustWaste = ((prydm.DSCL0034 * (1 - prymdm.DSCL0034))+ (prydm.CVIB0070 * (1 - prymdm.CVIB0070))) / (prydm.DryInput + prydm.DryCasing);
                            prydm.HotDustWaste = (prydm.RV0054 * (1 - prymdm.RV0054)) / (prydm.DryInput + prydm.DryCasing);

                        }
                        ResultYield.Add(prydm);
                    }
                }
                if (ResultYield.Count > 0)
                {
                    PPReportYieldDietModel sumYield = new PPReportYieldDietModel();
                    double? sumOutput = ResultYield.Sum(x => x.Output);
                    double? sumInput = ResultYield.Sum(x => x.Input);
                    double? sumCasing = ResultYield.Sum(x => x.Casing);
                    sumYield.WetYield = sumOutput / sumInput;
                    if (prymdm != null)
                    {
                        double? sumDryInput = ((sumInput - prymdm.Flake) * (1 - prymdm.MCKrosok)) + (prymdm.Flake * (1 - prymdm.MCFlake));
                        double? sumDryCasing = sumCasing * prymdm.DM;

                        sumYield.DryYield = (sumOutput * (1 - prymdm.MCPacking)) / (sumDryInput + sumDryCasing);
                    }

                    ResultYield.Insert(0, sumYield);
                }
                int recordsTotal = ResultYield.Count();
                
                // total number of rows count     
                int recordsFiltered = ResultYield.Count();

                // Paging     
                //var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
                var data = ResultYield.Skip(skip).Take(pageSize).ToList();
                
                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
                
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<PPReportYieldWhiteModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region Function Helper
        public int GetWeeksInYear(int year)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = new DateTime(year, 12, 31);
            Calendar cal = dfi.Calendar;
            return cal.GetWeekOfYear(date1, dfi.CalendarWeekRule,
                                                dfi.FirstDayOfWeek);
        }
        #endregion
    }
}
