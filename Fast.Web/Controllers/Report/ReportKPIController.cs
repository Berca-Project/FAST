using ExcelDataReader;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Fast.Web.Models;
using Fast.Web.Models.Report;
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Utils;
using System.Globalization;
using OfficeOpenXml;
using Fast.Web.Resources;
using OfficeOpenXml.Style;
using System.Configuration;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Fast.Web.Models.LPH;

namespace Fast.Web.Controllers.Report
{
    public static class DateTimeExtensions
    {
        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = (7 + (dt.DayOfWeek - startOfWeek)) % 7;
            return dt.AddDays(-1 * diff).Date;
        }
    }


    public class ReportKPIController : BaseController<UserModel>
	{
        private static string HEADER_COLOR = "#99ccff";
        private readonly IUserAppService _userAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _empService;
		private readonly IUserLogAppService _userLogAppService;
		private readonly IReportKPICRRAppService _reportKPICRRAppService;
		private readonly IReportKPIDIMAppService _reportKPIDIMAppService;
		private readonly IReportKPIYieldAppService _reportKPIYieldAppService;
		private readonly IReportKPIProdVolAppService _reportKPIProdVolAppService;
		private readonly IReportKPIWorkHourAppService _reportKPIWorkHourAppService;
		private readonly IReportKPIStickPerPackAppService _reportKPIStickPerPackAppService;
		private readonly IReportKPIDustAppService _reportKPIDustAppService;
		private readonly IReportKPITobaccoWeightAppService _reportKPITobaccoWeightAppService;
		private readonly IReportKPIRipperInfoAppService _reportKPIRipperInfoAppService;
		private readonly IReportKPICRRConversionAppService _reportKPICRRConversionAppService;
		private readonly IReportKPITargetAppService _reportKPITargetAppService;
        private readonly IInputTargetAppService _inputTargetAppService;
        private readonly ILocationAppService _locationAppService;

        public ReportKPIController(IUserAppService userAppService,
			ILoggerAppService logger,
			IUserLogAppService userLogAppService,
			IEmployeeAppService empService,
			IReportKPICRRAppService reportKPICRRService,
			IReportKPIDIMAppService reportKPIDIMService,
			IReportKPIYieldAppService reportKPIYieldService,
			IReportKPIProdVolAppService reportKPIProdVolService,
			IReportKPIWorkHourAppService reportKPIWorkHourService,
			IReportKPIStickPerPackAppService reportKPIStickPerPackService,
			IReportKPIDustAppService reportKPIDustService,
			IReportKPITobaccoWeightAppService reportKPITobaccoWeightService,
			IReportKPIRipperInfoAppService reportKPIRipperInfoService,
            IReportKPICRRConversionAppService reportKPICRRConversionService,
            IReportKPITargetAppService reportKPITargetService,
            ILocationAppService locationAppService,
            IInputTargetAppService inputTargetAppService,
            IReferenceAppService referenceAppService)
		{
			_referenceAppService = referenceAppService;
			_userAppService = userAppService;
			_logger = logger;
			_userLogAppService = userLogAppService;
			_empService = empService;
			_reportKPICRRAppService = reportKPICRRService;
			_reportKPIDIMAppService = reportKPIDIMService;
			_reportKPIYieldAppService = reportKPIYieldService;
			_reportKPIProdVolAppService = reportKPIProdVolService;
			_reportKPIWorkHourAppService = reportKPIWorkHourService;
			_reportKPIStickPerPackAppService = reportKPIStickPerPackService;
			_reportKPIDustAppService = reportKPIDustService;
			_reportKPITobaccoWeightAppService = reportKPITobaccoWeightService;
			_reportKPIRipperInfoAppService = reportKPIRipperInfoService;
			_reportKPITargetAppService = reportKPITargetService;
            _locationAppService = locationAppService;
            _inputTargetAppService = inputTargetAppService;
            _reportKPICRRConversionAppService = reportKPICRRConversionService;
		}


		// GET: ShiftDaily
		public ActionResult Index(string page="")
		{
            if (Session["ResultLog"] != null)
			{
				ViewBag.ResultLog = Session["ResultLog"].ToString();
				Session["ResultLog"] = null;
			}

            ViewBag.Main = page;

            return View();
		}

        public ActionResult ReportKPI(string PC = "", int year = 0, int week = 1, int showall = 0)
        {
            var model = new ReportWeeklyKPIModel();
            try
            {
                if (year == 0)
                    year = DateTime.Now.Year;

                ViewBag.PC = PC;
                ViewBag.year = year;
                ViewBag.week = week;
                ViewBag.showall = showall;

                var dataProdVol = _reportKPIProdVolAppService.GetAll(true).DeserializeToReportKPIProdVolModelList().Where(x=>x.Year == year).ToList();
                var dataDIM = _reportKPIDIMAppService.GetAll(true).DeserializeToReportKPIDIMModelList().Where(x => x.Year == year).ToList();
                var dataYield = _reportKPIYieldAppService.GetAll(true).DeserializeToReportKPIYieldModelList().Where(x => x.Year == year).ToList();
                var dataCRR = _reportKPICRRAppService.GetAll(true).DeserializeToReportKPICRRModelList().Where(x => x.Year == year).ToList();
                var dataWorkHour = _reportKPIWorkHourAppService.GetAll(true).DeserializeToReportKPIWorkHourModelList().Where(x => x.Year == year).ToList();
                var dataStickPerPack = _reportKPIStickPerPackAppService.GetAll(true).DeserializeToReportKPIStickPerPackModelList();
                var dataInputTarget = _inputTargetAppService.GetAll(true).DeserializeToInputTargetList();
                var dataCRRConversion = _reportKPICRRConversionAppService.GetAll(true).DeserializeToReportKPICRRConversionModelList().Where(x => x.Year == year).ToList();
                var dataTarget = _reportKPITargetAppService.GetAll(true).DeserializeToReportKPITargetModelList();

                var dataDust = _reportKPIDustAppService.GetAll(true).DeserializeToReportKPIDustModelList();
                var dataTobaccoWeight = _reportKPITobaccoWeightAppService.GetAll(true).DeserializeToReportKPITobaccoWeightModelList().Where(x => x.Year == year).ToList();
                var dataRipperInfo = _reportKPIRipperInfoAppService.GetAll(true).DeserializeToReportKPIRipperInfoModelList().Where(x => x.Year == year).ToList();

                if (PC != "")
                {
                    dataProdVol = dataProdVol.Where(x => x.ProductionCenter == PC).ToList();
                    dataDIM = dataDIM.Where(x => x.ProductionCenter == PC).ToList();
                    dataYield = dataYield.Where(x => x.ProductionCenter == PC).ToList();
                    dataCRR = dataCRR.Where(x => x.ProductionCenter == PC).ToList();
                    dataWorkHour = dataWorkHour.Where(x => x.ProductionCenter == PC).ToList();
                    dataStickPerPack = dataStickPerPack.Where(x => x.ProductionCenter == PC).ToList();
                    dataCRRConversion = dataCRRConversion.Where(x => x.ProductionCenter == PC).ToList();

                    dataTarget = dataTarget.Where(x => x.ProductionCenter == PC).ToList();

                    dataDust = dataDust.Where(x => x.ProductionCenter == PC).ToList();
                    dataTobaccoWeight = dataTobaccoWeight.Where(x => x.ProductionCenter == PC).ToList();
                    dataRipperInfo = dataRipperInfo.Where(x => x.ProductionCenter == PC).ToList();

                    var location = _locationAppService.FindBy("ParentID", 1, true).DeserializeToLocationList().Where(x => x.Code == PC).FirstOrDefault();
                    dataInputTarget = dataInputTarget.Where(x => x.ProdCenterID == location.ID).ToList();
                } else
                {
                    dataTarget = dataTarget.Where(x => x.ProductionCenter == "ID").ToList();
                }

                //********************************************************** TARGET & OB tak kumpulkan biar gampang ngisinya
                model.Header = new List<string>() { "OB20", "Internal Target" };

                model.Volume = new List<decimal>() { 0, 0 };

                model.Uptime = new List<decimal>() { 0, 0 };
                model.CRR = new List<decimal>() { 0, 0 };
                model.ActYield = new List<decimal>() { 0, 0 };
                model.DIMWaste = new List<decimal>() { 0, 0 };
                model.Claimable = new List<decimal>() { 0, 0 };
                model.TaxStampWaste = new List<decimal>() { 0, 0 };
                model.ClaimableInThStick = new List<decimal>() { 0,0 };

                model.FilterRod = new List<decimal>() { 0, 0 };
                model.HingeLidBlank = new List<decimal>() { 0, 0 };
                model.InnerLiner = new List<decimal>() { 0, 0 };
                model.TippingPaper = new List<decimal>() { 0, 0 };
                model.WrappingFilm = new List<decimal>() { 0, 0 };
                model.InnerFrame = new List<decimal>() { 0, 0 };
                model.CigarettePaper = new List<decimal>() { 0, 0 };
                
                var GICutfillerWet = new List<decimal>();
                var GRCutfillerActualWet = new List<decimal>();
                var WIPActualWet = new List<decimal>();
                var GRRSWet = new List<decimal>();

                if (dataTarget != null && dataTarget.Count() > 0)
                {
                    model.Volume[0] = dataTarget.Where(x => x.KPI == "Volume").Select(x=>x.TargetOB).DefaultIfEmpty(0).First();
                    model.Volume[1] = dataTarget.Where(x => x.KPI == "Volume").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.Uptime[0] = dataTarget.Where(x => x.KPI == "Uptime").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.Uptime[1] = dataTarget.Where(x => x.KPI == "Uptime").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.CRR[0] = dataTarget.Where(x => x.KPI == "CRR").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.CRR[1] = dataTarget.Where(x => x.KPI == "CRR").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.ActYield[0] = dataTarget.Where(x => x.KPI == "ActYield").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.ActYield[1] = dataTarget.Where(x => x.KPI == "ActYield").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.DIMWaste[0] = dataTarget.Where(x => x.KPI == "DIMWaste").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.DIMWaste[1] = dataTarget.Where(x => x.KPI == "DIMWaste").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.Claimable[0] = dataTarget.Where(x => x.KPI == "Claimable").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.Claimable[1] = dataTarget.Where(x => x.KPI == "Claimable").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.TaxStampWaste[0] = dataTarget.Where(x => x.KPI == "TaxStampWaste").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.TaxStampWaste[1] = dataTarget.Where(x => x.KPI == "TaxStampWaste").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.FilterRod[0] = dataTarget.Where(x => x.KPI == "FilterRod").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.FilterRod[1] = dataTarget.Where(x => x.KPI == "FilterRod").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.HingeLidBlank[0] = dataTarget.Where(x => x.KPI == "HingeLidBlank").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.HingeLidBlank[1] = dataTarget.Where(x => x.KPI == "HingeLidBlank").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.InnerLiner[0] = dataTarget.Where(x => x.KPI == "InnerLiner").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.InnerLiner[1] = dataTarget.Where(x => x.KPI == "InnerLiner").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.TippingPaper[0] = dataTarget.Where(x => x.KPI == "TippingPaper").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.TippingPaper[1] = dataTarget.Where(x => x.KPI == "TippingPaper").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.WrappingFilm[0] = dataTarget.Where(x => x.KPI == "WrappingFilm").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.WrappingFilm[1] = dataTarget.Where(x => x.KPI == "WrappingFilm").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.InnerFrame[0] = dataTarget.Where(x => x.KPI == "InnerFrame").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.InnerFrame[1] = dataTarget.Where(x => x.KPI == "InnerFrame").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();

                    model.CigarettePaper[0] = dataTarget.Where(x => x.KPI == "CigarettePaper").Select(x => x.TargetOB).DefaultIfEmpty(0).First();
                    model.CigarettePaper[1] = dataTarget.Where(x => x.KPI == "CigarettePaper").Select(x => x.TargetInternal).DefaultIfEmpty(0).First();
                }

                if (dataInputTarget != null && dataInputTarget.Count() > 0)
                {
                    var month = DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture);

                    dataInputTarget = dataInputTarget.Where(x => x.Month.ToLower() == month.ToLower() && x.Value > 0 && (x.Version == "Internal Target" || x.Version == "OB")).ToList();
                    if (dataInputTarget != null && dataInputTarget.Count() > 0)
                    {
                        var dataTarget_OB20 = dataInputTarget.Where(x => x.Version == "OB").ToList();
                        if (dataTarget_OB20 != null && dataTarget_OB20.Count() > 0)
                        {
                            if (model.Uptime[0] == 0)
                            {
                                var uptime = dataTarget_OB20.Where(x => x.KPI == "Uptime").FirstOrDefault();
                                if (uptime != null && uptime.Value != 0)
                                    model.Uptime[0] = uptime.Value;
                            }

                            if (model.CRR[0] == 0)
                            {
                                var CRR = dataTarget_OB20.Where(x => x.KPI == "CRR").FirstOrDefault();
                                if (CRR != null && CRR.Value != 0)
                                    model.CRR[0] = CRR.Value;
                            }
                        }

                        var dataTarget_Internal = dataInputTarget.Where(x => x.Version == "Internal Target").ToList();
                        if (dataTarget_Internal != null && dataTarget_Internal.Count() > 0)
                        {
                            if (model.Uptime[1] == 0)
                            {
                                var uptime = dataTarget_Internal.Where(x => x.KPI == "Uptime").FirstOrDefault();
                                if (uptime != null && uptime.Value != 0)
                                    model.Uptime[1] = uptime.Value;
                            }

                            if (model.CRR[1] == 0)
                            {
                                var CRR = dataTarget_Internal.Where(x => x.KPI == "CRR").FirstOrDefault();
                                if (CRR != null && CRR.Value != 0)
                                    model.CRR[1] = CRR.Value;
                            }
                        }
                    }
                }

                var totalWeek = GetWeekNumber(DateTime.Now);
                var startweek = week;

                for (var i = startweek; i < totalWeek; i++)
                {
                    model.Header.Add("W" + i.ToString("00"));

                    var tempYield = dataYield.Where(x => x.Week == i).ToList();
                    var tempRipper = dataRipperInfo.Where(x => x.Week == i).ToList();
                    var tempTobacco = dataTobaccoWeight.Where(x => x.Week == i).ToList();
                    var tempDust = dataDust;
                    var tempCRR = dataCRRConversion.Where(x => x.Week == i).ToList();

                    GICutfillerWet.Add(tempYield.Sum(x => x.TotTKGGI));

                    if (tempTobacco != null && tempTobacco.Count() > 0)
                    {
                        GRCutfillerActualWet.Add(tempYield.Where(x => x.MaterialGroupDesc.ToLower().Contains("cigarette var")).Sum(x => x.GoodsReceiptQty * (tempTobacco.FirstOrDefault().Mean / 1000000000000000)));
                    }
                    else
                    {
                        GRCutfillerActualWet.Add(0);
                    }

                    WIPActualWet.Add(tempYield.Where(x => x.MaterialGroupDesc.ToLower().Contains("reject cig")).Sum(x => x.TotTKGWIP));
                    GRRSWet.Add(tempRipper.Select(x => x.QtyRec).Sum());
                }
                model.Header.Add("YTD");

                //******************************************************************************** perhitungan dimulai dari sini. cemunguudh
                //********************************************************************* variabel yang akan sering dipakai, disendirikan saja
                var sumProdVol = new List<decimal>();
                for (var i = 0; i < startweek; i++)
                {
                    sumProdVol.Add(0);
                }
                for (var i = startweek; i < totalWeek; i++)
                {
                    var ProdVol = dataProdVol.Where(x => x.Week == i).Sum(x => x.ConfirmedQuantity);
                    sumProdVol.Add(ProdVol);
                }

                //******************************************************************************************************************* VOLUME
                
                decimal tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    tempTotal += sumProdVol[i];
                    model.Volume.Add(sumProdVol[i]);
                }
                model.Volume.Add(tempTotal);

                //******************************************************************************************************************* UPTIME
               
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    var workhours = dataWorkHour.Where(x => x.Week == i).ToList();

                    if (workhours != null && workhours.Count() > 0)
                    {
                        foreach(var brand in workhours)
                        {
                            var kali = dataStickPerPack.Where(x => x.Packer == brand.Packer).FirstOrDefault();

                            if (kali != null & kali.Stick > 0)
                                temp += brand.WorkHour * kali.Stick * 60;
                        }
                    }

                    if (temp > 0)
                    {
                        //tempTotal += sumProdVol[i] * 100000 / temp;
                        model.Uptime.Add(Math.Round((sumProdVol[i]*100000/temp), 3));
                    } else
                    {
                        model.Uptime.Add(0);
                    }
                    
                }
                model.Uptime.Add(0);


                //*********************************************************************************************************** Uptime Theo
                model.UptimeTheo = new List<decimal>();
                tempTotal = 0;
                for (var i = 0; i< model.Uptime.Count(); i++)
                {
                    if (model.Uptime[i] > 0)
                    {
                        tempTotal += model.Volume[i] * 100 / model.Uptime[i];
                        model.UptimeTheo.Add(Math.Round((model.Volume[i] * 100 / model.Uptime[i]), 3));
                    }
                    else
                    {
                        model.UptimeTheo.Add(0);
                    }
                }
                model.UptimeTheo[model.Uptime.Count() - 1] = Math.Round(tempTotal, 3);
                if (model.UptimeTheo[model.Uptime.Count() - 1] > 0)
                    model.Uptime[model.Uptime.Count() - 1] = Math.Round((model.Volume[model.Uptime.Count() - 1] * 100 / model.UptimeTheo[model.Uptime.Count() - 1]), 3);

                //********************************************************************************************************************* CRR
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var sumMacPac = dataCRR.Where(x => x.Week == i && x.MachineType.Trim().ToLower() != "linkup").Sum(x => x.Quantity);
                    var pembagi = sumMacPac + sumProdVol[i];
                    decimal temp = 0;

                    if (pembagi != 0)
                        temp = sumMacPac * 100 / pembagi;

                    tempTotal += temp * sumProdVol[i];
                    model.CRR.Add(Math.Round(temp, 3));
                }
                model.CRR.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.CRR[model.CRR.Count()-1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //*************************************************************************************************************** ACT YIELD
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i - startweek] != 0)
                        temp = (GRCutfillerActualWet[i - startweek] + WIPActualWet[i - startweek] + GRRSWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    tempTotal += temp * sumProdVol[i];
                    model.ActYield.Add(Math.Round(temp, 3));
                }
                model.ActYield.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.ActYield[model.ActYield.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //*************************************************************************************************************** DIM WASTE

                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() != "" && x.Description.Trim().ToLower() != "cutfiller" && x.Description.Trim().ToLower() != "rejected cigarettes").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() != "" && x.Description.Trim().ToLower() != "cutfiller" && x.Description.Trim().ToLower() != "rejected cigarettes").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.DIMWaste.Add(Math.Round(temp, 3));
                }
                model.DIMWaste.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.DIMWaste[model.DIMWaste.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //*************************************************************************************************************** CLAIMABLE

                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var tempCRR = dataCRRConversion.Where(x => x.Week == i).ToList();
                    decimal temp = 0;

                    if (tempCRR != null && tempCRR.Count() > 0)
                    {
                        temp = (tempCRR.FirstOrDefault().ClaimableInThStick / 1000000);
                    }

                    tempTotal += temp;
                    model.ClaimableInThStick.Add(Math.Round(temp, 3));
                }
                model.ClaimableInThStick.Add(Math.Round(tempTotal, 3));

                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (model.Volume[i - startweek] != 0)
                        temp = model.ClaimableInThStick[i-startweek] * 100.0M / model.Volume[i - startweek];

                    model.Claimable.Add(Math.Round(temp, 3));
                }
                model.Claimable.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.Claimable[model.Claimable.Count() - 1] = Math.Round((model.ClaimableInThStick[model.ClaimableInThStick.Count() - 1] / sumProdVol.Sum()), 3);

                //********************************************************************************************************* TAX STAMP WASTE

                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "sticker").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "sticker").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.TaxStampWaste.Add(Math.Round(temp, 3));
                }
                model.TaxStampWaste.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.TaxStampWaste[model.TaxStampWaste.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //************************************************************************************************************* Filter Rod
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "filter rod").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "filter rod").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.FilterRod.Add(Math.Round(temp, 3));
                }
                model.FilterRod.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.FilterRod[model.FilterRod.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);


                //******************************************************************************************************** Hinge Lid Blank
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "hinge lid blank").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "hinge lid blank").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.HingeLidBlank.Add(Math.Round(temp, 3));
                }
                model.HingeLidBlank.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.HingeLidBlank[model.HingeLidBlank.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //************************************************************************************************************ Inner Liner
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "inner liner").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "inner liner").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.InnerLiner.Add(Math.Round(temp, 3));
                }
                model.InnerLiner.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.InnerLiner[model.InnerLiner.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //********************************************************************************************************** Tipping Paper
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "tipping paper").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "tipping paper").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.TippingPaper.Add(Math.Round(temp, 3));
                }
                model.TippingPaper.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.TippingPaper[model.TippingPaper.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //********************************************************************************************************** Wrapping Film
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "wrapping film").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "wrapping film").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.WrappingFilm.Add(Math.Round(temp, 3));
                }
                model.WrappingFilm.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.WrappingFilm[model.WrappingFilm.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //*********************************************************************************************************** Inner Frame
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "inner frame").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "inner frame").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.InnerFrame.Add(Math.Round(temp, 3));
                }
                model.InnerFrame.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.InnerFrame[model.InnerFrame.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

                //******************************************************************************************************** Cigarette Paper
                
                tempTotal = 0;
                for (var i = startweek; i < totalWeek; i++)
                {
                    var Good = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "cigarette paper").Sum(x => x.GoodsIssueValue);
                    var wWaste = dataDIM.Where(x => x.Week == i && x.Description.Trim().ToLower() == "cigarette paper").Sum(x => x.StdUsgVal);

                    decimal temp = 0;

                    if (wWaste != 0)
                        temp = (Good - wWaste) * 100 / wWaste;

                    tempTotal += temp * sumProdVol[i];
                    model.CigarettePaper.Add(Math.Round(temp, 3));
                }
                model.CigarettePaper.Add(0);
                if (sumProdVol.Sum() > 0)
                    model.CigarettePaper[model.CigarettePaper.Count() - 1] = Math.Round((tempTotal / sumProdVol.Sum()), 3);

            }
            catch (Exception e)
            {

            }

            return PartialView(model);
        }
        public ActionResult ReportYield(string PC = "", int year = 0, int showall = 0)
        {
            var model = new ReportWeeklyYieldModel();
            try
            {
                int week = 1; // fix saja
                if (year == 0)
                    year = DateTime.Now.Year;

                ViewBag.PC = PC;
                ViewBag.year = year;
                ViewBag.week = week;
                ViewBag.showall = showall;

                var dataYield = _reportKPIYieldAppService.GetAll(true).DeserializeToReportKPIYieldModelList().Where(x => x.Year == year).ToList();
                var dataDust = _reportKPIDustAppService.GetAll(true).DeserializeToReportKPIDustModelList();
                var dataTobaccoWeight = _reportKPITobaccoWeightAppService.GetAll(true).DeserializeToReportKPITobaccoWeightModelList().Where(x => x.Year == year).ToList();
                var dataRipperInfo = _reportKPIRipperInfoAppService.GetAll(true).DeserializeToReportKPIRipperInfoModelList().Where(x => x.Year == year).ToList();
                var dataCRRConversion = _reportKPICRRConversionAppService.GetAll(true).DeserializeToReportKPICRRConversionModelList().Where(x => x.Year == year).ToList();

                if (PC != "")
                {
                    dataYield = dataYield.Where(x => x.ProductionCenter == PC).ToList();
                    dataDust = dataDust.Where(x => x.ProductionCenter == PC).ToList();
                    dataTobaccoWeight = dataTobaccoWeight.Where(x => x.ProductionCenter == PC).ToList();
                    dataRipperInfo = dataRipperInfo.Where(x => x.ProductionCenter == PC).ToList();
                    dataCRRConversion = dataCRRConversion.Where(x => x.ProductionCenter == PC).ToList();
                }
                    

                var totalWeek = GetWeekNumber(DateTime.Now);
                var startweek = week;

                model.Header = new List<string>();
                

                var GICutfillerWet = new List<decimal>();
                var GRCutfillerSpecWet = new List<decimal>();
                var GRCutfillerActualWet = new List<decimal>();
                var WIPSpecWet = new List<decimal>();
                var WIPActualWet = new List<decimal>();
                var GIRJRipperWet = new List<decimal>();
                var GRRSWet = new List<decimal>();

                //CRR Conversion
                var SaponMakerWet = new List<decimal>();
                var SaponPackerWet = new List<decimal>();
                var LMDustCollectorWet = new List<decimal>();

                var GICutfillerDry = new List<decimal>();
                var GRCutfillerActualDry = new List<decimal>();
                var WIPActualDry = new List<decimal>();
                var GIRJRipperDry = new List<decimal>();
                var GRRSDry = new List<decimal>();
                var MCCutfiller = new List<decimal>();

                var SaponMakerDry = new List<decimal>();
                var SaponPackerDry = new List<decimal>();
                var LMDustCollectorDry = new List<decimal>();

                var MeanMC = new List<decimal>();
                var Mean = new List<decimal>();
                var C2Weight = new List<decimal>();

                for (var i = startweek; i < totalWeek; i++)
                {
                    model.Header.Add("W" + i.ToString("00"));

                    var tempYield = dataYield.Where(x => x.Week == i).ToList();
                    var tempRipper = dataRipperInfo.Where(x => x.Week == i).ToList();
                    var tempTobacco = dataTobaccoWeight.Where(x => x.Week == i).ToList();
                    var tempDust = dataDust;
                    var tempCRR = dataCRRConversion.Where(x => x.Week == i).ToList();

                    GICutfillerWet.Add(tempYield.Sum(x => x.TotTKGGI));
                    GRCutfillerSpecWet.Add(tempYield.Where(x => x.MaterialGroupDesc.ToLower().Contains("cigarette var")).Sum(x => x.TotTKGGR));
                    
                    if (tempTobacco != null && tempTobacco.Count() > 0)
                    {
                        GRCutfillerActualWet.Add(tempYield.Where(x => x.MaterialGroupDesc.ToLower().Contains("cigarette var")).Sum(x => x.GoodsReceiptQty * (tempTobacco.FirstOrDefault().Mean / 1000000000000000)));
                        MeanMC.Add(tempTobacco.FirstOrDefault().MeanMC / 10000000000);
                        Mean.Add(tempTobacco.FirstOrDefault().Mean / 1000000000000000);
                        GRCutfillerActualDry.Add(GRCutfillerActualWet[i - startweek] * (1 - ((tempTobacco.FirstOrDefault().MeanMC / 10000000000) / 100)));
                    }
                    else
                    {
                        GRCutfillerActualWet.Add(0);
                        MeanMC.Add(0);
                        Mean.Add(0);
                        GRCutfillerActualDry.Add(0);
                    }

                    WIPSpecWet.Add(tempYield.Where(x => x.MaterialGroupDesc.ToLower().Contains("reject cig")).Sum(x => x.TotTKGWIP));
                    WIPActualWet.Add(WIPSpecWet[i-startweek]);
                    GIRJRipperWet.Add(tempRipper.Select(x => x.QtyIss).Sum());
                    GRRSWet.Add(tempRipper.Select(x => x.QtyRec).Sum());

                    if (tempCRR != null && tempCRR.Count() > 0)
                    {
                        SaponMakerWet.Add(tempCRR.FirstOrDefault().SaponTembakau / 1000000);
                        SaponPackerWet.Add((tempCRR.FirstOrDefault().SaponCigarette/1000000) * (decimal)0.717);
                        LMDustCollectorWet.Add(tempCRR.FirstOrDefault().DustHalusLMnoRipper / 1000000);
                        MCCutfiller.Add(tempCRR.FirstOrDefault().AverageFinalOv / 1000000);
                        C2Weight.Add(tempCRR.FirstOrDefault().C2Weight / 1000000);
                    } else
                    {
                        SaponMakerWet.Add(0);
                        SaponPackerWet.Add(0);
                        LMDustCollectorWet.Add(0);
                        MCCutfiller.Add(0);
                        C2Weight.Add(0);
                    }
                    
                    GICutfillerDry.Add(GICutfillerWet[i - startweek] * (1 - (MCCutfiller[i - startweek]/100)));
                    

                    if (tempDust!=null && tempDust.Count() > 0)
                    {
                        WIPActualDry.Add(WIPActualWet[i - startweek] * (1 - ((decimal)tempDust.FirstOrDefault().RS / 100)) * (decimal)0.95);
                        GIRJRipperDry.Add(GIRJRipperWet[i - startweek] * (1 - ((decimal)tempDust.FirstOrDefault().RS / 100)));
                        GRRSDry.Add(GRRSWet[i - startweek] * (1 - ((decimal)tempDust.FirstOrDefault().RS / 100)));
                        SaponMakerDry.Add(SaponMakerWet[i - startweek] * (1-((decimal)tempDust.FirstOrDefault().FloorSweeping/100)));
                        SaponPackerDry.Add(SaponPackerWet[i - startweek] * (1 - ((decimal)tempDust.FirstOrDefault().FloorSweeping / 100)));
                        LMDustCollectorDry.Add(LMDustCollectorWet[i - startweek] * (1 - ((decimal)tempDust.FirstOrDefault().Dust / 100)));

                    }
                    else
                    {
                        WIPActualDry.Add(0);
                        GIRJRipperDry.Add(0);
                        GRRSDry.Add(0);
                        SaponMakerDry.Add(0);
                        SaponPackerDry.Add(0);
                        LMDustCollectorDry.Add(0);

                    }
                }

                //********************************************************************************************************************* WET
                model.SecSpecWetYield = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i-startweek] != 0)
                        temp = (GRCutfillerSpecWet[i - startweek] + WIPSpecWet[i - startweek] + GRRSWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    model.SecSpecWetYield.Add(Math.Round(temp, 3));
                }

                model.SecActualWetYield = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i - startweek] != 0)
                        temp = (GRCutfillerActualWet[i - startweek] + WIPActualWet[i - startweek] + GRRSWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    model.SecActualWetYield.Add(Math.Round(temp, 3));
                }

                model.FloorSweptWet = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i - startweek] != 0)
                        temp = (SaponMakerWet[i - startweek] + SaponPackerWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    model.FloorSweptWet.Add(Math.Round(temp, 3));
                }

                model.DustHalusWet = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i - startweek] != 0)
                        temp = (LMDustCollectorWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    model.DustHalusWet.Add(Math.Round(temp, 3));
                }

                model.RippingLossWet = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerWet[i - startweek] != 0)
                        temp = (GIRJRipperWet[i - startweek] - GRRSWet[i - startweek]) * 100 / GICutfillerWet[i - startweek];

                    model.RippingLossWet.Add(Math.Round(temp, 3));
                }

                model.UnaccountableWet = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    var sum = model.SecActualWetYield[i - startweek] + model.FloorSweptWet[i - startweek] + model.DustHalusWet[i - startweek] + model.RippingLossWet[i - startweek];
                    decimal temp = 0;
                    if (sum > 0)
                        temp = 100 - sum;

                    model.UnaccountableWet.Add(Math.Round(temp, 3));
                }

                //********************************************************************************************************************* DRY

                model.SecActualDryYield = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerDry[i - startweek] != 0)
                        temp = (GRCutfillerActualDry[i - startweek] + WIPActualDry[i - startweek] + GRRSDry[i - startweek]) * 100 / GICutfillerDry[i - startweek];

                    model.SecActualDryYield.Add(Math.Round(temp, 3));
                }

                model.FloorSweptDry = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerDry[i - startweek] != 0)
                        temp = (SaponMakerDry[i - startweek] + SaponPackerDry[i - startweek]) * 100 / GICutfillerDry[i - startweek];

                    model.FloorSweptDry.Add(Math.Round(temp, 3));
                }

                model.DustHalusDry = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerDry[i - startweek] != 0)
                        temp = (LMDustCollectorDry[i - startweek] * 100) / GICutfillerDry[i - startweek];

                    model.DustHalusDry.Add(Math.Round(temp, 3));
                }

                model.RippingLossDry = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    decimal temp = 0;
                    if (GICutfillerDry[i - startweek] != 0)
                        temp = ((GIRJRipperDry[i - startweek] - GRRSDry[i - startweek]) * 100) / GICutfillerDry[i - startweek];

                    model.RippingLossDry.Add(Math.Round(temp, 3));
                }

                model.UnaccountableDry = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    var sum = model.SecActualDryYield[i - startweek] + model.FloorSweptDry[i - startweek] + model.DustHalusDry[i - startweek] + model.RippingLossDry[i - startweek];
                    decimal temp = 0;
                    if (sum > 0)
                        temp = 100 - sum;

                    model.UnaccountableDry.Add(Math.Round(temp, 3));
                }

                //**************************************************************************************************************** TAMBAHAN
                model.MCCutfiller = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    model.MCCutfiller.Add(Math.Round(MCCutfiller[i - startweek], 3));
                }

                model.PackOV = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    model.PackOV.Add(Math.Round(MeanMC[i - startweek], 3));
                }

                model.WeightATW = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    model.WeightATW.Add(Math.Round(Mean[i - startweek], 3));
                }

                model.WeightC2Weight = new List<decimal>();
                for (var i = startweek; i < totalWeek; i++)
                {
                    model.WeightC2Weight.Add(Math.Round(C2Weight[i - startweek], 3));
                }
            }
            catch (Exception e)
            {

            }

            return PartialView(model);
        }
        public ActionResult ReportTRS(string dateFrom = "", string dateTo = "", string PC="PJ")
        {
            var model = new ReportTRSModel();
            try
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

                var EndDate = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);

                var monday4weekbefore = EndDate.AddDays(-21).StartOfWeek(DayOfWeek.Monday);
                var StartDate = (dateFrom == "") ? monday4weekbefore : Convert.ToDateTime(dateFrom);

                if (StartDate > EndDate)
                {
                    Session["ResultLog"] = "error_Date Start is higher than Date End";
                } else
                {
                    ViewBag.PC = PC;
                    ViewBag.dateFrom = StartDate.ToString("dd-MMM-yy");
                    ViewBag.dateTo = EndDate.ToString("dd-MMM-yy");
                    ViewBag.weekFrom = GetWeekNumber(StartDate);
                    ViewBag.weekTo = GetWeekNumber(EndDate);

                    model.header = new List<string> { "date", "time", "speed", "CTW", "sample", "sampling_time", "recovery", "remark" };

                    string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    var myQuery = @"SELECT * FROM LPHSubmissions WHERE (convert(date,[Date]) BETWEEN '" + StartDate.Date.ToShortDateString() + "' AND '" + EndDate.Date.ToShortDateString() + "' AND LPHHeader = 'MakerController' AND IsComplete = 1 AND IsDeleted = 0)";

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
                    List<LPHSubmissionsModel> submissions = jsondata.DeserializeToLPHSubmissionsList();

                    //ambil data sesuai filter lokasi
                    List<long> locations = new List<long>();
                    if (!string.IsNullOrWhiteSpace(PC))
                    {
                        var location = _locationAppService.FindBy("ParentID", 1, true).DeserializeToLocationList().Where(x => x.Code == PC).FirstOrDefault();

                        if (location != null && location.ID > 0)
                        {
                            locations.Add(location.ID);

                            string deps = _locationAppService.FindBy("ParentID", location.ID, true);
                            var depsM = deps.DeserializeToLocationList();

                            foreach (var dep in depsM)
                            {
                                locations.Add(dep.ID);

                                string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                                var subdepsM = subdeps.DeserializeToLocationList();

                                foreach (var subdep in subdepsM)
                                {
                                    locations.Add(subdep.ID);
                                }
                            }

                            submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                    }

                    if (submissions.Count() > 0)
                    {
                        var submissionsID = submissions.Select(x => x.ID).ToList();
                        var LPHsID = submissions.Select(x => x.LPHID).ToList();

                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        myQuery = @"SELECT * FROM LPHs WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", LPHsID) + ")";

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
                        List<LPHModel> LPHList = jsondata.DeserializeToLPHList();

                        //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                        foreach (var item in submissions.ToList())
                        {
                            LPHModel lphModel = LPHList.Where(x => x.ID == item.LPHID).FirstOrDefault();

                            // chanif: exclude LPH yang sudah dihapus
                            if (lphModel == null)
                            {
                                submissions.Remove(item);
                                continue;
                            }
                            else if (lphModel.IsDeleted)
                            {
                                submissions.Remove(item);
                                continue;
                            }
                        }

                        if (submissions.Count() > 0)
                        {
                            var LPHIDs = submissions.Select(x => x.LPHID).ToList();

                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHExtras WHERE IsDeleted = 0 AND HeaderName = 'trsrate' AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

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
                            List<LPHExtrasModel> extraList = jsondata.DeserializeToLPHExtraList().OrderBy(x => x.LPHID).ToList();

                            model.temp = new List<List<dynamic>>();
                            foreach (var subshere in submissions)
                            {
                                var thisExtra = extraList.Where(x => x.LPHID == subshere.LPHID).ToList();
                                var isi = new List<dynamic>();

                                isi.Add(subshere.LPHID);
                                isi.Add(subshere.Machine);
                                isi.Add(subshere.Shift);

                                foreach (var th in model.header)
                                {
                                    var tempval = thisExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                    dynamic content = "";

                                    if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                    {
                                        if (tempval.ValueType.Trim() == "Numeric")
                                        {
                                            if (tempval.Value.Contains("."))
                                            {
                                                //tak anggap double
                                                double number = 0;
                                                Double.TryParse(tempval.Value, out number);
                                                content = number;
                                            }
                                            else
                                            {
                                                //tak anggep long
                                                Int64 number = 0;
                                                Int64.TryParse(tempval.Value, out number);
                                                content = number;
                                            }
                                        }
                                        else if (tempval.ValueType.Trim() == "ImageURL")
                                        {
                                            if (tempval.Value == "_no_image.png")
                                            {
                                                content = "no image";
                                            }
                                            else
                                            {
                                                content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + subshere.LPHHeader.ToLower() + "/" + tempval.Value;
                                            }
                                        }
                                        else
                                        {
                                            content = tempval.Value;
                                        }

                                        isi.Add(content);
                                    }
                                    thisExtra.Remove(tempval);
                                }

                                if(isi.Count()==11)
                                    model.temp.Add(isi);

                                

                                
                            }

                            model.content = new List<TRSModel>();
                            foreach (var dt in model.temp)
                            {
                                var temp = new TRSModel();
                                var counter = 0;

                                temp.LPHID = dt[counter++];
                                temp.Maker = dt[counter++];
                                temp.Shift = dt[counter++];
                                temp.Date = dt[counter++];
                                temp.Week = GetWeekNumber(DateTime.Parse(temp.Date));
                                temp.Time = dt[counter++];
                                temp.Speed = dt[counter++];
                                temp.CTW = dt[counter++];
                                temp.Sample = dt[counter++];
                                temp.Sampling_time = dt[counter++];
                                temp.Recovery = dt[counter++];
                                temp.Remark = dt[counter++];

                                model.content.Add(temp);
                            }

                            model.content = model.content.OrderBy(x => x.Date).ThenBy(x => x.Maker).ToList();

                            foreach (var data in model.content)
                            {
                                var tempDate = model.content.Where(x => x.Date == data.Date).ToList();
                                var tempMaker = tempDate.Where(x => x.Maker == data.Maker).ToList();
                                var tempShift = tempMaker.Where(x => x.Shift == data.Shift).ToList();

                                var dataDate = tempDate.Select(x => (decimal)x.Recovery).ToList();
                                var dataMaker = tempMaker.Select(x => (decimal)x.Recovery).ToList();
                                var dataShift = tempShift.Select(x => (decimal)x.Recovery).ToList();

                                data.TRSRecoveryDate = 0;
                                data.TRSRecoveryMaker = 0;
                                data.TRSRecoveryShift = 0;
                                //data.ConfidenceMaker = 0;
                                //data.ConfidenceDate = 0;
                                try
                                {
                                    data.TRSRecoveryDate = Math.Round(tempDate.Sum(x => (decimal)x.Recovery) / (decimal)tempDate.Count(), 3);
                                    data.TRSRecoveryMaker = Math.Round(tempMaker.Sum(x => (decimal)x.Recovery) / (decimal)tempMaker.Count(), 3);
                                    data.TRSRecoveryShift = Math.Round(tempShift.Sum(x => (decimal)x.Recovery) / (decimal)tempShift.Count(), 3);
                                }
                                catch (Exception e)
                                {

                                }
                            }

                        }
                        else
                        {
                            Session["ResultLog"] = "error_No Data";
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Session["ResultLog"] = "error_Detail error: "+ex.InnerException;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return PartialView(model);
        }

        public ActionResult DownloadXlsTRS(string dateFrom = "", string dateTo = "", string PC = "")
        {
            var model = new ReportTRSModel();
            try
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

                var EndDate = (dateTo == "") ? DateTime.Now : Convert.ToDateTime(dateTo);

                var monday4weekbefore = EndDate.AddDays(-21).StartOfWeek(DayOfWeek.Monday);
                var StartDate = (dateFrom == "") ? monday4weekbefore : Convert.ToDateTime(dateFrom);

                if (StartDate > EndDate)
                {
                    Session["ResultLog"] = "error_Date Start is higher than Date End";
                }
                else
                {
                    ViewBag.PC = PC;
                    ViewBag.dateFrom = StartDate.ToString("dd-MMM-yy");
                    ViewBag.dateTo = EndDate.ToString("dd-MMM-yy");
                    ViewBag.weekFrom = GetWeekNumber(StartDate);
                    ViewBag.weekTo = GetWeekNumber(EndDate);

                    model.header = new List<string> { "date", "time", "speed", "CTW", "sample", "sampling_time", "recovery", "remark" };

                    string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                    var myQuery = @"SELECT * FROM LPHSubmissions WHERE (convert(date,[Date]) BETWEEN '" + StartDate.Date.ToShortDateString() + "' AND '" + EndDate.Date.ToShortDateString() + "' AND LPHHeader = 'MakerController' AND IsComplete = 1 AND IsDeleted = 0)";

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
                    List<LPHSubmissionsModel> submissions = jsondata.DeserializeToLPHSubmissionsList();

                    //ambil data sesuai filter lokasi
                    List<long> locations = new List<long>();
                    if (!string.IsNullOrWhiteSpace(PC))
                    {
                        var location = _locationAppService.FindBy("ParentID", 1, true).DeserializeToLocationList().Where(x => x.Code == PC).FirstOrDefault();

                        if (location != null && location.ID > 0)
                        {
                            locations.Add(location.ID);

                            string deps = _locationAppService.FindBy("ParentID", location.ID, true);
                            var depsM = deps.DeserializeToLocationList();

                            foreach (var dep in depsM)
                            {
                                locations.Add(dep.ID);

                                string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                                var subdepsM = subdeps.DeserializeToLocationList();

                                foreach (var subdep in subdepsM)
                                {
                                    locations.Add(subdep.ID);
                                }
                            }

                            submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                    }

                    if (submissions.Count() > 0)
                    {
                        var submissionsID = submissions.Select(x => x.ID).ToList();
                        var LPHsID = submissions.Select(x => x.LPHID).ToList();

                        strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                        myQuery = @"SELECT * FROM LPHs WHERE IsDeleted = 0 AND ID IN (" + string.Join(",", LPHsID) + ")";

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
                        List<LPHModel> LPHList = jsondata.DeserializeToLPHList();

                        //setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
                        foreach (var item in submissions.ToList())
                        {
                            LPHModel lphModel = LPHList.Where(x => x.ID == item.LPHID).FirstOrDefault();

                            // chanif: exclude LPH yang sudah dihapus
                            if (lphModel == null)
                            {
                                submissions.Remove(item);
                                continue;
                            }
                            else if (lphModel.IsDeleted)
                            {
                                submissions.Remove(item);
                                continue;
                            }
                        }

                        if (submissions.Count() > 0)
                        {
                            var LPHIDs = submissions.Select(x => x.LPHID).ToList();

                            strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                            myQuery = @"SELECT * FROM LPHExtras WHERE IsDeleted = 0 AND HeaderName = 'trsrate' AND LPHID IN (" + string.Join(",", LPHIDs) + ")";

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
                            List<LPHExtrasModel> extraList = jsondata.DeserializeToLPHExtraList().OrderBy(x => x.LPHID).ToList();

                            model.temp = new List<List<dynamic>>();
                            foreach (var subshere in submissions)
                            {
                                var thisExtra = extraList.Where(x => x.LPHID == subshere.LPHID).ToList();
                                var isi = new List<dynamic>();

                                isi.Add(subshere.LPHID);
                                isi.Add(subshere.Machine);
                                isi.Add(subshere.Shift);

                                foreach (var th in model.header)
                                {
                                    var tempval = thisExtra.Where(x => x.FieldName == th).FirstOrDefault();

                                    dynamic content = "";

                                    if (tempval != null && !String.IsNullOrEmpty(tempval.Value))
                                    {
                                        if (tempval.ValueType.Trim() == "Numeric")
                                        {
                                            if (tempval.Value.Contains("."))
                                            {
                                                //tak anggap double
                                                double number = 0;
                                                Double.TryParse(tempval.Value, out number);
                                                content = number;
                                            }
                                            else
                                            {
                                                //tak anggep long
                                                Int64 number = 0;
                                                Int64.TryParse(tempval.Value, out number);
                                                content = number;
                                            }
                                        }
                                        else if (tempval.ValueType.Trim() == "ImageURL")
                                        {
                                            if (tempval.Value == "_no_image.png")
                                            {
                                                content = "no image";
                                            }
                                            else
                                            {
                                                content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + subshere.LPHHeader.ToLower() + "/" + tempval.Value;
                                            }
                                        }
                                        else
                                        {
                                            content = tempval.Value;
                                        }

                                        isi.Add(content);
                                    }
                                    thisExtra.Remove(tempval);
                                }

                                if (isi.Count() == 11)
                                    model.temp.Add(isi);




                            }

                            model.content = new List<TRSModel>();
                            foreach (var dt in model.temp)
                            {
                                var temp = new TRSModel();
                                var counter = 0;

                                temp.LPHID = dt[counter++];
                                temp.Maker = dt[counter++];
                                temp.Shift = dt[counter++];
                                temp.Date = dt[counter++];
                                temp.Week = GetWeekNumber(DateTime.Parse(temp.Date));
                                temp.Time = dt[counter++];
                                temp.Speed = dt[counter++];
                                temp.CTW = dt[counter++];
                                temp.Sample = dt[counter++];
                                temp.Sampling_time = dt[counter++];
                                temp.Recovery = dt[counter++];
                                temp.Remark = dt[counter++];

                                model.content.Add(temp);
                            }

                            model.content = model.content.OrderBy(x => x.Date).ThenBy(x => x.Maker).ToList();

                            foreach (var data in model.content)
                            {
                                var tempDate = model.content.Where(x => x.Date == data.Date).ToList();
                                var tempMaker = tempDate.Where(x => x.Maker == data.Maker).ToList();
                                var tempShift = tempMaker.Where(x => x.Shift == data.Shift).ToList();

                                var dataDate = tempDate.Select(x => (decimal)x.Recovery).ToList();
                                var dataMaker = tempMaker.Select(x => (decimal)x.Recovery).ToList();
                                var dataShift = tempShift.Select(x => (decimal)x.Recovery).ToList();

                                data.TRSRecoveryDate = 0;
                                data.TRSRecoveryMaker = 0;
                                data.TRSRecoveryShift = 0;
                                //data.ConfidenceMaker = 0;
                                //data.ConfidenceDate = 0;
                                try
                                {
                                    data.TRSRecoveryDate = Math.Round(tempDate.Sum(x => (decimal)x.Recovery) / (decimal)tempDate.Count(), 3);
                                    data.TRSRecoveryMaker = Math.Round(tempMaker.Sum(x => (decimal)x.Recovery) / (decimal)tempMaker.Count(), 3);
                                    data.TRSRecoveryShift = Math.Round(tempShift.Sum(x => (decimal)x.Recovery) / (decimal)tempShift.Count(), 3);
                                }
                                catch (Exception e)
                                {

                                }
                            }

                        }
                        else
                        {
                            Session["ResultLog"] = "error_No Data";
                        }
                    }
                }



                ExcelPackage Ep = new ExcelPackage();
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Sheet 1");

                using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                {
                    var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                    excelImage.SetPosition(0, 0, 0, 0);
                }

                Sheet.Cells["A3"].Value = UIResources.Title;
                Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                Sheet.Cells["B3"].Value = "Report KPI > Report TRS Rate";
                Sheet.Cells["B4"].Value = AccountName;
                Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                int startrow = 1;
                int numrow = 1;
                int numcol = 1;


                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));

                Sheet.Cells[8, 1].Value = "Date";
                Sheet.Cells[9, 1].Value = "Production Center";

                Sheet.Cells[8, 2].Value = dateFrom +" - "+ dateTo;
                Sheet.Cells[9, 2].Value = PC;

                startrow = 12;
                numrow = 12;
                numcol = 1;

                var makers = model.content.Select(x => x.Maker).Distinct().OrderBy(x => x).ToList();
                var hariawal = DateTime.Parse(ViewBag.dateTo).AddDays(-6);

                Sheet.Cells[numrow++, numcol].Value = "DATA WEEKLY";
                Sheet.Cells[numrow, numcol++].Value = "Maker";
                for (var i = ViewBag.weekFrom; i <= ViewBag.weekTo; i++)
                {
                    Sheet.Cells[numrow, numcol++].Value = "W"+i;
                }
                Sheet.Cells[numrow, numcol++].Value = "AVG";

                var monthly = new List<List<double>>();
                foreach (var maker in makers)
                {
                    int counter = 0;
                    double jumlah = 0;
                    var bulanan = new List<double>();

                    numrow++;
                    numcol = 1;

                    Sheet.Cells[numrow, numcol++].Value = maker;



                    for (var i = ViewBag.weekFrom; i <= ViewBag.weekTo; i++)
                    {
                        double sumweek = 0;
                        var markerweek = model.content.Where(x => x.Week == i && x.Maker == maker).Select(x => x.Recovery).ToList();
                        if (markerweek != null && markerweek.Count() > 0)
                        {
                            try
                            {
                                foreach (var lala in markerweek)
                                {
                                    sumweek += lala;
                                }
                            }
                            catch (Exception x)
                            {
                            }
                        }

                        Sheet.Cells[numrow, numcol++].Value = Math.Round(sumweek, 2);
                        bulanan.Add(Math.Round(sumweek, 2));
                        jumlah += sumweek;
                        counter++;
                    }
                    Sheet.Cells[numrow, numcol++].Value = (jumlah > 0 ? Math.Round(jumlah / counter, 2) : 0);
                    bulanan.Add(jumlah > 0 ? Math.Round(jumlah / counter, 2) : 0);
                    monthly.Add(bulanan);
                }

                numrow++;
                numcol = 1;
                var total = new List<double>();

                Sheet.Cells[numrow, numcol++].Value = "Grand Total";
                for (var i = ViewBag.weekFrom; i <= ViewBag.weekTo; i++)
                {
                    double sumweek = 0;
                    for (int j = 0; j < makers.Count(); j++)
                    {
                        sumweek += monthly[j][i - ViewBag.weekFrom];
                    }

                    var aver = sumweek / makers.Count();

                    Sheet.Cells[numrow, numcol++].Value = Math.Round(aver, 2);
                    total.Add(Math.Round(aver, 2));
                }
                Sheet.Cells[numrow, numcol++].Value = Math.Round(total.Average(x => x), 2);

                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(HEADER_COLOR);

                Sheet.Cells[startrow, 1, numrow, numcol--].AutoFitColumns();
                using (var range = Sheet.Cells[startrow, 1, startrow+1, numcol])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                }

                Sheet.Cells[startrow, 1, numrow, numcol].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                startrow = numrow+3;
                numrow = startrow;
                numcol = 1;

                Sheet.Cells[numrow++, numcol].Value = "DATA DAILY (LAST 7 DAYS)";
                Sheet.Cells[numrow, numcol++].Value = "Maker";
                foreach (var maker in makers)
                {
                    Sheet.Cells[numrow, numcol++].Value = maker;
                }
                
                for (var i = 0; i < 7; i++)
                {
                    int counter = 0;
                    double jumlah = 0;
                    DateTime sekarang = hariawal.AddDays(i);

                    numrow++;
                    numcol = 1;

                    Sheet.Cells[numrow, numcol++].Value = sekarang.ToString("dddd, dd-MMM-yy");
                        
                        foreach (var maker in makers)
                        {
                            double sumweek = 0;
                            try
                            {
                                var markerweek = model.content.Where(x => DateTime.Parse(x.Date) == sekarang && x.Maker == maker).Select(x => x.Recovery).ToList();
                                if (markerweek != null && markerweek.Count() > 0)
                                {
                                    foreach (var lala in markerweek)
                                    {
                                        sumweek += lala;
                                    }
                                }
                            }
                            catch (Exception x)
                            {

                            }

                            Sheet.Cells[numrow, numcol++].Value = Math.Round(sumweek, 2);

                            jumlah += sumweek;
                            counter++;
                        }
                }

                Sheet.Cells[startrow, 1, numrow, numcol--].AutoFitColumns();
                using (var range = Sheet.Cells[startrow, 1, startrow+1, numcol])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                }

                Sheet.Cells[startrow, 1, numrow, numcol].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                startrow = numrow + 3;
                numrow = startrow;
                numcol = 1;

                Sheet.Cells[numrow++, numcol].Value = "RAW DATA";
                Sheet.Cells[numrow, numcol++].Value = "LPHID";
                Sheet.Cells[numrow, numcol++].Value = "Date";
                Sheet.Cells[numrow, numcol++].Value = "Maker";
                Sheet.Cells[numrow, numcol++].Value = "Shift";
                Sheet.Cells[numrow, numcol++].Value = "Time";
                Sheet.Cells[numrow, numcol++].Value = "Speed";
                Sheet.Cells[numrow, numcol++].Value = "CTW";
                Sheet.Cells[numrow, numcol++].Value = "Sample";
                Sheet.Cells[numrow, numcol++].Value = "Sampling Time";
                Sheet.Cells[numrow, numcol++].Value = "Recovery";
                Sheet.Cells[numrow, numcol++].Value = "Remark";
                Sheet.Cells[numrow, numcol++].Value = "TRS Recovery Rate by Maker (%)";
                Sheet.Cells[numrow, numcol++].Value = "TRS Recovery Rate by Shift (%)";
                Sheet.Cells[numrow, numcol++].Value = "TRS Recovery Rate by Date (%)";

                if (model.content != null && model.content.Count() > 0)
                {
                    foreach (var data in model.content)
                    {
                        numrow++;
                        numcol = 1;

                        Sheet.Cells[numrow, numcol++].Value = data.LPHID;
                        Sheet.Cells[numrow, numcol++].Value = data.Date;
                        Sheet.Cells[numrow, numcol++].Value = data.Maker;
                        Sheet.Cells[numrow, numcol++].Value = data.Shift;
                        Sheet.Cells[numrow, numcol++].Value = data.Time;
                        Sheet.Cells[numrow, numcol++].Value = data.Speed;
                        Sheet.Cells[numrow, numcol++].Value = data.CTW;
                        Sheet.Cells[numrow, numcol++].Value = data.Sample;
                        Sheet.Cells[numrow, numcol++].Value = data.Sampling_time;
                        Sheet.Cells[numrow, numcol++].Value = data.Recovery;
                        Sheet.Cells[numrow, numcol++].Value = data.Remark;
                        Sheet.Cells[numrow, numcol++].Value = data.TRSRecoveryMaker;
                        Sheet.Cells[numrow, numcol++].Value = data.TRSRecoveryShift;
                        Sheet.Cells[numrow, numcol++].Value = data.TRSRecoveryDate;
                    }
                }
                else
                {
                    numrow++;
                    numcol = 1;
                    Sheet.Cells[numrow, numcol++].Value = "No Data";
                }

                Sheet.Cells[startrow, 1, numrow, numcol--].AutoFitColumns();
                using (var range = Sheet.Cells[startrow, 1, startrow + 1, numcol])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                }

                Sheet.Cells[startrow, 1, numrow, numcol].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + "ReportKPI - Report TRS Rate.xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult DataProdVol(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPIProdVolModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIProdVolAppService.Find(filters).DeserializeToReportKPIProdVolModelList().OrderBy(x=>x.Week).ThenBy(x=>x.ID).ToList();
            return PartialView(model);
        }

        public bool DeleteProdVol(long ID)
        {
            var data = _reportKPIProdVolAppService.GetById(ID, true).DeserializeToReportKPIProdVolModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIProdVolModel>.Serialize(data);
            _reportKPIProdVolAppService.Update(update);

            return true;
        }

        public ActionResult DataDIM(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPIKPIDIMModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIDIMAppService.Find(filters).DeserializeToReportKPIDIMModelList().OrderBy(x => x.Week).ThenBy(x => x.ID).ToList();
            return PartialView(model);
        }
        public bool DeleteDIM(long ID)
        {
            var data = _reportKPIDIMAppService.GetById(ID, true).DeserializeToReportKPIDIMModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIDIMModel>.Serialize(data);
            _reportKPIDIMAppService.Update(update);

            return true;
        }

        public ActionResult DataYield(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPIYieldModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));

            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIYieldAppService.Find(filters).DeserializeToReportKPIYieldModelList().OrderBy(x => x.Week).ThenBy(x => x.ID).ToList();
            return PartialView(model);
        }
        public bool DeleteYield(long ID)
        {
            var data = _reportKPIYieldAppService.GetById(ID, true).DeserializeToReportKPIYieldModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIYieldModel>.Serialize(data);
            _reportKPIYieldAppService.Update(update);

            return true;
        }

        public ActionResult DataCRR(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPICRRModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPICRRAppService.Find(filters).DeserializeToReportKPICRRModelList().OrderBy(x => x.Week).ThenBy(x => x.ID).ToList();
            return PartialView(model);
        }
        public bool DeleteCRR(long ID)
        {
            var data = _reportKPICRRAppService.GetById(ID, true).DeserializeToReportKPICRRModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPICRRModel>.Serialize(data);
            _reportKPICRRAppService.Update(update);

            return true;
        }

        public ActionResult DataWorkHour(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPIWorkHourModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIWorkHourAppService.Find(filters).DeserializeToReportKPIWorkHourModelList().OrderBy(x => x.Week).ThenBy(x => x.Packer).ToList();
            return PartialView(model);
        }
        public bool DeleteWorkHour(long ID)
        {
            var data = _reportKPIWorkHourAppService.GetById(ID, true).DeserializeToReportKPIWorkHourModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIWorkHourModel>.Serialize(data);
            _reportKPIWorkHourAppService.Update(update);

            return true;
        }

        [HttpPost]
        public ActionResult EditWorkHour(ReportKPIWorkHourModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                if (FormData.ID > 0)
                {
                    var data = _reportKPIWorkHourAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPIWorkHourModel>.Serialize(FormData);
                    _reportKPIWorkHourAppService.Update(update);
                } else
                {
                    string insert = JsonHelper<ReportKPIWorkHourModel>.Serialize(FormData);
                    _reportKPIWorkHourAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: "+e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataWorkHour" });
        }

        public ActionResult DataStickPerPack(string PC = "")
        {
            var model = new DataKPIStickPerPackModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            ViewBag.PC = PC;
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIStickPerPackAppService.Find(filters).DeserializeToReportKPIStickPerPackModelList().OrderBy(x => x.Packer).ToList();
            return PartialView(model);
        }
        public bool DeleteStickPerPack(long ID)
        {
            var data = _reportKPIStickPerPackAppService.GetById(ID, true).DeserializeToReportKPIStickPerPackModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIStickPerPackModel>.Serialize(data);
            _reportKPIStickPerPackAppService.Update(update);

            return true;
        }

        [HttpPost]
        public ActionResult EditStickPerPack(ReportKPIStickPerPackModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                if (FormData.ID > 0)
                {
                    var data = _reportKPIStickPerPackAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPIStickPerPackModel>.Serialize(FormData);
                    _reportKPIStickPerPackAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPIStickPerPackModel>.Serialize(FormData);
                    _reportKPIStickPerPackAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataStickPerPack" });
        }


        public ActionResult DataDust(string PC = "")
        {
            var model = new DataKPIDustModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            ViewBag.PC = PC;
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPIDustAppService.Find(filters).DeserializeToReportKPIDustModelList().OrderBy(x => x.ProductionCenter).ToList();
            return PartialView(model);
        }
        public bool DeleteDust(long ID)
        {
            var data = _reportKPIDustAppService.GetById(ID, true).DeserializeToReportKPIDustModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIDustModel>.Serialize(data);
            _reportKPIDustAppService.Update(update);

            return true;
        }
        [HttpPost]
        public ActionResult EditDust(ReportKPIDustModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                if (FormData.ID > 0)
                {
                    var data = _reportKPIDustAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPIDustModel>.Serialize(FormData);
                    _reportKPIDustAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPIDustModel>.Serialize(FormData);
                    _reportKPIDustAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataDust" });
        }

        public ActionResult DataTarget(string PC = "")
        {
            var model = new DataKPITargetModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            ViewBag.PC = PC;
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPITargetAppService.Find(filters).DeserializeToReportKPITargetModelList().OrderBy(x => x.ProductionCenter).ToList();
            return PartialView(model);
        }
        public bool DeleteTarget(long ID)
        {
            var data = _reportKPITargetAppService.GetById(ID, true).DeserializeToReportKPITargetModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPITargetModel>.Serialize(data);
            _reportKPITargetAppService.Update(update);

            return true;
        }
        [HttpPost]
        public ActionResult EditTarget(ReportKPITargetModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                if (FormData.ID > 0)
                {
                    var data = _reportKPITargetAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPITargetModel>.Serialize(FormData);
                    _reportKPITargetAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPITargetModel>.Serialize(FormData);
                    _reportKPITargetAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataTarget" });
        }


        public ActionResult DataTobaccoWeight(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPITobaccoWeightModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            model.Data = _reportKPITobaccoWeightAppService.Find(filters).DeserializeToReportKPITobaccoWeightModelList().OrderBy(x => x.Week).ThenBy(x => x.ProductionCenter).ToList();
            return PartialView(model);
        }
        public bool DeleteTobaccoWeight(long ID)
        {
            var data = _reportKPITobaccoWeightAppService.GetById(ID, true).DeserializeToReportKPITobaccoWeightModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPITobaccoWeightModel>.Serialize(data);
            _reportKPITobaccoWeightAppService.Update(update);

            return true;
        }
        [HttpPost]
        public ActionResult EditTobaccoWeight(ReportKPITobaccoWeightModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                FormData.Mean = Math.Round((FormData.Mean * 1000000000000000), 3);
                FormData.Stdev = Math.Round((FormData.Stdev * 1000000000000000), 3);
                FormData.MeanMC = Math.Round((FormData.MeanMC * 10000000000), 3);
                FormData.StdevMC = Math.Round((FormData.StdevMC * 10000000000), 3);

                if (FormData.ID > 0)
                {
                    var data = _reportKPITobaccoWeightAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPITobaccoWeightModel>.Serialize(FormData);
                    _reportKPITobaccoWeightAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPITobaccoWeightModel>.Serialize(FormData);
                    _reportKPITobaccoWeightAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataTobaccoWeight" });
        }


        public ActionResult DataRipperInfo(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPIRipperInfoModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            try
            {
                model.Data = _reportKPIRipperInfoAppService.Find(filters).DeserializeToReportKPIRipperInfoModelList().OrderBy(x => x.Week).ThenBy(x => x.Material).ToList();

            }
            catch (Exception e)
            {
                var lala = e.InnerException;
            }
            return PartialView(model);
        }

        public bool DeleteRipperInfo(long ID)
        {
            var data = _reportKPIRipperInfoAppService.GetById(ID, true).DeserializeToReportKPIRipperInfoModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPIRipperInfoModel>.Serialize(data);
            _reportKPIRipperInfoAppService.Update(update);

            return true;
        }
        [HttpPost]
        public ActionResult EditRipperInfo(ReportKPIRipperInfoModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                if (FormData.ID > 0)
                {
                    var data = _reportKPIRipperInfoAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPIRipperInfoModel>.Serialize(FormData);
                    _reportKPIRipperInfoAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPIRipperInfoModel>.Serialize(FormData);
                    _reportKPIRipperInfoAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataRipperInfo" });
        }


        public ActionResult DataCRRConversion(int year = 0, string week = "", string PC = "")
        {
            var model = new DataKPICRRConversionModel();

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("IsDeleted", "0"));
            if (year == 0)
                year = DateTime.Now.Year;
            ViewBag.Year = year;
            ViewBag.Week = week;
            ViewBag.PC = PC;

            filters.Add(new QueryFilter("Year", year));
            if (week != "")
                filters.Add(new QueryFilter("Week", week));
            if (PC != "")
                filters.Add(new QueryFilter("ProductionCenter", PC));

            try
            {
                model.Data = _reportKPICRRConversionAppService.Find(filters).DeserializeToReportKPICRRConversionModelList().OrderBy(x => x.Week).ToList();
            }
            catch (Exception e)
            {
                var lala = e.InnerException;
            }
            return PartialView(model);
        }
        public bool DeleteCRRConversion(long ID)
        {
            var data = _reportKPICRRConversionAppService.GetById(ID, true).DeserializeToReportKPICRRConversionModel();
            data.ModifiedBy = AccountName;
            data.ModifiedDate = DateTime.Now;
            data.IsDeleted = true;

            string update = JsonHelper<ReportKPICRRConversionModel>.Serialize(data);
            _reportKPICRRConversionAppService.Update(update);

            return true;
        }

        [HttpPost]
        public ActionResult EditCRRConversion(ReportKPICRRConversionModel FormData)
        {
            try
            {
                FormData.ModifiedBy = AccountName;
                FormData.ModifiedDate = DateTime.Now;

                FormData.DebuFinal = Math.Round((FormData.DebuFinal * 1000000), 3);
                FormData.DustHalusLM = Math.Round((FormData.DustHalusLM * 1000000), 3);
                FormData.DustHalusLMnoRipper = Math.Round((FormData.DustHalusLMnoRipper * 1000000), 3);
                FormData.SaponTembakau = Math.Round((FormData.SaponTembakau * 1000000), 3);
                FormData.SaponCigarette = Math.Round((FormData.SaponCigarette * 1000000), 3);
                FormData.ClaimableInThStick = Math.Round((FormData.ClaimableInThStick * 1000000), 3);
                FormData.AverageFinalOv = Math.Round((FormData.AverageFinalOv * 1000000), 3);
                FormData.C2Weight = Math.Round((FormData.C2Weight * 1000000), 3);

                if (FormData.ID > 0)
                {
                    var data = _reportKPICRRConversionAppService.GetById(FormData.ID, true);
                    string update = JsonHelper<ReportKPICRRConversionModel>.Serialize(FormData);
                    _reportKPICRRConversionAppService.Update(update);
                }
                else
                {
                    string insert = JsonHelper<ReportKPICRRConversionModel>.Serialize(FormData);
                    _reportKPICRRConversionAppService.Add(insert);
                }

                Session["ResultLog"] = "success_Data Submitted";
            }
            catch (Exception e)
            {
                Session["ResultLog"] = "error_Error: " + e.InnerException;
            }

            return RedirectToAction("Index", new { page = "DataCRRConversion" });
        }

        public ActionResult DownloadXls(string data,int year = 0, string week = "", string PC = "")
        {
            try
            {
                ExcelPackage Ep = new ExcelPackage();
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Sheet 1");

                using (System.Drawing.Image image = System.Drawing.Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
                {
                    var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                    excelImage.SetPosition(0, 0, 0, 0);
                }

                Sheet.Cells["A3"].Value = UIResources.Title;
                Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
                Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
                Sheet.Cells["B3"].Value = "Report KPI - Master Data - " + data;
                Sheet.Cells["B4"].Value = AccountName;
                Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");

                int startrow = 1;
                int numrow = 1;
                int numcol = 1;

                if (data == "ProdVol" || data == "DIM" || data == "Yield" || data == "CRR" || data == "WorkHour" || data == "StickPerPack" || data == "Dust" || data == "Target" || data == "TobaccoWeight" || data == "RipperInfo" || data == "CRRConversion")
                {
                    ICollection<QueryFilter> filters = new List<QueryFilter>();
                    filters.Add(new QueryFilter("IsDeleted", "0"));
                    if (data == "ProdVol" || data == "DIM" || data == "Yield" || data == "CRR" || data == "WorkHour" || data == "TobaccoWeight" || data == "RipperInfo" || data == "CRRConversion")
                    {
                        if (year == 0)
                            year = DateTime.Now.Year;

                        filters.Add(new QueryFilter("Year", year));
                        if (week != "")
                            filters.Add(new QueryFilter("Week", week));

                        Sheet.Cells[8, 1].Value = "Year";
                        Sheet.Cells[9, 1].Value = "Week";

                        Sheet.Cells[8, 2].Value = year.ToString();
                        if (week == "")
                            Sheet.Cells[9, 2].Value = "All Week";
                        else
                            Sheet.Cells[9, 2].Value = (Int32.Parse(week)).ToString("00");

                        startrow = 12;
                        numrow = 12;
                    } else if (data == "StickPerPack" || data == "Dust" || data == "Target")
                    {
                        Sheet.Cells[8, 1].Value = "Production Center";

                        if (PC == "")
                            Sheet.Cells[8, 2].Value = "ALL";
                        else
                            Sheet.Cells[8, 2].Value = PC;

                        startrow = 10;
                        numrow = 10;
                    }
                    

                    
                    numcol = 1;

                    if (data == "ProdVol")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Order Number";
                        Sheet.Cells[numrow, numcol++].Value = "Material";
                        Sheet.Cells[numrow, numcol++].Value = "Material Description";
                        Sheet.Cells[numrow, numcol++].Value = "Order Type";
                        Sheet.Cells[numrow, numcol++].Value = "Resource";
                        Sheet.Cells[numrow, numcol++].Value = "Target Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Confirmed Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Balance";
                        Sheet.Cells[numrow, numcol++].Value = "Base Unit";
                        Sheet.Cells[numrow, numcol++].Value = "Start Date";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIProdVolAppService.Find(filters).DeserializeToReportKPIProdVolModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.OrderNumber;
                            Sheet.Cells[numrow, numcol++].Value = dt.Material;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialDescription;
                            Sheet.Cells[numrow, numcol++].Value = dt.OrderType;
                            Sheet.Cells[numrow, numcol++].Value = dt.Resource;
                            Sheet.Cells[numrow, numcol++].Value = dt.TargetQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.ConfirmedQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.Balance;
                            Sheet.Cells[numrow, numcol++].Value = dt.BaseUnit;
                            Sheet.Cells[numrow, numcol++].Value = dt.StartDate.ToString("dd-MMM-yy");
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "DIM")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Company Code";
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Issuing Plant";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Process Order";
                        Sheet.Cells[numrow, numcol++].Value = "Component";
                        Sheet.Cells[numrow, numcol++].Value = "Material Description";
                        Sheet.Cells[numrow, numcol++].Value = "Base Unit";
                        Sheet.Cells[numrow, numcol++].Value = "Origin Group";
                        Sheet.Cells[numrow, numcol++].Value = "Description";
                        Sheet.Cells[numrow, numcol++].Value = "TMC";
                        Sheet.Cells[numrow, numcol++].Value = "Standard Price";
                        Sheet.Cells[numrow, numcol++].Value = "Currency";
                        Sheet.Cells[numrow, numcol++].Value = "Lead Material";
                        Sheet.Cells[numrow, numcol++].Value = "Lead Material Desc";
                        Sheet.Cells[numrow, numcol++].Value = "GR Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "GR Qty With Addbacks";
                        Sheet.Cells[numrow, numcol++].Value = "BaseUnit";
                        Sheet.Cells[numrow, numcol++].Value = "Std Usg Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Std Usg Val";
                        Sheet.Cells[numrow, numcol++].Value = "Waste Factor Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Waste Factor Value";
                        Sheet.Cells[numrow, numcol++].Value = "Std Usg Qty Waste";
                        Sheet.Cells[numrow, numcol++].Value = "Std Usg Val Waste";
                        Sheet.Cells[numrow, numcol++].Value = "Goods Issue Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Goods Issue Value";
                        Sheet.Cells[numrow, numcol++].Value = "Stock Take Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Stock Take Value";
                        Sheet.Cells[numrow, numcol++].Value = "Tot Actual Cons Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Tot Actual Cons Val";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Qty %";
                        Sheet.Cells[numrow, numcol++].Value = "Usg Dev Val";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Val %";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Qty Waste";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Qty Waste %";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Val Waste";
                        Sheet.Cells[numrow, numcol++].Value = "Use Dev Val Waste %";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIDIMAppService.Find(filters).DeserializeToReportKPIDIMModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;

                            Sheet.Cells[numrow, numcol++].Value = dt.CompanyCode;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.IssuingPlant;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProcessOrder;
                            Sheet.Cells[numrow, numcol++].Value = dt.Component;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialDescription;
                            Sheet.Cells[numrow, numcol++].Value = dt.BaseUnit;
                            Sheet.Cells[numrow, numcol++].Value = dt.OriginGroup;
                            Sheet.Cells[numrow, numcol++].Value = dt.Description;
                            Sheet.Cells[numrow, numcol++].Value = dt.TMC;
                            Sheet.Cells[numrow, numcol++].Value = dt.StandardPrice;
                            Sheet.Cells[numrow, numcol++].Value = dt.Currency;
                            Sheet.Cells[numrow, numcol++].Value = dt.LeadMaterial;
                            Sheet.Cells[numrow, numcol++].Value = dt.LeadMaterialDesc;
                            Sheet.Cells[numrow, numcol++].Value = dt.GRQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.GRQtyWithAddbacks;
                            Sheet.Cells[numrow, numcol++].Value = dt.BaseUnit2;
                            Sheet.Cells[numrow, numcol++].Value = dt.StdUsgQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.StdUsgVal;
                            Sheet.Cells[numrow, numcol++].Value = dt.WasteFactorQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.WasteFactorValue;
                            Sheet.Cells[numrow, numcol++].Value = dt.StdUsgQtyWaste;
                            Sheet.Cells[numrow, numcol++].Value = dt.StdUsgValWaste;
                            Sheet.Cells[numrow, numcol++].Value = dt.GoodsIssueQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.GoodsIssueValue;
                            Sheet.Cells[numrow, numcol++].Value = dt.StockTakeQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.StockTakeValue;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotActualConsQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotActualConsVal;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevQtyPercent;
                            Sheet.Cells[numrow, numcol++].Value = dt.UsgDevVal;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevValPercent;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevQtyWaste;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevQtyWastePercent;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevValWaste;
                            Sheet.Cells[numrow, numcol++].Value = dt.UseDevValWastePercent;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "Yield")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Material Group";
                        Sheet.Cells[numrow, numcol++].Value = "Material";
                        Sheet.Cells[numrow, numcol++].Value = "Material Description";
                        Sheet.Cells[numrow, numcol++].Value = "Base Unit";
                        Sheet.Cells[numrow, numcol++].Value = "Category";
                        Sheet.Cells[numrow, numcol++].Value = "Material Group Desc";
                        Sheet.Cells[numrow, numcol++].Value = "Goods Receipt Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Goods Issue Qty";
                        Sheet.Cells[numrow, numcol++].Value = "WIP Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Stock Take Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Scrap Qty";
                        Sheet.Cells[numrow, numcol++].Value = "Tot TKG GI";
                        Sheet.Cells[numrow, numcol++].Value = "Tot TKG GR";
                        Sheet.Cells[numrow, numcol++].Value = "Tot TKG Non BOM";
                        Sheet.Cells[numrow, numcol++].Value = "Tot TKG WIP";
                        Sheet.Cells[numrow, numcol++].Value = "Yield %";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIYieldAppService.Find(filters).DeserializeToReportKPIYieldModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialGroup;
                            Sheet.Cells[numrow, numcol++].Value = dt.Material;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialDescription;
                            Sheet.Cells[numrow, numcol++].Value = dt.BaseUnit;
                            Sheet.Cells[numrow, numcol++].Value = dt.Category;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialGroupDesc;
                            Sheet.Cells[numrow, numcol++].Value = dt.GoodsReceiptQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.GoodsIssueQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.WIPQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.StockTakeQuantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.ScrapQty;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotTKGGI;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotTKGGR;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotTKGNonBOM;
                            Sheet.Cells[numrow, numcol++].Value = dt.TotTKGWIP;
                            Sheet.Cells[numrow, numcol++].Value = dt.YieldPercent;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "CRR")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Posting Date";
                        Sheet.Cells[numrow, numcol++].Value = "Machine";
                        Sheet.Cells[numrow, numcol++].Value = "Machine Type";
                        Sheet.Cells[numrow, numcol++].Value = "Shift";
                        Sheet.Cells[numrow, numcol++].Value = "Order Number";
                        Sheet.Cells[numrow, numcol++].Value = "PO Lead Material";
                        Sheet.Cells[numrow, numcol++].Value = "Reject Material";
                        Sheet.Cells[numrow, numcol++].Value = "Material Description";
                        Sheet.Cells[numrow, numcol++].Value = "Quantity";
                        Sheet.Cells[numrow, numcol++].Value = "Base Unit";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPICRRAppService.Find(filters).DeserializeToReportKPICRRModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.PostingDate.ToString("dd-MMM-yy");
                            Sheet.Cells[numrow, numcol++].Value = dt.Machine;
                            Sheet.Cells[numrow, numcol++].Value = dt.MachineType;
                            Sheet.Cells[numrow, numcol++].Value = dt.Shift;
                            Sheet.Cells[numrow, numcol++].Value = dt.OrderNumber;
                            Sheet.Cells[numrow, numcol++].Value = dt.POLeadMaterial;
                            Sheet.Cells[numrow, numcol++].Value = dt.RejectMaterial;
                            Sheet.Cells[numrow, numcol++].Value = dt.MaterialDescription;
                            Sheet.Cells[numrow, numcol++].Value = dt.Quantity;
                            Sheet.Cells[numrow, numcol++].Value = dt.BaseUnit;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "WorkHour")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Packer";
                        Sheet.Cells[numrow, numcol++].Value = "Work Hour";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIWorkHourAppService.Find(filters).DeserializeToReportKPIWorkHourModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.Packer;
                            Sheet.Cells[numrow, numcol++].Value = dt.WorkHour;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "StickPerPack")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Packer";
                        Sheet.Cells[numrow, numcol++].Value = "Stick per Pack";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIStickPerPackAppService.Find(filters).DeserializeToReportKPIStickPerPackModelList();
                        if (PC != "")
                            datalist = datalist.Where(x => x.Packer.Contains(PC)).ToList();

                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Packer;
                            Sheet.Cells[numrow, numcol++].Value = dt.Stick;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "Dust")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Dust";
                        Sheet.Cells[numrow, numcol++].Value = "Winnower";
                        Sheet.Cells[numrow, numcol++].Value = "Floor Sweeping";
                        Sheet.Cells[numrow, numcol++].Value = "RS";
                        Sheet.Cells[numrow, numcol++].Value = "GR Spec";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIDustAppService.Find(filters).DeserializeToReportKPIDustModelList();
                        if (PC != "")
                            datalist = datalist.Where(x => x.ProductionCenter.Contains(PC)).ToList();

                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Dust;
                            Sheet.Cells[numrow, numcol++].Value = dt.Winnower;
                            Sheet.Cells[numrow, numcol++].Value = dt.FloorSweeping;
                            Sheet.Cells[numrow, numcol++].Value = dt.RS;
                            Sheet.Cells[numrow, numcol++].Value = dt.GRSpec;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "Target")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "KPI";
                        Sheet.Cells[numrow, numcol++].Value = "Internal Target";
                        Sheet.Cells[numrow, numcol++].Value = "OB Target";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPITargetAppService.Find(filters).DeserializeToReportKPITargetModelList();
                        if (PC != "")
                            datalist = datalist.Where(x => x.ProductionCenter.Contains(PC)).ToList();

                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.KPI;
                            Sheet.Cells[numrow, numcol++].Value = dt.TargetInternal;
                            Sheet.Cells[numrow, numcol++].Value = dt.TargetOB;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "TobaccoWeight")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Tobacco Weight Mean";
                        Sheet.Cells[numrow, numcol++].Value = "Tobacco Weight Stdev";
                        Sheet.Cells[numrow, numcol++].Value = "MC Mean";
                        Sheet.Cells[numrow, numcol++].Value = "MC Stdev";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPITobaccoWeightAppService.Find(filters).DeserializeToReportKPITobaccoWeightModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.Mean / 1000000000000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.Stdev / 1000000000000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.MeanMC / 10000000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.StdevMC / 10000000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "RipperInfo")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Material";
                        Sheet.Cells[numrow, numcol++].Value = "OrderNum";
                        Sheet.Cells[numrow, numcol++].Value = "Actual Start Date";
                        Sheet.Cells[numrow, numcol++].Value = "Description";
                        Sheet.Cells[numrow, numcol++].Value = "Qty Iss";
                        Sheet.Cells[numrow, numcol++].Value = "Qty Rec";
                        Sheet.Cells[numrow, numcol++].Value = "Yield";
                        Sheet.Cells[numrow, numcol++].Value = "Val Issued";
                        Sheet.Cells[numrow, numcol++].Value = "Val Receiv";
                        Sheet.Cells[numrow, numcol++].Value = "Val Differ";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPIRipperInfoAppService.Find(filters).DeserializeToReportKPIRipperInfoModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.Material;
                            Sheet.Cells[numrow, numcol++].Value = dt.OrderNum;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ActualStartDate).ToString("dd-MMM-yy");
                            Sheet.Cells[numrow, numcol++].Value = dt.Description;
                            Sheet.Cells[numrow, numcol++].Value = dt.QtyIss;
                            Sheet.Cells[numrow, numcol++].Value = dt.QtyRec;
                            Sheet.Cells[numrow, numcol++].Value = dt.Yield;
                            Sheet.Cells[numrow, numcol++].Value = dt.ValIssued;
                            Sheet.Cells[numrow, numcol++].Value = dt.ValReceiv;
                            Sheet.Cells[numrow, numcol++].Value = dt.ValDiffer;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                    else if (data == "CRRConversion")
                    {
                        Sheet.Cells[numrow, numcol++].Value = "Production Center";
                        Sheet.Cells[numrow, numcol++].Value = "Year";
                        Sheet.Cells[numrow, numcol++].Value = "Week";
                        Sheet.Cells[numrow, numcol++].Value = "Debu Final";
                        Sheet.Cells[numrow, numcol++].Value = "Dust Halus LM";
                        Sheet.Cells[numrow, numcol++].Value = "Dust Halus LM no Ripper";
                        Sheet.Cells[numrow, numcol++].Value = "Sapon Tembakau";
                        Sheet.Cells[numrow, numcol++].Value = "Sapon Cigarette";
                        Sheet.Cells[numrow, numcol++].Value = "Claimable In Th Stick";
                        Sheet.Cells[numrow, numcol++].Value = "Average Final Ov";
                        Sheet.Cells[numrow, numcol++].Value = "C2 Weight";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted By";
                        Sheet.Cells[numrow, numcol++].Value = "Inserted Date";

                        var datalist = _reportKPICRRConversionAppService.Find(filters).DeserializeToReportKPICRRConversionModelList();
                        foreach (var dt in datalist)
                        {
                            numrow++;
                            numcol = 1;
                            Sheet.Cells[numrow, numcol++].Value = dt.ProductionCenter;
                            Sheet.Cells[numrow, numcol++].Value = dt.Year;
                            Sheet.Cells[numrow, numcol++].Value = dt.Week;
                            Sheet.Cells[numrow, numcol++].Value = dt.DebuFinal / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.DustHalusLM / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.DustHalusLMnoRipper / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.SaponTembakau / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.SaponCigarette / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.ClaimableInThStick / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.AverageFinalOv / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.C2Weight / 1000000;
                            Sheet.Cells[numrow, numcol++].Value = dt.ModifiedBy;
                            Sheet.Cells[numrow, numcol++].Value = ((DateTime)dt.ModifiedDate).ToString("dd-MMM-yy HH:mm");
                        }
                    }
                }

                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(HEADER_COLOR);

                Sheet.Cells[startrow, 1, numrow, numcol--].AutoFitColumns();
                using (var range = Sheet.Cells[startrow, 1, startrow, numcol])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                }

                Sheet.Cells[startrow, 1, numrow, numcol].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename="+ "ReportKPI - Master Data - " + data + ".xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult FormUpload(string ID)
		{
			var model = new ReportKPIFormModel();
            model.Type = ID;
			model.Year = DateTime.Now.Year;
			model.Week = GetWeekNumber(DateTime.Now).ToString("00");

			return PartialView(model);
		}

		[HttpPost]
		public ActionResult FormUpload(ReportKPIFormModel FormData)
		{
			IExcelDataReader reader = null;

			try
			{
				if (!ModelState.IsValid)
				{
					Session["ResultLog"] = "error_Data not valid";
					return RedirectToAction("Index");
				}

				if (FormData.Excel != null && FormData.Excel.ContentLength > 0)
				{
					Stream stream = FormData.Excel.InputStream;

					if (FormData.Excel.FileName.ToLower().EndsWith(".xls"))
					{
						reader = ExcelReaderFactory.CreateBinaryReader(stream);
					}
					else if (FormData.Excel.FileName.ToLower().EndsWith(".xlsx"))
					{
						reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
					}
					else
					{
						Session["ResultLog"] = "error_File not supported";
						return RedirectToAction("Index");
					}

					int fieldcount = reader.FieldCount;
					int rowcount = reader.RowCount;
					DataTable dt = new DataTable();
					DataTable dt_ = reader.AsDataSet().Tables[0];

					try
					{
						if (FormData.Type == "ProdVol")
						{
							var modelList = new List<ReportKPIProdVolModel>();

							for (int index = 1; index < dt_.Rows.Count; index++)
							{
								DateTime DataTime;
								if (dt_.Rows[index][0].ToString() != "" && DateTime.TryParse(dt_.Rows[index][9].ToString(), out DataTime))
								{
									var model = new ReportKPIProdVolModel();

									model.Year = DataTime.Year;
									model.Week = GetWeekNumber(DataTime);

									var col = 0;
									model.OrderNumber = dt_.Rows[index][col++].ToString();
									model.Material = dt_.Rows[index][col++].ToString();
									model.MaterialDescription = dt_.Rows[index][col++].ToString();
									model.OrderType = dt_.Rows[index][col++].ToString();
									model.Resource = dt_.Rows[index][col++].ToString();
                                    model.ProductionCenter = System.Text.RegularExpressions.Regex.Replace(model.Resource.Substring(model.Resource.Length - 4), @"[\d-]", string.Empty);
                                    model.TargetQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.ConfirmedQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.Balance = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.BaseUnit = dt_.Rows[index][col++].ToString();
									model.StartDate = (DateTime)dt_.Rows[index][col++];

									model.ModifiedBy = AccountName;
									model.ModifiedDate = DateTime.Now;

									modelList.Add(model);
								}
							}

							if (modelList.Count > 0)
							{
                                long lala = 0;
								foreach (var data in modelList)
								{
									string RawData = JsonHelper<ReportKPIProdVolModel>.Serialize(data);
									lala = _reportKPIProdVolAppService.Add(RawData);
								}

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
							else
							{
								Session["ResultLog"] = "error_Data is empty or columns does not match the example";
							}
						}
						else if (FormData.Type == "DIM")
						{
							var modelList = new List<ReportKPIDIMModel>();

							for (int index = 1; index < dt_.Rows.Count; index++)
							{
								int temp;
								if (Int32.TryParse(dt_.Rows[index][0].ToString(), out temp))
								{
									var model = new ReportKPIDIMModel();

									model.Year = FormData.Year;
									model.Week = Int32.Parse(FormData.Week);

									var col = 0;
									model.CompanyCode = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.ProductionCenter = dt_.Rows[index][col++].ToString();
									model.IssuingPlant = dt_.Rows[index][col++].ToString();
									model.ProcessOrder = dt_.Rows[index][col++].ToString();
									model.Component = dt_.Rows[index][col++].ToString();
									model.MaterialDescription = dt_.Rows[index][col++].ToString();
									model.BaseUnit = dt_.Rows[index][col++].ToString();
									model.OriginGroup = dt_.Rows[index][col++].ToString();
									model.Description = dt_.Rows[index][col++].ToString();
									model.TMC = dt_.Rows[index][col++].ToString();
									model.StandardPrice = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.Currency = dt_.Rows[index][col++].ToString();
									model.LeadMaterial = dt_.Rows[index][col++].ToString();
									model.LeadMaterialDesc = dt_.Rows[index][col++].ToString();
									model.GRQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.GRQtyWithAddbacks = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.BaseUnit2 = dt_.Rows[index][col++].ToString();
									model.StdUsgQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StdUsgVal = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.WasteFactorQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.WasteFactorValue = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StdUsgQtyWaste = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StdUsgValWaste = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.GoodsIssueQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.GoodsIssueValue = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StockTakeQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StockTakeValue = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotActualConsQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotActualConsVal = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevQtyPercent = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UsgDevVal = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevValPercent = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevQtyWaste = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevQtyWastePercent = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevValWaste = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.UseDevValWastePercent = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));


									model.ModifiedBy = AccountName;
									model.ModifiedDate = DateTime.Now;

									modelList.Add(model);
								}
							}

							if (modelList.Count > 0)
							{
                                long lala = 0;
								foreach (var data in modelList)
								{
									string RawData = JsonHelper<ReportKPIDIMModel>.Serialize(data);
									lala = _reportKPIDIMAppService.Add(RawData);
								}

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
							else
							{
								Session["ResultLog"] = "error_Data is empty or columns does not match the example";
							}
						}
						else if (FormData.Type == "Yield")
						{
							var modelList = new List<ReportKPIYieldModel>();

							for (int index = 1; index < dt_.Rows.Count; index++)
							{
								if (dt_.Rows[index][2].ToString() != "")
								{
									var model = new ReportKPIYieldModel();

									model.Year = FormData.Year;
									model.Week = Int32.Parse(FormData.Week);
                                    model.ProductionCenter = FormData.PC;

                                    var col = 0;
									model.MaterialGroup = dt_.Rows[index][col++].ToString();
									model.Material = dt_.Rows[index][col++].ToString();
									model.MaterialDescription = dt_.Rows[index][col++].ToString();
									model.BaseUnit = dt_.Rows[index][col++].ToString();
									model.Category = dt_.Rows[index][col++].ToString();
									model.MaterialGroupDesc = dt_.Rows[index][col++].ToString();
									model.GoodsReceiptQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.GoodsIssueQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.WIPQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.StockTakeQuantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.ScrapQty = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotTKGGI = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotTKGGR = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotTKGNonBOM = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.TotTKGWIP = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.YieldPercent = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

									model.ModifiedBy = AccountName;
									model.ModifiedDate = DateTime.Now;

									modelList.Add(model);
								}
							}
							if (modelList.Count > 0)
							{
                                long lala = 0;
                                var counter = 0;
								var MaterialGroup = "";
								foreach (var data in modelList)
								{
									if (counter == 0 || (data.MaterialGroup != "" && MaterialGroup != data.MaterialGroup))
									{
										MaterialGroup = data.MaterialGroup;
									}
									else
									{
										data.MaterialGroup = MaterialGroup;
									}

									string RawData = JsonHelper<ReportKPIYieldModel>.Serialize(data);
									lala = _reportKPIYieldAppService.Add(RawData);

									counter++;
								}

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                } else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
							}
							else
							{
								Session["ResultLog"] = "error_Data is empty or columns does not match the example";
							}
						}
						else if (FormData.Type == "CRR")
						{
							var modelList = new List<ReportKPICRRModel>();

							for (int index = 1; index < dt_.Rows.Count; index++)
							{
								DateTime DataTime;
								if (DateTime.TryParse(dt_.Rows[index][0].ToString(), out DataTime))
								{
									var model = new ReportKPICRRModel();

                                    model.Year = FormData.Year;
                                    model.Week = Int32.Parse(FormData.Week);

                                    var col = 0;
									model.PostingDate = (DateTime)dt_.Rows[index][col++];
									model.Machine = dt_.Rows[index][col++].ToString();
                                    model.ProductionCenter = System.Text.RegularExpressions.Regex.Replace(model.Machine.Substring(model.Machine.Length - 4), @"[\d-]", string.Empty);
                                    model.MachineType = dt_.Rows[index][col++].ToString();
									model.Shift = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.OrderNumber = dt_.Rows[index][col++].ToString();
									model.POLeadMaterial = dt_.Rows[index][col++].ToString();
									model.RejectMaterial = dt_.Rows[index][col++].ToString();
									model.MaterialDescription = dt_.Rows[index][col++].ToString();
									model.Quantity = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
									model.BaseUnit = dt_.Rows[index][col++].ToString();

									model.ModifiedBy = AccountName;
									model.ModifiedDate = DateTime.Now;

									modelList.Add(model);
								}
							}

							if (modelList.Count > 0)
							{
                                long lala = 0;

								foreach (var data in modelList)
								{
									string RawData = JsonHelper<ReportKPICRRModel>.Serialize(data);
									lala = _reportKPICRRAppService.Add(RawData);
								}

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
							else
							{
								Session["ResultLog"] = "error_Data is empty or columns does not match the example";
							}
						}
                        else if (FormData.Type == "WorkHour")
                        {
                            var modelList = new List<ReportKPIWorkHourModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (int.TryParse(dt_.Rows[index][0].ToString(), out int n) && dt_.Rows[index][2].ToString() != "")
                                {
                                    var model = new ReportKPIWorkHourModel();

                                    var col = 0;

                                    model.Year = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Week = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Packer = dt_.Rows[index][col++].ToString();
                                    model.ProductionCenter = model.Packer.Substring(0, 2);
                                    model.WorkHour = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPIWorkHourAppService.GetAll(true).DeserializeToReportKPIWorkHourModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.Week == data.Week && x.Packer == data.Packer && x.Year == data.Year).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPIWorkHourAppService.GetById(check[0].ID, true).DeserializeToReportKPIWorkHourModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;
                                        exist.WorkHour = data.WorkHour;

                                        string update = JsonHelper<ReportKPIWorkHourModel>.Serialize(exist);
                                        _reportKPIWorkHourAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPIWorkHourModel>.Serialize(data);
                                        lala = _reportKPIWorkHourAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "StickPerPack")
                        {
                            var modelList = new List<ReportKPIStickPerPackModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (dt_.Rows[index][0].ToString() != "" && int.TryParse(dt_.Rows[index][1].ToString(), out int n))
                                {
                                    var model = new ReportKPIStickPerPackModel();

                                    var col = 0;
                                    model.Packer = dt_.Rows[index][col++].ToString();
                                    model.ProductionCenter = model.Packer.Substring(0, 2);
                                    model.Stick = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPIStickPerPackAppService.GetAll(true).DeserializeToReportKPIStickPerPackModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.Packer == data.Packer).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPIStickPerPackAppService.GetById(check[0].ID, true).DeserializeToReportKPIStickPerPackModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;
                                        exist.Stick = data.Stick;

                                        string update = JsonHelper<ReportKPIStickPerPackModel>.Serialize(exist);
                                        _reportKPIStickPerPackAppService.Update(update);
                                        lala = 1;
                                    } else
                                    {
                                        string RawData = JsonHelper<ReportKPIStickPerPackModel>.Serialize(data);
                                        lala = _reportKPIStickPerPackAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "Dust")
                        {
                            var modelList = new List<ReportKPIDustModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (dt_.Rows[index][0].ToString() != "" && int.TryParse(dt_.Rows[index][1].ToString(), out int n))
                                {
                                    var model = new ReportKPIDustModel();

                                    var col = 0;
                                    model.ProductionCenter = dt_.Rows[index][col++].ToString();
                                    model.Dust = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Winnower = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.FloorSweeping = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.RS = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.GRSpec = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPIDustAppService.GetAll(true).DeserializeToReportKPIDustModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.ProductionCenter == data.ProductionCenter).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPIDustAppService.GetById(check[0].ID, true).DeserializeToReportKPIDustModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;
                                        exist.Dust = data.Dust;
                                        exist.Winnower = data.Winnower;
                                        exist.FloorSweeping = data.FloorSweeping;
                                        exist.RS = data.RS;
                                        exist.GRSpec = data.GRSpec;

                                        string update = JsonHelper<ReportKPIDustModel>.Serialize(exist);
                                        _reportKPIDustAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPIDustModel>.Serialize(data);
                                        lala = _reportKPIDustAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "Target")
                        {
                            var modelList = new List<ReportKPITargetModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (dt_.Rows[index][0].ToString() != "" && dt_.Rows[index][1].ToString() != "")
                                {
                                    var model = new ReportKPITargetModel();

                                    var col = 0;
                                    model.ProductionCenter = dt_.Rows[index][col++].ToString();
                                    model.KPI = dt_.Rows[index][col++].ToString();
                                    model.TargetInternal = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.TargetOB = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    if (model.ProductionCenter == "ID" || model.ProductionCenter == "PB" || model.ProductionCenter == "PJ" || model.ProductionCenter == "PK" || model.ProductionCenter == "PI")
                                        modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPITargetAppService.GetAll(true).DeserializeToReportKPITargetModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.ProductionCenter == data.ProductionCenter).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPITargetAppService.GetById(check[0].ID, true).DeserializeToReportKPITargetModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;
                                        exist.ProductionCenter = data.ProductionCenter;
                                        exist.KPI = data.KPI;
                                        exist.TargetInternal = data.TargetInternal;
                                        exist.TargetOB = data.TargetOB;

                                        string update = JsonHelper<ReportKPITargetModel>.Serialize(exist);
                                        _reportKPITargetAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPITargetModel>.Serialize(data);
                                        lala = _reportKPITargetAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "TobaccoWeight")
                        {
                            var modelList = new List<ReportKPITobaccoWeightModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (int.TryParse(dt_.Rows[index][1].ToString(), out int n) && int.TryParse(dt_.Rows[index][2].ToString(), out int m))
                                {
                                    var model = new ReportKPITobaccoWeightModel();

                                    var col = 0;

                                    model.ProductionCenter = dt_.Rows[index][col++].ToString();
                                    model.Year = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Week = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Mean = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Stdev = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.MeanMC = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.StdevMC = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.Mean = Math.Round((model.Mean * 1000000000000000), 3);
                                    model.Stdev = Math.Round((model.Stdev * 1000000000000000), 3);
                                    model.MeanMC = Math.Round((model.MeanMC * 10000000000), 3);
                                    model.StdevMC = Math.Round((model.StdevMC * 10000000000), 3);

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPITobaccoWeightAppService.GetAll(true).DeserializeToReportKPITobaccoWeightModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.Week == data.Week && x.ProductionCenter == data.ProductionCenter && x.Year == data.Year).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPITobaccoWeightAppService.GetById(check[0].ID, true).DeserializeToReportKPITobaccoWeightModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;
                                        exist.Mean = data.Mean;
                                        exist.Stdev = data.Stdev;
                                        exist.MeanMC = data.MeanMC;
                                        exist.StdevMC = data.StdevMC;

                                        string update = JsonHelper<ReportKPITobaccoWeightModel>.Serialize(exist);
                                        _reportKPITobaccoWeightAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPITobaccoWeightModel>.Serialize(data);
                                        lala = _reportKPITobaccoWeightAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "RipperInfo")
                        {
                            var modelList = new List<ReportKPIRipperInfoModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (int.TryParse(dt_.Rows[index][1].ToString(), out int n) && int.TryParse(dt_.Rows[index][2].ToString(), out int m) && DateTime.TryParse(dt_.Rows[index][5].ToString(), out DateTime DataTime))
                                {
                                    var model = new ReportKPIRipperInfoModel();

                                    var col = 0;
                                    model.ProductionCenter = dt_.Rows[index][col++].ToString();
                                    model.Year = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Week = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Material = dt_.Rows[index][col++].ToString();
                                    model.OrderNum = dt_.Rows[index][col++].ToString();
                                    model.ActualStartDate = (DateTime)dt_.Rows[index][col++];
                                    model.Description = dt_.Rows[index][col++].ToString();
                                    model.QtyIss = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.QtyRec = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Yield = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.ValIssued = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.ValReceiv = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.ValDiffer = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPIRipperInfoAppService.GetAll(true).DeserializeToReportKPIRipperInfoModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.ProductionCenter == data.ProductionCenter && x.Week == data.Week && x.OrderNum == data.OrderNum && x.Year == data.Year).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPIRipperInfoAppService.GetById(check[0].ID, true).DeserializeToReportKPIRipperInfoModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;

                                        exist.QtyIss = data.QtyIss;
                                        exist.QtyRec = data.QtyRec;
                                        exist.Yield = data.Yield;
                                        exist.ValIssued = data.ValIssued;
                                        exist.ValReceiv = data.ValReceiv;
                                        exist.ValDiffer = data.ValDiffer;

                                        string update = JsonHelper<ReportKPIRipperInfoModel>.Serialize(exist);
                                        _reportKPIRipperInfoAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPIRipperInfoModel>.Serialize(data);
                                        lala = _reportKPIRipperInfoAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else if (FormData.Type == "CRRConversion")
                        {
                            var modelList = new List<ReportKPICRRConversionModel>();

                            for (int index = 1; index < dt_.Rows.Count; index++)
                            {
                                if (int.TryParse(dt_.Rows[index][1].ToString(), out int n) && int.TryParse(dt_.Rows[index][2].ToString(), out int m))
                                {
                                    var model = new ReportKPICRRConversionModel();
                                    var col = 0;

                                    model.ProductionCenter = dt_.Rows[index][col++].ToString();
                                    model.Year = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.Week = Int32.Parse((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.DebuFinal = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.DustHalusLM = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.DustHalusLMnoRipper = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.SaponTembakau = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.SaponCigarette = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.ClaimableInThStick = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.AverageFinalOv = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));
                                    model.C2Weight = Convert.ToDecimal((dt_.Rows[index][col].ToString() == "" ? "0" : dt_.Rows[index][col++].ToString()).Replace(",", ""));

                                    model.DebuFinal = Math.Round((model.DebuFinal * 1000000), 3);
                                    model.DustHalusLM = Math.Round((model.DustHalusLM * 1000000), 3);
                                    model.DustHalusLMnoRipper = Math.Round((model.DustHalusLMnoRipper * 1000000), 3);
                                    model.SaponTembakau = Math.Round((model.SaponTembakau * 1000000), 3);
                                    model.SaponCigarette = Math.Round((model.SaponCigarette * 1000000), 3);
                                    model.ClaimableInThStick = Math.Round((model.ClaimableInThStick * 1000000), 3);
                                    model.AverageFinalOv = Math.Round((model.AverageFinalOv * 1000000), 3);
                                    model.C2Weight = Math.Round((model.C2Weight * 1000000), 3);

                                    model.ModifiedBy = AccountName;
                                    model.ModifiedDate = DateTime.Now;

                                    modelList.Add(model);
                                }
                            }

                            if (modelList.Count > 0)
                            {
                                long lala = 0;

                                var allData = _reportKPICRRConversionAppService.GetAll(true).DeserializeToReportKPICRRConversionModelList();

                                foreach (var data in modelList)
                                {
                                    var check = allData.Where(x => x.ProductionCenter == data.ProductionCenter && x.Week == data.Week && x.Year == data.Year).ToList();

                                    if (check != null && check.Count() > 0)
                                    {
                                        var exist = _reportKPICRRConversionAppService.GetById(check[0].ID, true).DeserializeToReportKPICRRConversionModel();
                                        exist.ModifiedBy = AccountName;
                                        exist.ModifiedDate = DateTime.Now;

                                        exist.DebuFinal = data.DebuFinal;
                                        exist.DustHalusLM = data.DustHalusLM;
                                        exist.DustHalusLMnoRipper = data.DustHalusLMnoRipper;
                                        exist.SaponTembakau = data.SaponTembakau;
                                        exist.SaponCigarette = data.SaponCigarette;
                                        exist.ClaimableInThStick = data.ClaimableInThStick;
                                        exist.AverageFinalOv = data.AverageFinalOv;
                                        exist.C2Weight = data.C2Weight;

                                        string update = JsonHelper<ReportKPICRRConversionModel>.Serialize(exist);
                                        _reportKPICRRConversionAppService.Update(update);
                                        lala = 1;
                                    }
                                    else
                                    {
                                        string RawData = JsonHelper<ReportKPICRRConversionModel>.Serialize(data);
                                        lala = _reportKPICRRConversionAppService.Add(RawData);
                                    }
                                }

                                if (lala > 0)
                                {
                                    Session["ResultLog"] = "success_Upload Succeed";
                                }
                                else
                                {
                                    Session["ResultLog"] = "error_Failed to save to DB";
                                }
                            }
                            else
                            {
                                Session["ResultLog"] = "error_Data is empty or columns does not match the example";
                            }
                        }
                        else
                        {
							Session["ResultLog"] = "error_Data type not exist";
						}
					}
					catch(Exception e)
					{
						Session["ResultLog"] = "error_Wrong data type, please check data type & download example file";
					}

					reader.Close();
					reader.Dispose();
				}
				else
				{
					Session["ResultLog"] = "error_File Corrupted";
				}
			}
			catch (Exception ex)
			{
				Session["ResultLog"] = "error_Upload Failed";
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				reader = null;
			}

            return RedirectToAction("Index", new { page = "Data"+FormData.Type });
		}

        public int GetWeekNumber(DateTime date)
		{
			CultureInfo ciCurr = CultureInfo.CurrentCulture;
			int weekNum = ciCurr.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
			return weekNum;
		}

        public decimal CalculateLinest(decimal[] y, decimal[] x)
        {
            decimal linest = 0;
            if (y.Length == x.Length)
            {
                decimal avgY = y.Average();
                decimal avgX = x.Average();
                decimal[] dividend = new decimal[y.Length];
                decimal[] divisor = new decimal[y.Length];
                for (int i = 0; i < y.Length; i++)
                {
                    dividend[i] = (x[i] - avgX) * (y[i] - avgY);
                    double temp = Math.Pow( (double)(x[i] - avgX), 2);
                    divisor[i] = (decimal)temp;
                }
                linest = dividend.Sum() / divisor.Sum();
            }
            return linest;
        }

        public decimal getStandardDeviation(List<decimal> doubleList)
        {
            decimal average = doubleList.Average();
            decimal sumOfDerivation = 0;
            foreach (decimal value in doubleList)
            {
                sumOfDerivation += (value) * (value);
            }
            double sumOfDerivationAverage = (double)sumOfDerivation / (doubleList.Count - 1);
            return (decimal)Math.Sqrt(sumOfDerivationAverage - (double)(average * average));
        }


        public static List<SelectListItem> BindDropDownDataType()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();

			_menuList.Add(new SelectListItem
			{
				Text = "Production Volume",
				Value = "ProdVol"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Yield",
				Value = "Yield"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "DIM",
				Value = "DIM"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Uptime CRR",
				Value = "CRR"
			});

			return _menuList;

		}
	}
}
