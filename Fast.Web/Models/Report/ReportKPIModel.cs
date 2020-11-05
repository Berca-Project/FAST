using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class ReportWeeklyKPIModel
    {
        public List<string> Header { get; set; }
        public List<decimal> Volume { get; set; }
        public List<decimal> Uptime { get; set; }
        public List<decimal> CRR { get; set; }
        public List<decimal> SpecYield { get; set; }
        public List<decimal> ActYield { get; set; }
        public List<decimal> DryYield { get; set; }
        public List<decimal> DIMWaste { get; set; }
        public List<decimal> Claimable { get; set; }
        public List<decimal> TaxStampWaste { get; set; }
        public List<decimal> SecActualDryYield { get; set; }
        public List<decimal> FloorSwept { get; set; }
        public List<decimal> DustHalus { get; set; }
        public List<decimal> RippingLoss { get; set; }
        public List<decimal> Unaccountable { get; set; }
        public List<decimal> FilterRod { get; set; }
        public List<decimal> HingeLidBlank { get; set; }
        public List<decimal> InnerLiner { get; set; }
        public List<decimal> TippingPaper { get; set; }
        public List<decimal> WrappingFilm { get; set; }
        public List<decimal> InnerFrame { get; set; }
        public List<decimal> CigarettePaper { get; set; }
        public List<decimal> UptimeTheo { get; set; }
        public List<decimal> ClaimableInThStick { get; set; }
        public List<decimal> ClaimableU16 { get; set; }
        public List<decimal> ClaimableU12 { get; set; }
        public List<decimal> ClaimableUCool { get; set; }
        public List<decimal> VolumeYield { get; set; }
    }
    public class ReportWeeklyYieldModel
    {
        public List<string> Header { get; set; }
        public List<decimal> SecSpecWetYield { get; set; }
        public List<decimal> SecActualWetYield { get; set; }
        public List<decimal> FloorSweptWet { get; set; }
        public List<decimal> DustHalusWet { get; set; }
        public List<decimal> RippingLossWet { get; set; }
        public List<decimal> UnaccountableWet { get; set; }
        public List<decimal> SecSpecDryYield { get; set; }
        public List<decimal> SecActualDryYield { get; set; }
        public List<decimal> FloorSweptDry { get; set; }
        public List<decimal> DustHalusDry { get; set; }
        public List<decimal> RippingLossDry { get; set; }
        public List<decimal> UnaccountableDry { get; set; }

        public List<decimal> MCCutfiller { get; set; }
        public List<decimal> PackOV { get; set; }
        public List<decimal> WeightATW { get; set; }
        public List<decimal> WeightC2Weight { get; set; }
    }

    public class TRSModel
    {
        public dynamic LPHID { get; set; }
        public dynamic Maker { get; set; }
        public dynamic Date { get; set; }
        public dynamic Week { get; set; }
        public dynamic Shift { get; set; }
        public dynamic Time { get; set; }
        public dynamic Speed { get; set; }
        public dynamic CTW { get; set; }
        public dynamic Sample { get; set; }
        public dynamic Sampling_time { get; set; }
        public dynamic Recovery { get; set; }
        public dynamic Remark { get; set; }
        public dynamic TRSRecoveryMaker { get; set; }
        public dynamic TRSRecoveryShift { get; set; }
        public dynamic TRSRecoveryDate { get; set; }
        public dynamic ConfidenceMaker { get; set; }
        public dynamic ConfidenceDate { get; set; }
}

    public class ReportTRSModel
    {
        public List<string> header { get; set; }
        public List<List<dynamic>> temp { get; set; }
        public List<TRSModel> content { get; set; }
    }

    public class ReportKPIDIMModel: BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public int CompanyCode { get; set; }
        public string ProductionCenter { get; set; }
        public string IssuingPlant { get; set; }
        public string ProcessOrder { get; set; }
        public string Component { get; set; }
        public string MaterialDescription { get; set; }
        public string BaseUnit { get; set; }
        public string OriginGroup { get; set; }
        public string Description { get; set; }
        public string TMC { get; set; }
        public decimal StandardPrice { get; set; }
        public string Currency { get; set; }
        public string LeadMaterial { get; set; }
        public string LeadMaterialDesc { get; set; }
        public decimal GRQuantity { get; set; }
        public decimal GRQtyWithAddbacks { get; set; }
        public string BaseUnit2 { get; set; }
        public decimal StdUsgQty { get; set; }
        public decimal StdUsgVal { get; set; }
        public decimal WasteFactorQty { get; set; }
        public decimal WasteFactorValue { get; set; }
        public decimal StdUsgQtyWaste { get; set; }
        public decimal StdUsgValWaste { get; set; }
        public decimal GoodsIssueQty { get; set; }
        public decimal GoodsIssueValue { get; set; }
        public decimal StockTakeQuantity { get; set; }
        public decimal StockTakeValue { get; set; }
        public decimal TotActualConsQty { get; set; }
        public decimal TotActualConsVal { get; set; }
        public decimal UseDevQty { get; set; }
        public decimal UseDevQtyPercent { get; set; }
        public decimal UsgDevVal { get; set; }
        public decimal UseDevValPercent { get; set; }
        public decimal UseDevQtyWaste { get; set; }
        public decimal UseDevQtyWastePercent { get; set; }
        public decimal UseDevValWaste { get; set; }
        public decimal UseDevValWastePercent { get; set; }
    }

    public class ReportKPICRRModel: BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public DateTime PostingDate { get; set; }
        public string Machine { get; set; }
        public string MachineType { get; set; }
        public int Shift { get; set; }
        public string OrderNumber { get; set; }
        public string POLeadMaterial { get; set; }
        public string RejectMaterial { get; set; }
        public string MaterialDescription { get; set; }
        public decimal Quantity { get; set; }
        public string BaseUnit { get; set; }
    }

    public class ReportKPIProdVolModel: BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string OrderNumber { get; set; }
        public string Material { get; set; }
        public string MaterialDescription { get; set; }
        public string OrderType { get; set; }
        public string Resource { get; set; }
        public decimal TargetQuantity { get; set; }
        public decimal ConfirmedQuantity { get; set; }
        public decimal Balance { get; set; }
        public string BaseUnit { get; set; }
        public DateTime StartDate { get; set; }
    }

    public class ReportKPIYieldModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string MaterialGroup { get; set; }
        public string Material { get; set; }
        public string MaterialDescription { get; set; }
        public string BaseUnit { get; set; }
        public string Category { get; set; }
        public string MaterialGroupDesc { get; set; }
        public decimal GoodsReceiptQty { get; set; }
        public decimal GoodsIssueQty { get; set; }
        public decimal WIPQuantity { get; set; }
        public decimal StockTakeQuantity { get; set; }
        public decimal ScrapQty { get; set; }
        public decimal TotTKGGI { get; set; }
        public decimal TotTKGGR { get; set; }
        public decimal TotTKGNonBOM { get; set; }
        public decimal TotTKGWIP { get; set; }
        public decimal YieldPercent { get; set; }
    }

    public class ReportKPIWorkHourModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string Packer { get; set; }
        public decimal WorkHour { get; set; }
    }

    public class ReportKPIStickPerPackModel : BaseModel
    {
        public string ProductionCenter { get; set; }
        public string Packer { get; set; }
        public int Stick { get; set; }
    }

    public class ReportKPITobaccoWeightModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public decimal Mean { get; set; }
        public decimal Stdev { get; set; }
        public decimal MeanMC { get; set; }
        public decimal StdevMC { get; set; }
    }
    public class ReportKPIRipperInfoModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string Material { get; set; }
        public string OrderNum { get; set; }
        public DateTime ActualStartDate { get; set; }
        public string Description { get; set; }
        public decimal QtyIss { get; set; }
        public decimal QtyRec { get; set; }
        public decimal Yield { get; set; }
        public decimal ValIssued { get; set; }
        public decimal ValReceiv { get; set; }
        public decimal ValDiffer { get; set; }
    }
    public class ReportKPIDustModel : BaseModel
    {
        public string ProductionCenter { get; set; }
        public int Dust { get; set; }
        public int Winnower { get; set; }
        public int FloorSweeping { get; set; }
        public int RS { get; set; }
        public decimal GRSpec { get; set; }
    }
    public class ReportKPICRRConversionModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public decimal DebuFinal { get; set; }
        public decimal DustHalusLM { get; set; }
        public decimal DustHalusLMnoRipper { get; set; }
        public decimal SaponTembakau { get; set; }
        public decimal SaponCigarette { get; set; }
        public decimal ClaimableInThStick { get; set; }
        public decimal AverageFinalOv { get; set; }
        public decimal C2Weight { get; set; }
    }

    public class ReportKPITargetModel : BaseModel
    {
        public string ProductionCenter { get; set; }
        public string KPI { get; set; }
        public decimal TargetInternal { get; set; }
        public decimal TargetOB { get; set; }
    }


    public class DataKPIProdVolModel
    {
        public List<ReportKPIProdVolModel> Data { get; set; }
    }
    public class DataKPICRRModel
    {
        public List<ReportKPICRRModel> Data { get; set; }
    }
    public class DataKPIKPIDIMModel
    {
        public List<ReportKPIDIMModel> Data { get; set; }
    }
    public class DataKPIYieldModel
    {
        public List<ReportKPIYieldModel> Data { get; set; }
    }
    public class DataKPIWorkHourModel
    {
        public List<ReportKPIWorkHourModel> Data { get; set; }
    }
    public class DataKPIStickPerPackModel
    {
        public List<ReportKPIStickPerPackModel> Data { get; set; }
    }
    public class DataKPIDustModel
    {
        public List<ReportKPIDustModel> Data { get; set; }
    }
    public class DataKPITobaccoWeightModel
    {
        public List<ReportKPITobaccoWeightModel> Data { get; set; }
    }
    public class DataKPIRipperInfoModel
    {
        public List<ReportKPIRipperInfoModel> Data { get; set; }
    }
    public class DataKPICRRConversionModel
    {
        public List<ReportKPICRRConversionModel> Data { get; set; }
    }
    public class DataKPITargetModel
    {
        public List<ReportKPITargetModel> Data { get; set; }
    }

    public class ReportKPIFormModel
    {
        public string Type { get; set; }
        public string PC { get; set; }
        public int Year { get; set; }
        public string Week { get; set; }
        public int Month { get; set; }
        public HttpPostedFileBase Excel { get; set; }
    }
}