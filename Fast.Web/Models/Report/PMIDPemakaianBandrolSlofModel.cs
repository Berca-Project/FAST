using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PMIDPemakaianBandrolSlofModel
    {
        public long ID { get; set; }
        public long LPHID { get; set; }
        public Int32 Week { get; set; }
        public Nullable<DateTime> Date { get; set; }
        public string Shift { get; set; }
        public string Group { get; set; }
        public string Machine { get; set; }
        public string FACode { get; set; }
        public Nullable<double> TSRequest { get; set; }
        public Nullable<double> ManualTaxStamp { get; set; }
        public Nullable<double> StockAwal { get; set; }
        public Nullable<double> PackOnMachine { get; set; }
        public Nullable<double> SamplingCabinet { get; set; }
        public Nullable<double> PackOutofMachine { get; set; }
        public Nullable<double> OtherMachine { get; set; }
        public Nullable<double> TotalOpening { get; set; }
        public Nullable<double> CaseWTS { get; set; }
        public Nullable<double> TaxStampReturn { get; set; }
        public Nullable<double> SisaAkhirBandrol { get; set; }
        public Nullable<double> PackonSamplingCabinet { get; set; }
        public Nullable<double> PackOnMachineClosing { get; set; }
        public Nullable<double> PackOutofMachineClosing { get; set; }
        public Nullable<double> ClaimablePack { get; set; }
        public Nullable<double> IPCCounted { get; set; }
        public Nullable<double> IPCOther { get; set; }
        public Nullable<double> QASamplewStamp { get; set; }
        public Nullable<double> WIPtoOtherMachine { get; set; }
        public Nullable<double> TotalClosing { get; set; }
        public Nullable<double> Missing { get; set; }
        public Nullable<double> PercentMissing { get; set; }
        public Nullable<double> Usage { get; set; }
        public string PlusMinusBandrol { get; set; }
        public string TaxStampCode { get; set; }
        public string SubmissionStatus { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }
    }
}
