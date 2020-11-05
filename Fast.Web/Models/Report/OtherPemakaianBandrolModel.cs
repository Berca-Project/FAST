using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class OtherPemakaianBandrolModel
    {
        public long ID { get; set; }
        public long LPHID { get; set; }
        public Nullable<DateTime> Date { get; set; }
        public string Shift { get; set; }
        public string Machine { get; set; }
        public Nullable<double> SisaAwalWeek { get; set; }
        public Nullable<double> HasilBox { get; set; }
        public Nullable<double> HasilCounter { get; set; }
        public Nullable<double> HasilProduksi { get; set; }
        public Nullable<double> StockAwal { get; set; }
        public Nullable<double> ReworkFromMachine { get; set; }
        public Nullable<double> ReceiveFromRework { get; set; }
        public Nullable<double> StockAkhir { get; set; }
        public Nullable<double> TerimaBandrol { get; set; }
        public Nullable<double> Pakai { get; set; }
        public Nullable<double> BSAfkir { get; set; }
        public Nullable<double> SampleQA { get; set; }
        public Nullable<double> IPC { get; set; }
        public Nullable<double> SisaAkhirBandrol { get; set; }
        public Nullable<double> BandrolHilang { get; set; }
        public Nullable<double> PercentBandrolHilang { get; set; }
        public Nullable<double> ReworkTBPlus { get; set; }
        public Nullable<double> BandrolLepas { get; set; }
        public Nullable<double> Claimable { get; set; }
        public Nullable<double> LPHvsRework { get; set; }
        public string SubmissionStatus { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }
        public string Brand { get; set; }
    }
}
