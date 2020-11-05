using System;

namespace Fast.Domain.Entities
{
    public class PPReportYieldKreteks : BaseEntity
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public string Blend { get; set; }
        public Nullable<double> Leaf { get; set; }
        public Nullable<double> LeafOV { get; set; }
        public Nullable<double> Clove { get; set; }
        public Nullable<double> CloveOV { get; set; }
        public Nullable<double> CRES { get; set; }
        public Nullable<double> CRESOV { get; set; }
        public Nullable<double> DIET { get; set; }
        public Nullable<double> DIETOV { get; set; }
        public Nullable<double> RTC { get; set; }
        public Nullable<double> RTCOV { get; set; }
        public Nullable<double> SmallLamina { get; set; }
        public Nullable<double> SmallLaminaOV { get; set; }
        public Nullable<double> CloveSteamFlake { get; set; }
        public Nullable<double> CloveSteamFlakeOV { get; set; }
        public Nullable<double> InfeedMaterialWet { get; set; }
        public Nullable<double> InfeedMaterialDry { get; set; }
        public Nullable<double> BCWet { get; set; }
        public Nullable<double> BCDryMatter { get; set; }
        public Nullable<double> ACWet { get; set; }
        public Nullable<double> ACDryMatter { get; set; }
        public Nullable<double> CutFiller { get; set; }
        public Nullable<double> RipperShort { get; set; }
        public Nullable<double> Addback { get; set; }
        public Nullable<double> CFDry { get; set; }
        public Nullable<double> RSDry { get; set; }
        public Nullable<double> AddbackDry { get; set; }
        public Nullable<double> CFDryExclude { get; set; }
        public Nullable<double> AvgFinalOV { get; set; }
        public Nullable<double> RipperMC { get; set; }
        public Nullable<double> AddbackMC { get; set; }
        public Nullable<double> DryYield { get; set; }
        public Nullable<double> WetYield { get; set; }
    }
}
