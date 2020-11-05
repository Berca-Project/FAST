using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIProdVol : BaseEntity
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
}
