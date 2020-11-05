using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIYield : BaseEntity
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
}
