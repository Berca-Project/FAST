using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIDIM : BaseEntity
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
}
