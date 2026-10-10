namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class StockBreachItemDto
    {
        public int RMID { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string RMName { get; set; } = string.Empty;
        public string? Unit { get; set; }
        public string? GroupName { get; set; }
        public double QtyInStock { get; set; }
        public double? MinLevel { get; set; }
        public double? MaxLevel { get; set; }
        public int? ReorderPoint { get; set; }
        public string BreachType { get; set; } = "Normal"; // "MinStock", "Reorder", "MaxStock", "Normal"
        public double DeficitOrSurplus
        {
            get
            {
                if (BreachType == "MinStock" && MinLevel.HasValue)
                    return MinLevel.Value - QtyInStock;
                if (BreachType == "Reorder" && ReorderPoint.HasValue)
                    return ReorderPoint.Value - QtyInStock;
                if (BreachType == "MaxStock" && MaxLevel.HasValue)
                    return QtyInStock - MaxLevel.Value;
                return 0;
            }
        }
    }

    public class StockBreachSummaryDto
    {
        public int BelowReorderCount { get; set; }
        public int BelowMinCount { get; set; }
        public int AboveMaxCount { get; set; }
        public int TotalBreaches => BelowReorderCount + BelowMinCount + AboveMaxCount;
    }
}
