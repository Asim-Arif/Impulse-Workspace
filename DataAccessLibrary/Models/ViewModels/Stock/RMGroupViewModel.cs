using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMGroupViewModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
        public long GrpColor { get; set; }
    }
}
