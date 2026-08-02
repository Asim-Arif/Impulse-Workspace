using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class LotSearchResultModel
    {
        public long EntryID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string RecieptID { get; set; } = string.Empty;
        public long VendID { get; set; }
        public string VendorName { get; set; } = string.Empty;
        public string VendID1 { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public bool Authorized { get; set; }
        public bool Closed { get; set; }
        public string IssEmpID { get; set; } = string.Empty;
        public bool IsFactoryMaker { get; set; }
        public bool AlreadyReceived { get; set; }
    }
}
