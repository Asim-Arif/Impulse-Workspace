using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendGateRcvdInspectionMaterialItem
    {
        public int EntryID { get; set; } // VendRcvdDetailPO.EntryID
        public string MaterialName { get; set; } // Formatted as "{RMID1} RMName"
        public string RMSize { get; set; }
        public int GroupID { get; set; }
        public decimal QtyRcvd { get; set; }
        public string OrderNo { get; set; }
        public string VendorName { get; set; }
        public DateTime RcvDate { get; set; }
    }

    public class VendGateRcvdInspectionDetailViewModel
    {
        public int RMGP_RefID { get; set; } // Reference to RMGroupParameters.EntryID
        public string ParameterName { get; set; } // Description
        public string AQL { get; set; }
        public string SampleSize { get; set; }
        public string RejectOn { get; set; }
        public string TestSpecificationNo { get; set; }
        public string ActualRejection { get; set; } = "0";
        public string Status { get; set; } = "OK";
    }

    public class VendGateRcvdInspectionViewModel
    {
        public int EntryID { get; set; } // VendRcvdDetailPO_Inspection.EntryID
        public int VRDPO_RefID { get; set; } // VendRcvdDetailPO.EntryID
        public string Disposation { get; set; }
        public string Comments { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public bool IsPassed { get; set; } = true; // Maps to Status in DB (1 = Pass, 0 = Fail)

        public List<VendGateRcvdInspectionDetailViewModel> Details { get; set; } = new List<VendGateRcvdInspectionDetailViewModel>();
    }
}
