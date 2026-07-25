using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    // ---------------------------------------------------------------------------
    // Filter Model (header area)
    // ---------------------------------------------------------------------------
    public class VendorBillingFilterModel
    {
        public string VendorAccNo { get; set; }
        public string VendorAccTitle { get; set; }
        public string DebitHeadAccNo { get; set; }
        public string DebitHeadAccTitle { get; set; }
        public int DateRangeIndex { get; set; } = 1; // default: Last 15 Days
        public DateTime DateFrom { get; set; } = DateTime.Today.AddDays(-15);
        public DateTime DateTo { get; set; } = DateTime.Today;
        public DateTime PostingDate { get; set; } = DateTime.Today;
    }

    // ---------------------------------------------------------------------------
    // Tab 1 – Vendor Billing grid row
    // ---------------------------------------------------------------------------
    public class VendorBillingRowViewModel
    {
        public int EntryID { get; set; }
        public string PORefNo { get; set; }
        public DateTime RcvDate { get; set; }
        public string MaterialID { get; set; }
        public string RMName { get; set; }
        public float QtyRcvd { get; set; }
        public float ExtraQtyRcvd { get; set; }
        public float QtyRejected { get; set; }
        public float Rate { get; set; }
        public byte IssuanceType { get; set; }

        // Computed
        public float NetQty => QtyRcvd + ExtraQtyRcvd - QtyRejected;
        public float Value => NetQty * Rate;

        public string OrderType => IssuanceType == 1 ? "Annealing"
                                 : IssuanceType == 2 ? "Machining"
                                 : "Regular";

        // UI state
        public bool IsChecked { get; set; } = false;
        public bool IsEditingRate { get; set; } = false;
        public float EditingRate { get; set; }
    }

    // ---------------------------------------------------------------------------
    // Tab 2 – Material Deductions (read-only issuance grid)
    // ---------------------------------------------------------------------------
    public class MaterialDeductionRowViewModel
    {
        public int EntryID { get; set; }
        public string IssNo { get; set; }
        public DateTime DT { get; set; }
        public string RMID1 { get; set; }
        public string RMName { get; set; }
        public float Price { get; set; }
        public float Qty { get; set; }
        public float Amount => Qty * Price;
    }

    // ---------------------------------------------------------------------------
    // Tab 3 – Manual Deductions
    // ---------------------------------------------------------------------------
    public class ManualDeductionRowViewModel
    {
        public string Description { get; set; }
        public decimal Amount { get; set; }
    }

    // ---------------------------------------------------------------------------
    // Summary / Deduction footer
    // ---------------------------------------------------------------------------
    public class VendorBillingSummaryModel
    {
        public float TotalQtyAll { get; set; }
        public float TotalValueAll { get; set; }
        public float TotalQtySelected { get; set; }
        public float TotalValueSelected { get; set; }
        public decimal ShortTermDeduction { get; set; }
        public decimal LongTermDeduction { get; set; }
        public float MaterialDeduction { get; set; }
        public decimal ManualDeduction { get; set; }
        public decimal AmountToPost => (decimal)TotalValueSelected - ShortTermDeduction - LongTermDeduction - (decimal)MaterialDeduction - ManualDeduction;
    }

    // ---------------------------------------------------------------------------
    // Posting Model (passed to DAL for the atomic transaction)
    // ---------------------------------------------------------------------------
    public class VendorBillingPostModel
    {
        public string VendorAccNo { get; set; }
        public string DebitHeadAccNo { get; set; }
        public DateTime PostingDate { get; set; }
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public int BillNo { get; set; }
        public string VchrNo { get; set; }
        public decimal ManualDedAmt { get; set; }
        public decimal STDeduction { get; set; }
        public decimal LTDeduction { get; set; }
        public decimal MaterialDedAmt { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
        public List<int> CheckedEntryIDs { get; set; } = new();
        public List<int> MaterialIssuanceEntryIDs { get; set; } = new();
        public List<ManualDeductionRowViewModel> ManualDeductions { get; set; } = new();
    }

    // ---------------------------------------------------------------------------
    // Lookups
    // ---------------------------------------------------------------------------
    public class VendorLookupModel
    {
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string DisplayName => AccTitle;
    }

    public class AccountLookupModel
    {
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string DisplayName => $"{AccTitle} {{{AccNo}}}";
    }
}
