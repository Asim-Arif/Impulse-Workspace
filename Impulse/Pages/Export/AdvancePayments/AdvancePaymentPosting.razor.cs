using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;

namespace Impulse.Pages.Export.AdvancePayments
{
    public partial class AdvancePaymentPosting
    {
        [Parameter] public int EntryID { get; set; }

        [Inject] private IAdvancePaymentService AdvancePaymentService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = default!;
        [Inject] private IDBHelperService DBHelperService { get; set; } = default!;

        private AdvancePaymentViewModel? Payment { get; set; }
        private List<AdvancePaymentBankModel> Banks { get; set; } = new();
        private bool IsLoading { get; set; } = true;
        private bool isSaving { get; set; } = false;

        private DateTime PostingDate { get; set; } = DateTime.Today;

        private string NewDeductionTitle { get; set; } = string.Empty;
        private decimal NewDeductionAmount { get; set; }

        private List<DeductionModel> Deductions { get; set; } = new();

        /// <summary>
        /// Represents a deduction row with its chart-of-accounts number for ledger posting.
        /// AccountNo is loaded from GeneralData on page init — mirrors the VB6 ListView key.
        /// </summary>
        public class DeductionModel
        {
            public string Title { get; set; } = string.Empty;
            /// <summary>Chart-of-accounts number from GeneralData (e.g. PRC_FBCharges_AccNo)</summary>
            public string AccountNo { get; set; } = string.Empty;
            public decimal Amount { get; set; }
            public string Currency { get; set; } = "PKR";
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Banks = await AdvancePaymentService.GetBanksAsync();

                // ── Load account numbers from GeneralData (mirrors VB6 ShowMe logic) ──
                string fbChargesAccNo   = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_FBCharges_AccNo'");
                string withHoldingAccNo = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_WithHolding_AccNo'");
                string edsAccNo         = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_EDS_AccNo'");
                string postageAccNo     = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_Postage_AccNo'");
                string bscAccNo         = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_BSC_AccNo'");
                string bocAccNo         = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_BOC_AccNo'");

                // Build deduction rows in same order as VB6 frmAdvancePaymentPosting ShowMe()
                Deductions.Add(new DeductionModel { Title = "Bank Charges (F B Charges)", AccountNo = fbChargesAccNo,   Amount = 0, Currency = "USD" });
                Deductions.Add(new DeductionModel { Title = "With Holding Tax",            AccountNo = withHoldingAccNo, Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "EDS",                         AccountNo = edsAccNo,         Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Postage",                     AccountNo = postageAccNo,     Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Bank Service Charges",        AccountNo = bscAccNo,         Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Bank Other Charges",          AccountNo = bocAccNo,         Amount = 0, Currency = "PKR" });

                Payment = await AdvancePaymentService.GetAdvancePaymentAsync(EntryID);
                if (Payment != null)
                {
                    PostingDate = DateTime.Today;
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error loading data", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void AddDeduction()
        {
            if (string.IsNullOrWhiteSpace(NewDeductionTitle)) return;
            if (NewDeductionAmount <= 0) return;

            Deductions.Add(new DeductionModel { Title = NewDeductionTitle, Amount = NewDeductionAmount });
            
            NewDeductionTitle = string.Empty;
            NewDeductionAmount = 0;
        }

        private async Task PostVoucher()
        {
            if (Payment == null || isSaving) return;

            if (string.IsNullOrWhiteSpace(Payment.BankAccNo))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select a Bank Account.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Payment.CustomerAccNo))
            {
                NotificationServiceManager.ShowError("Validation", "Customer receivable account not found. Cannot post voucher.");
                return;
            }

            isSaving = true;
            try
            {
                // Convert page's DeductionModel list to PrcDeductionModel (DataAccessLibrary type)
                var deductionModels = Deductions.Select(d => new PrcDeductionModel
                {
                    Title     = d.Title,
                    AccountNo = d.AccountNo,
                    Amount    = d.Amount,
                    Currency  = d.Currency
                }).ToList();

                string vchrNo = await AdvancePaymentService.PostToFinancialAsync(
                    EntryID,
                    Payment,
                    deductionModels,
                    PostingDate);
                
                if (!string.IsNullOrEmpty(vchrNo))
                {
                    NotificationServiceManager.ShowSuccess("Posted", $"Advance Payment posted to financials. Voucher No: {vchrNo}");
                    NavigationManager.NavigateTo("/export/advance-payment-list");
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Posting Failed", ex.Message);
            }
            finally
            {
                isSaving = false;
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/export/advance-payment-list");
        }
    }
}
