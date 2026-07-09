using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services.Export;
using Radzen;

namespace Impulse.Pages.Export.AdvancePayments
{
    public partial class AdvancePaymentPosting
    {
        [Parameter] public int EntryID { get; set; }

        [Inject] private IAdvancePaymentService AdvancePaymentService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        private AdvancePaymentViewModel? Payment { get; set; }
        private List<AdvancePaymentBankModel> Banks { get; set; } = new();
        private bool IsLoading { get; set; } = true;

        private DateTime PostingDate { get; set; } = DateTime.Today;

        private string NewDeductionTitle { get; set; } = string.Empty;
        private decimal NewDeductionAmount { get; set; }

        private List<DeductionModel> Deductions { get; set; } = new();

        public class DeductionModel
        {
            public string Title { get; set; } = string.Empty;
            public decimal Amount { get; set; }
            public string Currency { get; set; } = "PKR";
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Banks = await AdvancePaymentService.GetBanksAsync();
                
                Deductions.Add(new DeductionModel { Title = "Bank Charges (F B Charges)", Amount = 0, Currency = "USD" });
                Deductions.Add(new DeductionModel { Title = "With Holding Tax", Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "EDS", Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Postage", Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Bank Service Charges", Amount = 0, Currency = "PKR" });
                Deductions.Add(new DeductionModel { Title = "Bank Other Charges", Amount = 0, Currency = "PKR" });

                Payment = await AdvancePaymentService.GetAdvancePaymentAsync(EntryID);
                if (Payment != null)
                {
                    PostingDate = DateTime.Today;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error loading data", Detail = ex.Message });
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
            if (Payment == null) return;

            if (string.IsNullOrWhiteSpace(Payment.BankAccNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Bank Account." });
                return;
            }

            try
            {
                // In a full implementation, we would pass the Deductions and PostingDate to the service
                // and build a complete VoucherMasterViewModel with VouchersDetails.
                // For now, it calls the stub that generates a VoucherNo and links it.
                string vchrNo = await AdvancePaymentService.PostToFinancialAsync(EntryID, Payment);
                
                if (!string.IsNullOrEmpty(vchrNo))
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Posted", Detail = $"Successfully posted to financials. Voucher No: {vchrNo}" });
                    NavigationManager.NavigateTo("/export/advance-payment-list");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Posting Failed", Detail = ex.Message });
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/export/advance-payment-list");
        }
    }
}
