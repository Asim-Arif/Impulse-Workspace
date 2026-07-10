using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services.Export;
using Radzen;

namespace Impulse.Pages.Export.CustomInvoices
{
    public partial class PostPrc
    {
        [Parameter] public int EntryID { get; set; }

        [Inject] private ICustomPaymentService CustomPaymentService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

        private PostPrcModel? Model { get; set; }
        private List<GenericDropDownModel> Banks { get; set; } = new();
        private bool IsLoading { get; set; } = true;

        private string NewDeductionTitle { get; set; } = string.Empty;
        private decimal NewDeductionAmount { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Banks = await CustomPaymentService.GetBanksAsync();

                Model = await CustomPaymentService.GetPrcDetailsAsync(EntryID);
                if (Model != null)
                {
                    Model.PostingDate = DateTime.Today;
                    Model.RealizationDate = DateTime.Today;
                    Model.SerialDate = DateTime.Today;
                    Model.BillDate = DateTime.Today;

                    Model.Deductions.Add(new PrcDeductionModel { Title = "Bank Charges (F B Charges)", Amount = 0, Currency = "USD" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "EDS", Amount = 0, Currency = "PKR" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "Bank Service Charges", Amount = 0, Currency = "PKR" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "WithHolding Tax", Amount = 0, Currency = "PKR" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "Exch. Diff. Account", Amount = 0, Currency = "PKR" });
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

            Model?.Deductions.Add(new PrcDeductionModel { Title = NewDeductionTitle, Amount = NewDeductionAmount });
            
            NewDeductionTitle = string.Empty;
            NewDeductionAmount = 0;
        }

        private async Task PostVoucher()
        {
            if (Model == null) return;

            if (string.IsNullOrWhiteSpace(Model.BankAccNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please select a Bank Account." });
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.SerialNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please enter Running Serial No." });
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.BillNo))
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Warning, Summary = "Validation", Detail = "Please enter Bill Number." });
                return;
            }

            try
            {
                string vchrNo = await CustomPaymentService.PostPrcAsync(Model);
                
                if (!string.IsNullOrEmpty(vchrNo))
                {
                    NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Posted", Detail = $"Successfully posted PRC. Voucher No: {vchrNo}" });
                    NavigationManager.NavigateTo("/export/custom-payment-status");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Posting Failed", Detail = ex.Message });
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo("/export/custom-payment-status");
        }
    }
}
