using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services;
using Impulse.Services.Export;

namespace Impulse.Pages.Export.CustomInvoices
{
    public partial class PostPrc
    {
        [Parameter] public int EntryID { get; set; }

        [Inject] private ICustomPaymentService CustomPaymentService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = default!;
        [Inject] private IDBHelperService DBHelperService { get; set; } = default!;

        private PostPrcModel? Model { get; set; }
        private List<GenericDropDownModel> Banks { get; set; } = new();
        private bool IsLoading { get; set; } = true;
        private bool isSaving { get; set; } = false;

        private string NewDeductionTitle { get; set; } = string.Empty;
        private decimal NewDeductionAmount { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Banks = await CustomPaymentService.GetPrcBanksAsync();

                Model = await CustomPaymentService.GetPrcDetailsAsync(EntryID);
                if (Model != null)
                {
                    Model.PostingDate = DateTime.Today;
                    Model.RealizationDate = DateTime.Today;
                    Model.SerialDate = DateTime.Today;
                    Model.BillDate = DateTime.Today;

                    // ── Load account numbers from GeneralData (mirrors VB6 ShowMe logic) ──
                    // The VB6 uses getGeneralDataValue("Key") to fetch account numbers from
                    // the GeneralData table. The ListView Item Key = AccNo, used as debit line.
                    string fbChargesAccNo   = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_FBCharges_AccNo'");
                    string edsAccNo         = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_EDS_AccNo'");
                    string bscAccNo         = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_BSC_AccNo'");
                    string withHoldingAccNo = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_WithHolding_AccNo'");
                    string exchDiffAccNo    = await DBHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='PRC_ExchDiff_AccNo'");

                    // Store ExchDiff AccNo on the model (used in PostPrcAsync for DR/CR logic)
                    Model.ExchDiffAccNo = exchDiffAccNo;

                    // Build deduction rows with their account numbers
                    // (Order matches VB6 frmPRC ShowMe — account # is the key that becomes the ledger entry)
                    Model.Deductions.Add(new PrcDeductionModel { Title = "Bank Charges (F B Charges)", AccountNo = fbChargesAccNo,   Amount = 0, Currency = "USD" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "EDS",                         AccountNo = edsAccNo,         Amount = 0, Currency = "PKR" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "Bank Service Charges",        AccountNo = bscAccNo,         Amount = 0, Currency = "PKR" });
                    Model.Deductions.Add(new PrcDeductionModel { Title = "WithHolding Tax",             AccountNo = withHoldingAccNo, Amount = 0, Currency = "PKR" });
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

            Model?.Deductions.Add(new PrcDeductionModel { Title = NewDeductionTitle, Amount = NewDeductionAmount });
            
            NewDeductionTitle = string.Empty;
            NewDeductionAmount = 0;
        }

        private async Task PostVoucher()
        {
            if (Model == null || isSaving) return;

            if (string.IsNullOrWhiteSpace(Model.BankAccNo))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select a Bank Account.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.SerialNo))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please enter Running Serial No.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.BillNo))
            {
                NotificationServiceManager.ShowWarning("Validation", "Please enter Bill Number.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.CustomerAccNo))
            {
                NotificationServiceManager.ShowError("Validation", "Customer receivable account not found. Cannot post voucher.");
                return;
            }

            isSaving = true;
            try
            {
                string vchrNo = await CustomPaymentService.PostPrcAsync(Model);
                
                if (!string.IsNullOrEmpty(vchrNo))
                {
                    NotificationServiceManager.ShowSuccess("Posted", $"PRC posted successfully. Voucher No: {vchrNo}");
                    NavigationManager.NavigateTo("/export/custom-payment-status");
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
            NavigationManager.NavigateTo("/export/custom-payment-status");
        }
    }
}
