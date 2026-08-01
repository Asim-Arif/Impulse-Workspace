using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services.Accounts;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts.MakerLoanTransfer
{
    public partial class MakerLoanTransfer : ComponentBase
    {
        [Inject] private IMakerLoanTransferService TransferService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<MakerAccountModel> Makers { get; set; } = new();

        private MakerAccountModel? _selectedMaker;
        private MakerAccountModel? SelectedMaker
        {
            get => _selectedMaker;
            set
            {
                if (_selectedMaker != value)
                {
                    _selectedMaker = value;
                    _ = LoadBalancesAsync(value);
                }
            }
        }

        private DateTime TransferDate { get; set; } = DateTime.Today;
        private MakerLoanBalanceModel Balances { get; set; } = new();
        private decimal LTToClear { get; set; } = 0m;
        private decimal STToClear { get; set; } = 0m;
        private decimal NewLTDeductionAmount { get; set; } = 0m;
        private string Remarks { get; set; } = string.Empty;

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Makers = await TransferService.GetMakerAccountsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load maker accounts: {ex.Message}"
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task LoadBalancesAsync(MakerAccountModel? maker)
        {
            if (maker != null && !string.IsNullOrWhiteSpace(maker.AccNo))
            {
                Balances = await TransferService.GetMakerBalancesAsync(maker.AccNo);
                NewLTDeductionAmount = Balances.LTDeductionAmount;
            }
            else
            {
                Balances = new MakerLoanBalanceModel();
                NewLTDeductionAmount = 0m;
            }
            StateHasChanged();
        }

        private async Task<IEnumerable<MakerAccountModel>> SearchMakersAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(Makers);
            return await Task.FromResult(Makers.Where(m =>
                (m.MakerName != null && m.MakerName.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.MakerNo != null && m.MakerNo.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.AccNo != null && m.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<bool> SaveAsync()
        {
            if (SelectedMaker == null || string.IsNullOrWhiteSpace(SelectedMaker.AccNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select a valid Maker account."
                });
                return false;
            }

            if (LTToClear <= 0 && STToClear <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter valid transfer amount for L.T to S.T or S.T to L.T."
                });
                return false;
            }

            if (LTToClear > Balances.LTBalance)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = $"L.T amount to transfer (PKR {LTToClear:N2}) exceeds current L.T balance (PKR {Balances.LTBalance:N2})."
                });
                return false;
            }

            if (STToClear > Balances.STBalance)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = $"S.T amount to transfer (PKR {STToClear:N2}) exceeds current S.T balance (PKR {Balances.STBalance:N2})."
                });
                return false;
            }

            IsSaving = true;
            StateHasChanged();

            try
            {
                var dto = new MakerLoanTransferDto
                {
                    AccNo = SelectedMaker.AccNo,
                    TransferDate = TransferDate,
                    LTToClear = LTToClear,
                    STToClear = STToClear,
                    NewLTDeductionAmount = NewLTDeductionAmount,
                    Remarks = Remarks
                };

                await TransferService.SaveTransferAsync(dto);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved Successfully",
                    Detail = "Maker loan transfer completed successfully."
                });

                return true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message
                });
                return false;
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private async Task SaveAndNewAsync()
        {
            if (await SaveAsync())
            {
                LTToClear = 0m;
                STToClear = 0m;
                Remarks = string.Empty;
                if (SelectedMaker != null)
                {
                    await LoadBalancesAsync(SelectedMaker);
                }
            }
        }

        private async Task SaveAndCloseAsync()
        {
            if (await SaveAsync())
            {
                NavigateBack();
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/financial");
            }
        }
    }
}
