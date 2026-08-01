using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services.Accounts;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts.MakerLoanClearance
{
    public partial class MakerLoanClearance : ComponentBase
    {
        [Inject] private IMakerLoanClearanceService ClearanceService { get; set; } = default!;
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

        private DateTime ClearanceDate { get; set; } = DateTime.Today;
        private MakerLoanBalanceModel Balances { get; set; } = new();
        private decimal LTToClear { get; set; } = 0m;
        private decimal STToClear { get; set; } = 0m;
        private decimal NewLTDeductionAmount { get; set; } = 0m;
        private string Remarks { get; set; } = string.Empty;

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;
        private bool IsUpdatingDed { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Makers = await ClearanceService.GetMakerAccountsAsync();
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
                Balances = await ClearanceService.GetMakerBalancesAsync(maker.AccNo);
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

        private async Task UpdateDeductionAmountAsync()
        {
            if (SelectedMaker == null || string.IsNullOrWhiteSpace(SelectedMaker.AccNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Selection Required",
                    Detail = "Please select a maker account first."
                });
                return;
            }

            if (NewLTDeductionAmount < 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Deduction amount cannot be negative."
                });
                return;
            }

            IsUpdatingDed = true;
            StateHasChanged();

            try
            {
                await ClearanceService.UpdateLTDeductionAmountAsync(SelectedMaker.AccNo, NewLTDeductionAmount);
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Updated",
                    Detail = "Long-term deduction amount updated successfully."
                });
                await LoadBalancesAsync(SelectedMaker);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Error",
                    Detail = ex.Message
                });
            }
            finally
            {
                IsUpdatingDed = false;
                StateHasChanged();
            }
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
                    Detail = "Please enter valid clearance amount for Long Term or Short Term loan."
                });
                return false;
            }

            if (LTToClear > Balances.LTBalance)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = $"L.T amount to clear (PKR {LTToClear:N2}) exceeds current L.T balance (PKR {Balances.LTBalance:N2})."
                });
                return false;
            }

            if (STToClear > Balances.STBalance)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = $"S.T amount to clear (PKR {STToClear:N2}) exceeds current S.T balance (PKR {Balances.STBalance:N2})."
                });
                return false;
            }

            IsSaving = true;
            StateHasChanged();

            try
            {
                var dto = new MakerLoanClearanceDto
                {
                    AccNo = SelectedMaker.AccNo,
                    ClearanceDate = ClearanceDate,
                    LTToClear = LTToClear,
                    STToClear = STToClear,
                    NewLTDeductionAmount = NewLTDeductionAmount,
                    Remarks = Remarks
                };

                await ClearanceService.SaveClearanceAsync(dto);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved Successfully",
                    Detail = "Maker loan clearance saved successfully."
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
