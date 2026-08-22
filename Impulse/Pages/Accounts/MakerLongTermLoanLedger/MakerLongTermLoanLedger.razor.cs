using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using Impulse.Services.Accounts;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts.MakerLongTermLoanLedger
{
    public partial class MakerLongTermLoanLedger : ComponentBase
    {
        [Inject] private IMakerLongLoanLedgerService LedgerService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;

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
                    _ = RefreshLedgerAsync();
                }
            }
        }

        private DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        private DateTime ToDate { get; set; } = DateTime.Today;

        private MakerLongLoanLedgerResultModel LedgerData { get; set; } = new();
        private bool IsLoading { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Makers = await LedgerService.GetMakerAccountsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load accounts: {ex.Message}"
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task RefreshLedgerAsync()
        {
            if (SelectedMaker == null || string.IsNullOrWhiteSpace(SelectedMaker.AccNo))
            {
                LedgerData = new MakerLongLoanLedgerResultModel();
                return;
            }

            IsLoading = true;
            StateHasChanged();

            try
            {
                LedgerData = await LedgerService.GetLedgerDataAsync(
                    SelectedMaker.AccNo,
                    SelectedMaker.MakerName,
                    FromDate,
                    ToDate
                );
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Ledger",
                    Detail = ex.Message
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private async Task<IEnumerable<MakerAccountModel>> SearchMakersAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(Makers);
            return await Task.FromResult(Makers.Where(m =>
                (m.MakerName != null && m.MakerName.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.MakerNo != null && m.MakerNo.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.AccNo != null && m.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task PrintLedgerAsync()
        {
            if (SelectedMaker == null || string.IsNullOrWhiteSpace(SelectedMaker.AccNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Selection Warning",
                    Detail = "Please select a maker/vendor account first."
                });
                return;
            }

            try
            {
                string selectionFormula = $"{{VMakerAdvancesLedger.AccNo}}='{SelectedMaker.AccNo}' AND {{VMakerAdvancesLedger.DT}} in Date({FromDate.Year}, {FromDate.Month}, {FromDate.Day}) to Date({ToDate.Year}, {ToDate.Month}, {ToDate.Day})";
                var req = new ReportRequest
                {
                    ReportName = "MakerAdvancesLedger_Long.rpt",
                    SelectionFormula = selectionFormula
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Print Error",
                    Detail = ex.Message
                });
            }
        }
    }
}
