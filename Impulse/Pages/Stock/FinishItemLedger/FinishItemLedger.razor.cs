using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.FinishItemLedger
{
    public partial class FinishItemLedger : ComponentBase
    {
        private bool IsLoading = false;
        private string CurrentUserName = "System";

        private DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        private DateTime ToDate { get; set; } = DateTime.Today;

        private List<ItemViewModel> Articles = new();
        private ItemViewModel _selectedArticle;
        public ItemViewModel SelectedArticle
        {
            get => _selectedArticle;
            set
            {
                if (_selectedArticle != value)
                {
                    _selectedArticle = value;
                    _ = OnArticleChanged();
                }
            }
        }

        private ItemHeaderInfoModel HeaderInfo;
        private List<FinishItemLedgerRowModel> LedgerRows = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Articles = await LedgerService.GetArticlesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private async Task<IEnumerable<ItemViewModel>> SearchArticles(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Articles;

            return await Task.FromResult(Articles.Where(x => x.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnArticleChanged()
        {
            if (SelectedArticle == null)
            {
                HeaderInfo = null;
                LedgerRows.Clear();
                await InvokeAsync(StateHasChanged);
                return;
            }

            try
            {
                HeaderInfo = await LedgerService.GetItemHeaderInfoAsync(SelectedArticle.ItemID);
                await LoadLedgerAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadLedgerAsync()
        {
            if (SelectedArticle == null) return;

            IsLoading = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var filter = new FinishItemLedgerFilterModel
                {
                    ItemID = SelectedArticle.ItemID,
                    FromDate = FromDate,
                    ToDate = ToDate
                };

                var rawRows = await LedgerService.GetLedgerAsync(filter);
                var processedRows = new List<FinishItemLedgerRowModel>();

                if (rawRows.Any())
                {
                    double openingBal = rawRows.First().OpeningBalance;
                    double currentBal = openingBal;

                    processedRows.Add(new FinishItemLedgerRowModel
                    {
                        DT = null,
                        RcvNo = string.Empty,
                        OpeningBalance = openingBal,
                        RunningBalance = openingBal
                    });

                    foreach (var r in rawRows)
                    {
                        if (string.IsNullOrEmpty(r.RcvNo) && r.DT == null)
                            continue;

                        currentBal = currentBal + r.QtyIn - r.QtyOut;
                        r.RunningBalance = currentBal;
                        processedRows.Add(r);
                    }
                }

                LedgerRows = processedRows;
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading ledger", ex.Message);
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task PrintLedgerReport()
        {
            if (SelectedArticle == null) return;

            try
            {
                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.FinishItemLedger,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "@ItemID", SelectedArticle.ItemID },
                        { "@DTFrom", FromDate },
                        { "@DTTo", ToDate }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintBinCardReport()
        {
            if (SelectedArticle == null) return;

            try
            {
                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.FinishItemLedgerBinCard,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "@ItemID", SelectedArticle.ItemID },
                        { "@DTFrom", FromDate },
                        { "@DTTo", ToDate }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }
    }
}
