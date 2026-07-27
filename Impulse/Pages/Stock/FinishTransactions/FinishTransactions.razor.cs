using BlazorContextMenu;
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

namespace Impulse.Pages.Stock.FinishTransactions
{
    public partial class FinishTransactions : ComponentBase
    {
        private bool IsLoading = false;
        private string CurrentUserName = "System";

        private DateTime FromDate { get; set; } = DateTime.Today;
        private DateTime ToDate { get; set; } = DateTime.Today;
        private bool IsCustomRange { get; set; } = false;

        private bool IncludeIssuance { get; set; } = true;
        private bool IncludeReceiving { get; set; } = true;

        private List<StoreLookupModel> Stores = new();
        private int SelectedStoreId { get; set; } = 0;

        private List<ItemViewModel> Articles = new();
        private ItemViewModel SelectedArticle { get; set; }

        private List<FinishTransactionModel> Transactions = new();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                Stores = await TransactionService.GetStoresAsync(CurrentUserName);
                Articles = await TransactionService.GetArticlesAsync();

                await LoadTransactionsAsync();
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

        private void OnRangeChanged(ChangeEventArgs e)
        {
            string val = e.Value?.ToString() ?? "Today";
            DateTime today = DateTime.Today;

            IsCustomRange = false;

            switch (val)
            {
                case "Today":
                    FromDate = today;
                    ToDate = today;
                    break;
                case "15":
                    FromDate = today.AddDays(-15);
                    ToDate = today;
                    break;
                case "30":
                    FromDate = today.AddDays(-30);
                    ToDate = today;
                    break;
                case "60":
                    FromDate = today.AddDays(-60);
                    ToDate = today;
                    break;
                case "90":
                    FromDate = today.AddDays(-90);
                    ToDate = today;
                    break;
                case "Custom":
                    IsCustomRange = true;
                    break;
            }
        }

        private async Task LoadTransactionsAsync()
        {
            IsLoading = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var filter = new FinishTransactionFilterModel
                {
                    FromDate = FromDate,
                    ToDate = ToDate,
                    IncludeIssuance = IncludeIssuance,
                    IncludeReceiving = IncludeReceiving,
                    StoreRefId = SelectedStoreId,
                    ItemID = SelectedArticle?.ItemID ?? "0"
                };

                Transactions = await TransactionService.GetTransactionsAsync(filter);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading transactions", ex.Message);
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task PrintTransactionsReport()
        {
            try
            {
                string formula = $"{{VStockTransactions.DT}} = Date({FromDate:yyyy,MM,dd}) to Date({ToDate:yyyy,MM,dd})";

                if (!IncludeIssuance)
                    formula += " and {VStockTransactions.EntryType} <> 0";
                if (!IncludeReceiving)
                    formula += " and {VStockTransactions.EntryType} <> 1";
                if (SelectedStoreId > 0)
                    formula += $" and {{VStockTransactions.Store_RefID}} = {SelectedStoreId}";
                if (SelectedArticle != null && SelectedArticle.ItemID != "0")
                    formula += $" and {{VStockTransactions.ItemID}} = '{SelectedArticle.ItemID}'";

                string heading = "Finished Stock Transactions";
                if (IncludeIssuance && !IncludeReceiving) heading = "Issuance Entries";
                else if (!IncludeIssuance && IncludeReceiving) heading = "Receiving Entries";

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.StockTransactions,
                    SelectionFormula = formula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Heading", heading },
                        { "DateRange", $"{FromDate:dd-MMM-yyyy} to {ToDate:dd-MMM-yyyy}" }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintFinishedStockReport()
        {
            try
            {
                string formula = "";
                if (SelectedArticle != null && SelectedArticle.ItemID != "0")
                {
                    formula = $"{{VItems_StockReport.ItemID}} = '{SelectedArticle.ItemID}'";
                }

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.FinishStockDetail,
                    SelectionFormula = formula
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintLocationWiseReport()
        {
            try
            {
                string formula = "{VItemsWithShelfWiseStock.NetQty} > 0";
                if (SelectedArticle != null && SelectedArticle.ItemID != "0")
                {
                    formula += $" and {{VItems.ItemID}} = '{SelectedArticle.ItemID}'";
                }

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.BinWiseFinishStockStatus,
                    SelectionFormula = formula
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task PrintRackWiseReport()
        {
            try
            {
                string formula = "{VItemsWithShelfWiseStock.NetQty} > 0";
                if (SelectedArticle != null && SelectedArticle.ItemID != "0")
                {
                    formula += $" and {{VItems.ItemID}} = '{SelectedArticle.ItemID}'";
                }

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Stock.RackWiseFinishStockStatus,
                    SelectionFormula = formula
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private async Task OnContextPrintSlip(ItemClickEventArgs e)
        {
            var item = e.Data as FinishTransactionModel;
            if (item == null)
            {
                NotificationService.ShowError("Error", "Unable to identify selected row data.");
                return;
            }

            try
            {
                string reportName = item.EntryNo.StartsWith("RCV", StringComparison.OrdinalIgnoreCase)
                    ? ReportNames.Stock.FinishStockReceivingSlip
                    : ReportNames.Stock.FinishStockIssuanceSlip;

                var req = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VStockTransactions.EntryNo}} = '{item.EntryNo}'"
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
