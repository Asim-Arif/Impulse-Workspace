using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.StockLedger
{
    public partial class StockLedger : ComponentBase
    {
        private bool IsLoadingInitialData = true;
        private bool IsLoading = false;

        private List<RMViewModel> Materials = new();
        public RMViewModel SelectedMaterial { get; set; }

        public DateTime FromDate { get; set; } = DateTime.Today.AddDays(-60);
        public DateTime ToDate { get; set; } = DateTime.Today;
        public bool IsCustomDate { get; set; } = false;

        public double OpeningBalance { get; set; } = 0;
        public List<StockLedgerTransactionModel> Transactions { get; set; }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Materials = await LedgerService.GetMaterialsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsLoadingInitialData = false;
            }
        }

        private async Task<IEnumerable<RMViewModel>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Materials;
                
            return await Task.FromResult(Materials.Where(x => x.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private void OnDateRangeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int days))
            {
                if (days == -1)
                {
                    IsCustomDate = true;
                }
                else
                {
                    IsCustomDate = false;
                    ToDate = DateTime.Today;
                    FromDate = DateTime.Today.AddDays(-days);
                }
            }
        }

        private async Task LoadLedgerAsync()
        {
            if (SelectedMaterial == null)
            {
                NotificationService.ShowWarning("Validation", "Please select a material first.");
                return;
            }

            IsLoading = true;
            try
            {
                OpeningBalance = await LedgerService.GetOpeningBalanceAsync(SelectedMaterial.RMID1, FromDate);
                Transactions = await LedgerService.GetTransactionsAsync(SelectedMaterial.RMID1, FromDate, ToDate);

                // Calculate running balance
                double currentBalance = OpeningBalance;
                foreach (var tran in Transactions)
                {
                    currentBalance += tran.Quantity;
                    tran.Balance = currentBalance;
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading ledger", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task PrintLedgerAsync()
        {
            if (SelectedMaterial == null)
            {
                NotificationService.ShowWarning("Validation", "Please select a material first.");
                return;
            }

            // Crystal formulas match what VB6 passed
            string selectionFormula = $"{{VSTockLedger.MaterialID}}='{SelectedMaterial.RMID1}' and ({{VStockLedger.TransactionDate}} in Date({FromDate.Year},{FromDate.Month},{FromDate.Day}) to Date({ToDate.Year},{ToDate.Month},{ToDate.Day}))";
            
            var formulaValues = new Dictionary<string, object>
            {
                { "FromTo", $"'From   {FromDate:dd-MMM-yyyy}    to     {ToDate:dd-MMM-yyyy}'" },
                { "OpeningBalance", OpeningBalance.ToString() }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.Stock.StockLedger,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await ReportNavigation.PrintReportAsync(request);
        }
    }
}
