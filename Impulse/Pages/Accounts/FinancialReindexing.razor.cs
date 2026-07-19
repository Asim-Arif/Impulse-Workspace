using DataAccessLibrary.Interface.Accounts;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Accounts
{
    public partial class FinancialReindexing
    {
        [Inject]
        public IChartOfAccountsDataAccess ChartOfAccountsDataAccess { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public bool IsProcessing { get; set; } = false;
        public string SuccessMessage { get; set; }
        public string ErrorMessage { get; set; }

        private async Task StartReindexing()
        {
            IsProcessing = true;
            SuccessMessage = string.Empty;
            ErrorMessage = string.Empty;
            
            try
            {
                await ChartOfAccountsDataAccess.ReindexFinancialBalancesAsync();
                SuccessMessage = "Financial balances have been successfully re-indexed.";
            }
            catch (Exception ex)
            {
                ErrorMessage = $"An error occurred during re-indexing: {ex.Message}";
            }
            finally
            {
                IsProcessing = false;
            }
        }

        private void GoToIndexPage()
        {
            NavigationManager.NavigateTo("/financial");
        }
    }
}
