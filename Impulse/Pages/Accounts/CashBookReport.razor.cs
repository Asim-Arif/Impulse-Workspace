using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Configurations;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Impulse.Pages.Accounts
{
    public partial class CashBookReport : ComponentBase
    {
        [Inject]
        private IAccountReportingAccessService AccountReportingAccessService { get; set; }
        
        [Inject]
        protected IReportNavigationService ReportNavigationService { get; set; }

        private bool IsLoading = false;
        private DateTime DTFrom { get; set; } = DateTime.Now.Date;
        private DateTime DTTo { get; set; } = DateTime.Now.Date;
        
        private List<Cash_Book_Report_ViewModel> ListToUse { get; set; } = new List<Cash_Book_Report_ViewModel>();
        
        private decimal OpeningBalance { get; set; } = 0;
        private decimal ClosingBalance { get; set; } = 0;
        private decimal TotalDebit { get; set; } = 0;
        private decimal TotalCredit { get; set; } = 0;

        protected override async Task OnInitializedAsync()
        {
            await RefreshCashLedger();
        }

        private async Task RefreshCashLedger()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                var listData = await AccountReportingAccessService.GetCashBookReport(DTFrom, DTTo);
                
                if (listData != null && listData.Any())
                {
                    ListToUse = listData.ToList();
                    OpeningBalance = ListToUse.First().OpeningBalance;
                    ClosingBalance = ListToUse.First().ClosingBalance;
                    TotalDebit = ListToUse.Sum(x => x.Debit);
                    TotalCredit = ListToUse.Sum(x => x.Credit);
                }
                else
                {
                    ListToUse = new List<Cash_Book_Report_ViewModel>();
                    OpeningBalance = 0;
                    ClosingBalance = 0;
                    TotalDebit = 0;
                    TotalCredit = 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in RefreshCashLedger: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/financialdashboard");
        }

        private async Task PrintCashBookReport()
        {
            string reportName = "TransRptCashnBank.rpt"; 
            
            var formulaValues = new Dictionary<string, object>
            {
                { "Company", "'CurrentCompany'" },
                { "FromTo", $"'{DTFrom:dd-MMM-yyyy} To {DTTo:dd-MMM-yyyy}'" }
            };

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = "", 
                FormulaValues = formulaValues
            };

            await ReportNavigationService.PrintReportAsync(request);
        }

        private async Task OnEditVoucherClick(BlazorContextMenu.ItemClickEventArgs e)
        {
            var row = e.Data as Cash_Book_Report_ViewModel;
            if (row != null)
            {
                // Navigation or modal for Edit Voucher
            }
        }
    }
}
