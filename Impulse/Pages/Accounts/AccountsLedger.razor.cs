using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Constants;
using CIP.Pages;
using CIP.Services;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;
using Radzen;
using System.Collections.Generic;

namespace CIP.Pages.Accounts
{
    public partial class AccountsLedger : ComponentBase
    {
        private bool IsLoading=true;
        [Inject]
        private IVouchersDataAccess IVouchersDataAccess { get; set; }
        [Inject]
        private IDBHelper IDBHelper { get; set; }
        [Inject]
        protected IAccountReportingAccess AccountReportingAccess { get; set; }

        [Inject]
        protected IReportNavigationService ReportNavigationService { get; set; }
        private List<GenericDropDownModel> Accounts = new List<GenericDropDownModel>();
        private List<AccountsReportingModel> AccountsList = new List<AccountsReportingModel>();
        private AccountsReportingModel CurrentAccount = new AccountsReportingModel();
        private List<AccountsReportingModel> LedgerDatafromDB { get; set; } = new List<AccountsReportingModel>();
        private DateTime DtFrom { get; set; } = DateTime.Now;
        private DateTime DtTo { get; set; } = DateTime.Now;

        private string? StrOpeningMsg = null;
        private string? StrClosingMsg = null;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                IsLoading = true;
                Accounts = await IVouchersDataAccess.GetValuesForDropDown("Accounts", "AccNo", "AccTitle", " WHERE Parent=0 ORDER BY AccNo");
                string maxDateString = await IDBHelper.getSingleStringValue("MAX(VDate)", "Vouchers");

                CurrentAccount.DTFrom = new DateTime(DateTime.Now.Year, 7, 1);
                CurrentAccount.DTTo = DateTime.Parse(maxDateString);

                AccountsList = Accounts
                    .Select(a => new AccountsReportingModel
                    {
                        AccNo = a.DropDownValue_ID,
                        AccTitle = a.DropDownValue_Description 
                    })
                    .ToList();

                IsLoading = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }
        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        //////////////////// NOW GET ACCOUNT NOs DATA 
        private List<AccountsReportingModel> accheadsforlist = new List<AccountsReportingModel>();
        private AccountsReportingModel? SelectedAccount = null;

        private async Task<IEnumerable<AccountsReportingModel>> GetAccounts(string searchText)
        {
            return await Task.FromResult(AccountsList.Where(x => x.AccTitle.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged(AccountsReportingModel? selectedhead)
        {
            SelectedAccount = selectedhead;
            CurrentAccount.AccNo = selectedhead?.AccNo ?? string.Empty;
        }

        private decimal? TotalDebit = 0;
        private decimal? TotalCredit = 0;

        private async Task RefreshLedger()
        {
            IsLoading = true;
            try
            {
                StrOpeningMsg = "< < <    Balance Forward   > > >";
                StrClosingMsg = "< < <    Last Balance   > > >";

                LedgerDatafromDB.Clear();

                decimal AccBal = await AccountReportingAccess.GetAccountOpeningBalance(CurrentAccount.AccNo,CurrentAccount.DTFrom);
                CurrentAccount.OpeningBalance = AccBal;
                var LedgerDataFromDb = await AccountReportingAccess.GetLedgerData(CurrentAccount.AccNo,CurrentAccount.DTFrom,CurrentAccount.DTTo);
                LedgerDatafromDB = LedgerDataFromDb.ToList();

                TotalDebit = LedgerDatafromDB.Sum(x => x.Debit);
                TotalCredit = LedgerDatafromDB.Sum(x => x.Credit);

                decimal ClBal = await AccountReportingAccess.GetAccountOpeningBalance(CurrentAccount.AccNo, CurrentAccount.DTTo);
                CurrentAccount.Balance = ClBal;
                
                StateHasChanged();

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private void GoToVoucherTransaction(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;
            CurrentAccount = vchr;

            if (CurrentAccount != null)
            {
                Navigation.NavigateTo($"/accounts/transactionregister/{CurrentAccount.VchrNo}");
            }
        }
        private void PrintVoucher(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;
            CurrentAccount = vchr;

            if (CurrentAccount != null)
            {
                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Voucher, //"PV.rpt",
                    Format = "xlsx",
                    SelectionFormula = $"{{VLedger.VchrNo}} = '{CurrentAccount.VchrNo}'",
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Test", "\"From\"" }
                    }
                };
                
                ReportNavigationService.PrintReportAsync(reportRequest);
                
            }
            //use below code example for Parameterized reports.
            /*DateTime myDT = DateTime.Now;
            var reportRequest = new ReportRequest
            {
                ReportName = "Yearly_PLS_Comparison.rpt",
                //SelectionFormula = $"{{VLedger.VchrNo}} = '{CurrentAccount.VchrNo}'",
                Parameters = new Dictionary<string, object>
                    {
                        { "@Year", 2024 }
                    }
            };

            ReportNavigationService.PrintReportAsync(reportRequest);
            */
        }

        private void PrintLedgerWithTitle(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;

            if (CurrentAccount != null)
            {

                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Accounts.LedgerWithTitle,
                    SelectionFormula = $"1=1",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", CurrentAccount.DTFrom },
                        { "@DTTo", CurrentAccount.DTTo},
                        { "@AccNo", CurrentAccount.AccNo},
                        { "@SubAccOf", ""}
                    }
                };

                ReportNavigationService.PrintReportAsync(reportRequest);
            }
        }

        private void PrintTransactionReport(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;

            if (CurrentAccount != null)
            {

                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Accounts.Transaction_Report,
                    SelectionFormula = $"{{VLedger.AccNo}} = '{CurrentAccount.AccNo}' AND {{VLedger.VDate}}=#{CurrentAccount.DTFrom}# TO #{CurrentAccount.DTTo}#",
                };

                ReportNavigationService.PrintReportAsync(reportRequest);
            }
        }

        private void PrintLedger(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;

            decimal openingBalance = CurrentAccount.OpeningBalance;
            string dateRange = $"'From {(CurrentAccount.DTFrom.ToString("dd-MMM-yyyy") ?? "")} To {(CurrentAccount.DTTo.ToString("dd-MMM-yyyy") ?? "")}'";
            if (CurrentAccount != null)
            {
                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Accounts.PrintLedgerReport,
                    SelectionFormula = $"1=1",
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", CurrentAccount.DTFrom },
                        { "@DTTo", CurrentAccount.DTTo},
                        { "@AccNo", CurrentAccount.AccNo},

                        { "@SubAccOf", ""}
                    }
                    ,
                    FormulaValues = new Dictionary<string, object>
                    {
                        //{ "AccNo",  CurrentAccount.AccNo },
                        //{ "AccTitle",  CurrentAccount.AccTitle }
                        //,
                        { "OpeningBalance",  openingBalance },
                        { "OpenCrDr",  openingBalance },
                        { "FromTo",  dateRange }
                    }
                };

                ReportNavigationService.PrintReportAsync(reportRequest);
            }
        }

    }
}