using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Radzen;
using System.Collections.Generic;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.JSInterop;

namespace Impulse.Pages.Accounts
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
        [Inject] 
        private AuthenticationStateProvider AuthStateProvider { get; set; }
        [Inject] 
        private IJSRuntime JSRuntime { get; set; }
        
        private async Task<string> GetUserName()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            return authState.User.Identity?.Name ?? "Guest";
        }

        private List<GenericDropDownModel> Accounts = new List<GenericDropDownModel>();
        private List<AccountsReportingModel> AccountsList = new List<AccountsReportingModel>();
        private AccountsReportingModel CurrentAccount = new AccountsReportingModel();
        private List<AccountsReportingModel> LedgerDatafromDB { get; set; } = new List<AccountsReportingModel>();
        private DateTime DtFrom { get; set; } = DateTime.Now;
        private DateTime DtTo { get; set; } = DateTime.Now;

        private string? StrOpeningMsg = null;
        private string? StrClosingMsg = null;

        public bool IsChequeModalVisible { get; set; } = false;
        public AccountsReportingModel? CurrentChequeRow { get; set; } = null;

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
            if (string.IsNullOrWhiteSpace(searchText))
                return AccountsList;

            return await Task.FromResult(AccountsList.Where(x => 
                (!string.IsNullOrEmpty(x.AccTitle) && x.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)) || 
                (!string.IsNullOrEmpty(x.AccNo) && x.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            ).ToList());
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
        private ChequeDetailModel? ChequeDetails = null;

        private async Task ShowChequeDetails(AccountsReportingModel row)
        {
            CurrentChequeRow = row;
            ChequeDetails = await AccountReportingAccess.GetChequeDetails(row.VchrNo, row.ChqNo);
            IsChequeModalVisible = true;
            StateHasChanged();
        }

        private void CloseChequeModal()
        {
            IsChequeModalVisible = false;
            CurrentChequeRow = null;
            StateHasChanged();
        }

        private async Task InsertBalanceTag(ItemClickEventArgs e)
        {
            var row = e.Data as AccountsReportingModel;
            if (row != null)
            {
                var userName = await GetUserName();
                await AccountReportingAccess.InsertBalanceTag(row.SNo, userName);
                row.BalanceTag_UserName = userName;
                row.BalanceTag_DTEntry = DateTime.Now;
                StateHasChanged();
            }
        }

        private async Task RemoveBalanceTag(ItemClickEventArgs e)
        {
            var row = e.Data as AccountsReportingModel;
            if (row != null)
            {
                await AccountReportingAccess.RemoveBalanceTag(row.SNo);
                row.BalanceTag_UserName = null;
                row.BalanceTag_DTEntry = null;
                StateHasChanged();
            }
        }

        private void OnInsertBalanceTagAppearing(ItemAppearingEventArgs e)
        {
            var row = e.Data as AccountsReportingModel;
            if (row != null && !string.IsNullOrEmpty(row.BalanceTag_UserName))
            {
                e.IsVisible = false;
            }
        }

        private void OnRemoveBalanceTagAppearing(ItemAppearingEventArgs e)
        {
            var row = e.Data as AccountsReportingModel;
            if (row != null && string.IsNullOrEmpty(row.BalanceTag_UserName))
            {
                e.IsVisible = false;
            }
        }

        private void PrintVoucher(ItemClickEventArgs e)
        {
            var vchr = e.Data as AccountsReportingModel;
            CurrentAccount = vchr;

            if (CurrentAccount != null && !string.IsNullOrEmpty(CurrentAccount.VchrNo))
            {
                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Voucher,
                    SelectionFormula = $"{{VLedger.VchrNo}} = '{CurrentAccount.VchrNo}'"
                };
                
                ReportNavigationService.PrintReportAsync(reportRequest);
            }
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
                    SelectionFormula = $"{{VLedger.AccNo}} = '{CurrentAccount.AccNo}' AND {{VLedger.VDate}} in Date({CurrentAccount.DTFrom.Year}, {CurrentAccount.DTFrom.Month}, {CurrentAccount.DTFrom.Day}) to Date({CurrentAccount.DTTo.Year}, {CurrentAccount.DTTo.Month}, {CurrentAccount.DTTo.Day})",
                };

                ReportNavigationService.PrintReportAsync(reportRequest);
            }
        }

        private void PrintLedger(ItemClickEventArgs e)
        {
            if (CurrentAccount != null && !string.IsNullOrEmpty(CurrentAccount.AccNo))
            {
                decimal openingBalance = CurrentAccount.OpeningBalance;
                string openCrDr = openingBalance >= 0 ? "Dr" : "Cr";
                string dateRange = $"From {CurrentAccount.DTFrom:dd-MMM-yyyy} To {CurrentAccount.DTTo:dd-MMM-yyyy}";
                string selectionFormula = $"{{VLedger.AccNo}}='{CurrentAccount.AccNo}' AND {{VLedger.VDate}} in Date({CurrentAccount.DTFrom.Year}, {CurrentAccount.DTFrom.Month}, {CurrentAccount.DTFrom.Day}) to Date({CurrentAccount.DTTo.Year}, {CurrentAccount.DTTo.Month}, {CurrentAccount.DTTo.Day})";

                var reportRequest = new ReportRequest
                {
                    ReportName = "Ledger.rpt",
                    SelectionFormula = selectionFormula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Company", "'IAA'" },
                        { "OpeningBalance", openingBalance },
                        { "OpenCrDr", $"'{openCrDr}'" },
                        { "FromTo", $"'{dateRange}'" }
                    }
                };

                ReportNavigationService.PrintReportAsync(reportRequest);
            }
        }

    }
}