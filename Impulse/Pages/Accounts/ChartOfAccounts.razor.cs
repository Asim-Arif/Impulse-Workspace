using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using System.Collections.ObjectModel;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;

namespace Impulse.Pages.Accounts
{
    public partial class ChartOfAccounts : ComponentBase
    {
        private List<ChartOfAccountsModel> ChartOfAccountsfromDB { get; set; } = new List<ChartOfAccountsModel>();
        [Inject]
        protected IChartOfAccountsDataAccess iChartOfAccountsDataAccess { get; set; }
        
        [Inject]
        protected IDBHelper IDBHelper { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        protected IReportNavigationService ReportNavigationService { get; set; }

        private ChartOfAccountsModel SelectedAccount;
        private ChartOfAccountsModel CurrentAccount;
        private bool ShowModal;
        private bool IsLoading;
        private bool IsAdding;
        private bool IsEdit;
        private bool IsContextMenuVisible; // Renamed from ShowContextMenu to avoid conflict
        private int ContextMenuLeftpx;
        private int ContextMenuToppx;
        private ElementReference gridRef;
        private bool bshowinactive=false;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                var AccountsFromDb = await iChartOfAccountsDataAccess.GetChartOfAccountsList(bshowinactive);
                ChartOfAccountsfromDB = AccountsFromDb.ToList();
                OnAfterRenderAsync(true);
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }
        private async Task Refreshlist(bool showinactive)
        {
            IsLoading = true;
            try
            {
                var AccountsFromDb = await iChartOfAccountsDataAccess.GetChartOfAccountsList(showinactive);
                ChartOfAccountsfromDB = AccountsFromDb.ToList();
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
        private void SelectAccount(ChartOfAccountsModel account)
        {
            SelectedAccount = account;
        }
        private async Task ShowAddModal()
        {
            CurrentAccount = new ChartOfAccountsModel { };
            IsAdding = true;
            ShowModal = true;
            accheadsforlist = await iChartOfAccountsDataAccess.GetAccountsHeads();
            CurrentAccount.OpenDate= DateTime.Today;
            CurrentAccount.BalType = 1;
        }

        //private ConfirmDialog msgdialog = null!;
        //[Inject] ToastService ToastService { get; set; } = default!;
        [Inject]
        protected Radzen.NotificationService Notification { get; set; }
        private async Task DeleteSelectedAccount(ItemClickEventArgs e)
        {
            try
            {
                StateHasChanged();

                var account = e.Data as ChartOfAccountsModel;
                SelectedAccount = account;

                if (SelectedAccount == null)
                    return;

                int isExitInTransaction = 0;

                // Check if it is a parent account
                int bParentAccount = await IDBHelper.getSingleIntValueasync(
                    "Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");

                // Check if account is fixed
                isExitInTransaction = await IDBHelper.getSingleIntValueasync(
                    "COUNT(*)", "FixedAccounts", $"WHERE AccNo = '{SelectedAccount.AccNo}'");

                if (isExitInTransaction > 0)
                {
                    NotificationServiceManager.ShowWarning("Cannot delete this account.", "This account has a fixed account number.");
                    return;
                }

                // Check if parent account has child accounts
                if (bParentAccount == 1)
                {
                    isExitInTransaction = await IDBHelper.getSingleIntValueasync(
                        "COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}'");

                    if (isExitInTransaction > 0)
                    {
                        NotificationServiceManager.ShowWarning("Cannot delete this account.", "This account has child accounts.");
                        return;
                    }
                }

                isExitInTransaction = await IDBHelper.getSingleIntValueasync(
                    "COUNT(*)", "Vouchers", $"WHERE AccNo = '{SelectedAccount.AccNo}'");

                if (isExitInTransaction > 0)
                {
                    NotificationServiceManager.ShowWarning("Cannot delete this account.", "There are transactions linked to this account.");
                    return;
                }

                await ChartOfAccountsService.DeleteAccountAsync(SelectedAccount.AccNo);
                await Refreshlist(bshowinactive);
                SelectedAccount = null;

                NotificationServiceManager.ShowSuccess("Account Deleted", "The selected account has been successfully removed.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Occurred", $"Failed to delete account: {ex.Message}");
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
        }

        private void ViewDetails()
        {
            if (SelectedAccount != null)
            {
                Console.WriteLine($"Viewing details for {SelectedAccount.AccTitle}");
                IsContextMenuVisible = false; // Updated to new variable name
            }
        }

        private async Task SaveAccount()
        {
            if (IsAdding)
            {
                if (string.IsNullOrWhiteSpace(CurrentAccount.AccTitle))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Proceed Without Account Title.");
                    return;
                }
                if (CurrentAccount.BalType == 2)
                {
                    CurrentAccount.Balance = -CurrentAccount.Balance;
                }
                else
                {
                    CurrentAccount.Balance = Math.Abs(CurrentAccount.Balance);
                }

                await GetNextAccountNo(CurrentAccount.HeadTypeNo, CurrentAccount.AccTitle, CurrentAccount.AccType, CurrentAccount.SubAccOf);
                await ChartOfAccountsService.SaveNewAccount(CurrentAccount);
            }
            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentAccount.AccTitle))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Proceed Without Account Title.");
                    return;
                }
                if (CurrentAccount.BalType == 2)
                {
                    CurrentAccount.Balance = -CurrentAccount.Balance;
                }
                else {
                    CurrentAccount.Balance = Math.Abs(CurrentAccount.Balance);
                }
                await ChartOfAccountsService.EditAccount(CurrentAccount, CurrentAccount.AccNo);
                await Refreshlist(bshowinactive);
            }
            HideModal();
            ResetForm();
        }

        private void ResetForm()
        {
            CurrentAccount.AccTitle= "";
            CurrentAccount.AccNo = "";
            CurrentAccount.OpeningBalance = 0;
            Selectedhead = null;
            Selectedhead_sub = null;

            StateHasChanged();
        }
        private void HideModal()
        {
            ShowModal = false;
            //CurrentAccount = null;
            IsEdit = false;
        }
        //////////////////// NOW GET ACCOUNT HEADS DATA 
        private List<ChartOfAccountsModel> accheadsforlist = new List<ChartOfAccountsModel>();
        private ChartOfAccountsModel? Selectedhead = null;
        private async Task<IEnumerable<ChartOfAccountsModel>> GetHeads(string searchText)
        {
            return await Task.FromResult(accheadsforlist.Where(x => x.HeadType.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged(ChartOfAccountsModel? selectedhead)
        {
            Selectedhead = selectedhead;
            CurrentAccount.HeadType= selectedhead.HeadType;
            CurrentAccount.HeadTypeNo = selectedhead.HeadTypeNo;
            if (CurrentAccount.HeadTypeNo != null)
            {
                accheadsforlist_sub = await iChartOfAccountsDataAccess.GetSubAccOfAccounts(CurrentAccount.HeadTypeNo);
            }
            else
            {
                CurrentAccount.SubAccOf = "";
            }
            if (CurrentAccount.AccTitle != null && CurrentAccount.HeadTypeNo != null)
            {
                await GetNextAccountNo(CurrentAccount.HeadTypeNo, CurrentAccount.AccTitle, false, CurrentAccount.SubAccOf);
            }
        }
        //////////////////// NOW GET SUB ACCOUNT HEADS DATA 
        private List<ChartOfAccountsModel> accheadsforlist_sub = new List<ChartOfAccountsModel>();
        private ChartOfAccountsModel? Selectedhead_sub = null;
        private async Task<IEnumerable<ChartOfAccountsModel>> GetHeads_sub(string searchText)
        {
           return await Task.FromResult(accheadsforlist_sub.Where(x => x.SubAccOf.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged_sub(ChartOfAccountsModel? selectedhead_sub)
        {
            Selectedhead_sub = selectedhead_sub;
            CurrentAccount.SubAccOf = selectedhead_sub.SubAccOf;
            CurrentAccount.SubAccTitle = selectedhead_sub.SubAccTitle;
            //accheadsforlist_sub = await iChartOfAccountsDataAccess.GetSubAccOfAccounts(CurrentAccount.HeadTypeNo);
            if (CurrentAccount.AccTitle != null && CurrentAccount.HeadTypeNo != null && CurrentAccount.SubAccOf != null)
            { 
                await GetNextAccountNo(CurrentAccount.HeadTypeNo,CurrentAccount.AccTitle, CurrentAccount.AccType, CurrentAccount.SubAccOf);
            }
        }
        private async Task GetNextAccountNo(string acctype, string accountname, Boolean isparent, string subaccof)
        {
            string NextAccountNo = await iChartOfAccountsDataAccess.GetNextAccountNumberAsync(acctype, accountname, isparent, subaccof);
            CurrentAccount.AccNo = NextAccountNo;
        }
        private async Task OnChangedAccTitle(ChangeEventArgs e)
        {
            if(IsEdit == false)
            { 
                await GetNextAccountNo(CurrentAccount.HeadTypeNo, CurrentAccount.AccTitle, CurrentAccount.AccType, CurrentAccount.SubAccOf);
            }
        }
        private async Task OnChangedMasterType(ChangeEventArgs e)
        {
            await GetNextAccountNo(CurrentAccount.HeadTypeNo, CurrentAccount.AccTitle, CurrentAccount.AccType, CurrentAccount.SubAccOf);
            if (CurrentAccount.AccType == true)
            {
                CurrentAccount.Balance = 0;
            }
        }
        private void OnChangedBalanceType(int value)
        {
            CurrentAccount.BalType = value;
        }
        void ViewAccount(ItemClickEventArgs e)
        {
            var account = e.Data as ChartOfAccountsModel;
            Console.WriteLine($"Viewing: {account?.AccNo}");
        }

        async void RefreshData(ItemClickEventArgs e)
        {
            await Refreshlist(bshowinactive);
            StateHasChanged();
        }
        async void ShowInActive(ItemClickEventArgs e, bool showinactive)
        {
            bshowinactive = showinactive;
            await Refreshlist(bshowinactive);
            StateHasChanged();
        }
        private async void EditSelectedAccount(ItemClickEventArgs e)
        {
            var account = e.Data as ChartOfAccountsModel;
            SelectedAccount = account;

            if (SelectedAccount != null)
            {
                accheadsforlist = await iChartOfAccountsDataAccess.GetAccountsHeads();

                if (SelectedAccount.Balance < 0)
                    SelectedAccount.BalType = 2; // Credit
                else
                    SelectedAccount.BalType = 1; // Debit

                CurrentAccount = new ChartOfAccountsModel
                {
                    AccNo = SelectedAccount.AccNo,
                    AccTitle = SelectedAccount.AccTitle,
                    Balance = SelectedAccount.Balance,
                    SubAccOf = SelectedAccount.SubAccOf,
                    SubAccTitle = SelectedAccount.SubAccTitle,
                    HeadTypeNo = SelectedAccount.HeadTypeNo,
                    HeadType = SelectedAccount.HeadType,
                    OpeningBalance = SelectedAccount.OpeningBalance,
                    BalType = SelectedAccount.BalType,
                    OpenDate = SelectedAccount.OpenDate
                };
                
                IsAdding = false;
                ShowModal = true;
                IsContextMenuVisible = false;
                IsEdit = true;
                StateHasChanged();
                //Console.WriteLine($"BalType Value Before Modal: {SelectedAccount.BalType}");
            }
        }

        private Char dialog = default!;

        //private async void MakeAccountActive(ItemClickEventArgs e)
        private async Task MakeAccountActive(ItemClickEventArgs e)

        {
            try
            {
                var account = e.Data as ChartOfAccountsModel;
                SelectedAccount = account;

                if (SelectedAccount == null)
                    return;

                int isExitInTransaction = 0;
                bool bactive = SelectedAccount.Active;
                if (bactive == true)
                {
                    int bParentAccount = await IDBHelper.getSingleIntValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                    isExitInTransaction = await IDBHelper.getSingleIntValueasync("COUNT(*)", "FixedAccounts", $"WHERE AccNo = '{SelectedAccount.AccNo}'");

                    if (isExitInTransaction > 0)
                    {
                        NotificationServiceManager.ShowWarning("Cannot Make This Account Inactive."," This Account Has A Fixed Account Number.");
                        return;
                    }

                    if (bParentAccount == 1)
                    {
                        isExitInTransaction = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}'");
                        if (isExitInTransaction > 0)
                        {
                            NotificationServiceManager.ShowWarning("Cannot Make This Account InActive.", "This Account Has Some Child Accounts.");
                            return;
                        }
                    }

                    isExitInTransaction = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Vouchers", $"WHERE AccNo = '{SelectedAccount.AccNo}'");
                    if (isExitInTransaction > 0)
                    {
                        NotificationServiceManager.ShowWarning("Cannot Make InActive.", "There is Some Transaction Related To This Account.");
                        return;
                    }

                    await ChartOfAccountsService.UpdateAccountStatusAsync(SelectedAccount.AccNo,SelectedAccount.Active);
                    NotificationServiceManager.ShowSuccess("Account Marked InActive", "Account Status Changed.");
                }
                else {
                    await ChartOfAccountsService.UpdateAccountStatusAsync(SelectedAccount.AccNo, SelectedAccount.Active);
                    NotificationServiceManager.ShowSuccess("Account Marked Active", "Account Status Changed.");
                }

                string action = SelectedAccount.Active ? "Make This Account Inactive" : "Activate This Account";
                StateHasChanged();

                StateHasChanged();
                await Refreshlist(bshowinactive);
                StateHasChanged();

            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Occurred", $"Failed To Change Account Status account: {ex.Message}");
            }

        }
        private void GoToIndexPage()
        {
            // Navigate to the index page ("/" is typically the root URL)
            Navigation.NavigateTo("/", true);
        }


        private void PrintChartOfAccounts(int itype)
        {
            string strcond = "";
            if (bshowinactive)
            {
                strcond = "1=1";   // show all records
            }
            else
            {
                strcond = "{VTempAccounts.Active} = 1";
            }

            string strreportname = "";
            if (itype == 1)
            {
                strreportname= ReportNames.Accounts.PrintAccounts;
            }
            else
            {
                strreportname = ReportNames.Accounts.AccountsWithOpening;
            }

            var reportRequest = new ReportRequest
            {

                ReportName = strreportname,
                SelectionFormula = strcond,
            };
            ReportNavigationService.PrintReportAsync(reportRequest);

        }

    }
}
