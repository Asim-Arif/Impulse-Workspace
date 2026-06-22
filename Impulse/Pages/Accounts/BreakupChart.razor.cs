using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.VariantTypes;
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
    public partial class BreakupChart : ComponentBase
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
        private bool bshowinactive = false;

        private DateTime TodayDT = DateTime.Now;
        private DateTime myDTStart = DateTime.Now.AddMonths(-1);

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
            CurrentAccount.OpenDate = DateTime.Today;
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
                    "COUNT(*)", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}' AND Parent=1");
                
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
                else
                {
                    CurrentAccount.Balance = Math.Abs(CurrentAccount.Balance);
                }
                await ChartOfAccountsService.EditAccount(CurrentAccount, CurrentAccount.AccNo);
                //await Refreshlist(bshowinactive);
            }
            await Refreshlist(bshowinactive);
            HideModal();
            ResetForm();
        }

        private void ResetForm()
        {
            CurrentAccount.AccTitle = "";
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
        private List<CashFlowHeadsModel> cashflowheadslist = new List<CashFlowHeadsModel>();
        private ChartOfAccountsModel? Selectedhead = null;

        private List<NotesToAccountsModel> notestoaccountsheadslist = new List<NotesToAccountsModel>();
        private ChartOfAccountsModel? SelectedNote = null;

        private async Task<IEnumerable<ChartOfAccountsModel>> GetHeads(string searchText)
        {
            return await Task.FromResult(accheadsforlist.Where(x => x.HeadType.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged(ChartOfAccountsModel? selectedhead)
        {
            Selectedhead = selectedhead;
            CurrentAccount.HeadType = selectedhead.HeadType;
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
                await GetNextAccountNo(CurrentAccount.HeadTypeNo, CurrentAccount.AccTitle, CurrentAccount.AccType, CurrentAccount.SubAccOf);
            }
        }
        private async Task GetNextAccountNo(string acctype, string accountname, Boolean isparent, string subaccof)
        {
            string NextAccountNo = await iChartOfAccountsDataAccess.GetNextAccountNumberAsync(acctype, accountname, isparent, subaccof);
            CurrentAccount.AccNo = NextAccountNo;
        }
        private async Task OnChangedAccTitle(ChangeEventArgs e)
        {
            if (IsEdit == false)
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
                        NotificationServiceManager.ShowWarning("Cannot Make This Account Inactive.", " This Account Has A Fixed Account Number.");
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

                    await ChartOfAccountsService.UpdateAccountStatusAsync(SelectedAccount.AccNo, SelectedAccount.Active);
                    NotificationServiceManager.ShowSuccess("Account Marked InActive", "Account Status Changed.");
                }
                else
                {
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
        private void PrintAccountsHeadsReport(int iTypes)
        {
            //if (CurrentAccount != null)
            //{
            DateTime NowDT = DateTime.Now;
            DateTime myDT = DateTime.Now.AddMonths(-1);
            string cond = "";
            if (iTypes == 0)
            {
                cond = "{VTempAccounts.Active}=TRUE";
            }
            else if (iTypes == 1)
            {
                cond = "{VTempAccounts.Active}=TRUE AND ({VTempAccounts.Marked}=1 OR {VTempAccounts.ParentMarked}=1)";
            }
            else if (iTypes == 2)
            {
                cond = "({VTempAccounts.opendate}>=#" + myDT + "# And {VTempAccounts.opendate} <= #" + NowDT + " #) And ({VTempAccounts.Marked}=1 And {VTempAccounts.ManualMarked}=0 And {VTempAccounts.ParentMarked}=0)";
            }
            else
            {
                cond = "{VTempAccounts.Active}=TRUE AND {VTempAccounts.Marked}=0 AND {VTempAccounts.ParentMarked}=0 AND {VTempAccounts.ManualMarked}=0";
            }
            //ReportNavigationService.PrintReportAsync("Accounts_Adjustment.rpt", cond);
            ReportNavigationService.PrintReportAsync_Old("AccountsHeadsReport.rpt", cond);
            //}
        }

        private bool ShowModal_CF;
        private bool IsEdit_CF;
        private CashFlowHeadsModel? SelectedHead_CF = null;
        private async Task ShowCashFlowModal()
        {
            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }
            bool parentaccount = false;
            string strsubaccof = "";
            int icount = 0;
            parentaccount = await IDBHelper.getSingleBoolValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            if (parentaccount == false)
            {
                strsubaccof = await IDBHelper.getSingleStringValue("SubAccOf", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo= '{strsubaccof}' AND ISNULL(CFHFS_RefID,0)<>0");
                if (icount > 0)
                {
                    NotificationServiceManager.ShowWarning("alert", "Cannot Proceed. Master Account Already Marked.");
                    return;
                }
            }
            else
            {
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}' AND ISNULL(CFHFS_RefID,0)<>0");
            }

            if (icount == 0)
            {
                if (parentaccount == true)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}' AND ISNULL(CFHFS_RefID,0)<>0 ");
                }
                else
                {
                    icount = 0;
                }
            }

            if (parentaccount == true)
            {
                if (icount == 0)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE ISNULL(Marked_CF,0)<>0 AND AccNo IN (SELECT AccNo FROM Accounts WHERE SubAccOf='{SelectedAccount.AccNo}' AND ISNULL(Marked_CF,0)<>0)");
                }
            }
            else
            {
                icount = 0;
            }

            if (icount > 0)
            {
                NotificationServiceManager.ShowWarning("alert", "Can't Add Related Account Already Marked For Cash Flow.");
                return;
            }

            CurrentAccount = SelectedAccount;
            SelectedAccount = null;

            IsEdit_CF = true;
            ShowModal_CF = true;
            cashflowheadslist = await iChartOfAccountsDataAccess.GetCashFlowHeads();

            var cfhead = new CashFlowHeadsModel();
            if (CurrentAccount.CFHFS_RefID != 0)
            {
                cfhead.EntryID = CurrentAccount.CFHFS_RefID;
                cfhead.Description = CurrentAccount.Description;
                SelectedHead_CF = cfhead;
            }
        }

        private void HideCashFlowModal()
        {
            ShowModal_CF = false;
            IsEdit_CF = false;
        }
        private async Task<IEnumerable<CashFlowHeadsModel>> SearchCashFlowHeads(string searchText)
        {
            return await Task.FromResult(cashflowheadslist.Where(x => x.Description.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged_CF(CashFlowHeadsModel? selectedhead)
        {
            SelectedHead_CF = selectedhead;
        }
        private async Task SaveCashFlowHeads()
        {
            if (SelectedHead_CF != null)
            {
                CurrentAccount.CFHFS_RefID = SelectedHead_CF.EntryID;
                CurrentAccount.Description = SelectedHead_CF.Description;
            }

            if (IsEdit_CF)
            {
                if (SelectedHead_CF.EntryID == 0)
                {
                    NotificationServiceManager.ShowWarning("alert", "Cannot Proceed Without Cash Flow Heads.");
                    return;
                }
                await ChartOfAccountsService.UpdateCashFlowHeads(CurrentAccount, CurrentAccount.AccNo, CurrentAccount.CFHFS_RefID, 3);
                await Refreshlist(bshowinactive);
            }
            HideCashFlowModal();

            StateHasChanged();
            await Refreshlist(bshowinactive);
            StateHasChanged();

        }

        private async Task MarkCashFlowAccount()
        {
            //CurrentAccount = new ChartOfAccountsModel { };

            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }
            bool parentaccount = false;
            string strsubaccof = "";
            int icount = 0;
            parentaccount = await IDBHelper.getSingleBoolValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            if (parentaccount == false)
            {
                strsubaccof = await IDBHelper.getSingleStringValue("SubAccOf", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo= '{strsubaccof}' AND ISNULL(Marked_CF,0)<>0");
                if (icount > 0)
                {
                    NotificationServiceManager.ShowWarning("alert", "Cannot Proceed. Master Account Already Marked.");
                    return;
                }
            }
            else
            {
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}' AND ISNULL(Marked_CF,0)<>0");
            }


            if (icount == 0)
            {
                if (parentaccount == true)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}' AND ISNULL(Marked_CF,0)<>0 ");
                }
                else
                {
                    icount = 0;
                }
            }

            if (parentaccount == true)
            {

                if (icount == 0)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE ISNULL(CFHFS_RefID,0)<>0 AND AccNo IN (SELECT AccNo FROM Accounts WHERE SubAccOf='{SelectedAccount.AccNo}' AND ISNULL(Marked_CF,0)<>0)");
                }
            }
            else
            {
                icount = 0;
            }
            if (icount > 0)
            {
                NotificationServiceManager.ShowWarning("alert", "Can't Add Related Account Already Marked For Cash Flow.");
                return;
            }
            if (SelectedAccount.Marked_CF == true)
            {
                await iChartOfAccountsDataAccess.UpdateCashFlowHeads(SelectedAccount, SelectedAccount.AccNo, 0, 1);
            }
            else
            {
                await iChartOfAccountsDataAccess.UpdateCashFlowHeads(SelectedAccount, SelectedAccount.AccNo, 0, 2);
            }
            CurrentAccount = SelectedAccount;

            StateHasChanged();
            await Refreshlist(bshowinactive);
            StateHasChanged();

            SelectedAccount = null;
        }

        private bool ShowModal_Notes;
        private bool IsEdit_Notes;
        private NotesToAccountsModel? SelectedHead_Notes = null;
        private async Task ShowNotesToAccountsModal()
        {
            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }
            bool parentaccount = false;
            string strsubaccof = "";
            int icount = 0;
            strsubaccof = await IDBHelper.getSingleStringValue("SubAccOf", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            parentaccount = await IDBHelper.getSingleBoolValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            if (parentaccount == false)
            {
                strsubaccof = await IDBHelper.getSingleStringValue("SubAccOf", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo= '{strsubaccof}' AND Marked=1");
                if (icount > 0)
                {
                    NotificationServiceManager.ShowWarning("alert", "Cannot Proceed. Master Account Already Marked.");
                    return;
                }
            }
            else
            {
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}' AND Marked=1");
            }

            if (icount == 0)
            {
                if (parentaccount == true)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}' AND Marked=1 ");
                }
                else
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}' AND Marked=1 ");
                }
            }

            if (parentaccount == true)
            {
                if (icount == 0)
                {
                    icount = 0;// await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE ISNULL(Marked_CF,0)<>0 AND AccNo IN (SELECT AccNo FROM Accounts WHERE SubAccOf='{SelectedAccount.AccNo}' AND ISNULL(Marked_CF,0)<>0)");
                }
            }
            else
            {
                icount = 0;
            }

            if (icount > 0)
            {
                NotificationServiceManager.ShowWarning("alert", "Can't Add Related Account Already Marked.");
                return;
            }

            CurrentAccount = SelectedAccount;
            SelectedAccount = null;

            IsEdit_Notes = true;
            ShowModal_Notes = true;
            notestoaccountsheadslist = await iChartOfAccountsDataAccess.GetNotesToAccountsHeads();

            var notes = new NotesToAccountsModel();
            if (CurrentAccount.CFHFS_RefID != 0)
            {
                notes.EntryID = CurrentAccount.AccGroup_RefID;
                notes.GroupName = CurrentAccount.GroupName;
                SelectedHead_Notes = notes;
            }
        }
        private void HideNotesToAccountsAccountModal()
        {
            ShowModal_Notes = false;
            IsEdit_Notes = false;
        }
        private async Task SaveNotesToAccountsAccount()
        {
            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }
            else if (SelectedHead_Notes == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Cannot Proceed Without Selecting Notes To Accounts Group.");
                return;
            }

            bool parentaccount = false;
            string strsubaccof = "";
            int icount = 0;
            parentaccount = await IDBHelper.getSingleBoolValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            if (parentaccount)
            {
                if (parentaccount == true & SelectedAccount.SubAccOf != "")
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "VAccountGroupsForBalancesheet_Accounts", $"WHERE SubAccOf= '{SelectedAccount.AccNo}'");
                }
                else
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "VAccountGroupsForBalancesheet_Accounts", $"WHERE LEFT(SubAccOf,6)= '{SelectedAccount.AccNo}'");
                }
            }
            else
            {
                icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "VAccountGroupsForBalancesheet_Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                if (icount == 0 & parentaccount == false)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "VAccountGroupsForBalancesheet_Accounts", $"WHERE AccNo= '{strsubaccof}'");
                }
            }

            if (icount > 0 & parentaccount == true)
            {
                NotificationServiceManager.ShowWarning("alert", "Child Account Already Added Agaist This Account... You Must Remove Child Account First To Add Master Account...!!!");
                return;
            }
            else if (icount > 0 & parentaccount == false)
            {
                NotificationServiceManager.ShowWarning("alert", "Master Account Already Added Agaist This Account... You Must Remove Master Account First To Add Child Account...!!!");
                return;
            }
            if (SelectedHead_Notes != null)
            {
                SelectedAccount.AccGroup_RefID = SelectedHead_Notes.EntryID;
                SelectedAccount.GroupName = SelectedHead_Notes.GroupName;
            }

            if (SelectedAccount.AccNo.Length == 2 & parentaccount)
            {
                await iChartOfAccountsDataAccess.DeleteAccountGroupsForBalancesheet(SelectedAccount.AccNo);
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, SelectedAccount.AccGroup_RefID, 1);
            }
            else if (SelectedAccount.AccNo.Length > 2 & parentaccount)
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, SelectedAccount.AccGroup_RefID, 2);
            }
            else
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, SelectedAccount.AccGroup_RefID, 2);
            }

            CurrentAccount = SelectedAccount;

            StateHasChanged();
            await Refreshlist(bshowinactive);
            HideNotesToAccountsAccountModal();
            StateHasChanged();

            SelectedAccount = null;
        }

        private async Task<IEnumerable<NotesToAccountsModel>> SearchNotesToAccounts(string searchText)
        {
            return await Task.FromResult(notestoaccountsheadslist.Where(x => x.GroupName.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private async Task SelectedResultChanged_Notes(NotesToAccountsModel? selectedhead)
        {
            SelectedHead_Notes = selectedhead;
        }
        private async Task ChangeNotesStatus()
        {
            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }

            if (SelectedAccount.Marked == true)
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 3);
            }
            else
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 4);
            }

            StateHasChanged();
            await Refreshlist(bshowinactive);
            StateHasChanged();
        }
        private async Task ChangeAccountCheckedStatus()
        {
            if (SelectedAccount == null)
            {
                NotificationServiceManager.ShowWarning("alert", "Please Select An Account First.");
                return;
            }

            bool parentaccount = false;
            string strsubaccof = "";
            int icount = 0;
            parentaccount = await IDBHelper.getSingleBoolValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            strsubaccof = await IDBHelper.getSingleStringValue("SubAccOf", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
            if (SelectedAccount.MarkedTB ==false)
            {
                if (parentaccount)
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo LIKE  '%{SelectedAccount.AccNo}%' AND Marked=1 AND Parent=0");
                }
                else
                {
                    icount = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE AccNo LIKE  '%{strsubaccof}%' AND Marked=1 AND Parent=1");
                }

                if (icount > 0 & parentaccount == false)
                {
                    NotificationServiceManager.ShowWarning("alert", "Master Account Already Added Agaist This Account... You Must Remove Master Account First To Add Child Account...!!!");
                    return;
                }
                else if (icount > 0 & parentaccount == true)
                {
                    NotificationServiceManager.ShowWarning("alert", "Child Account Already Added Agaist This Account... Child Account Already Added Agaist This Account...");
                    return;
                }
            }
            if (SelectedAccount.MarkedTB == true)
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 5);
            }
            else
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 6);
            }

            StateHasChanged();
            await Refreshlist(bshowinactive);
            StateHasChanged();
        }

        private async Task BreakLink(int itype)
        {
            if (itype==1)
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 7);
            }
            else
            {
                await iChartOfAccountsDataAccess.SaveAccountGroupsForBalancesheet(SelectedAccount, SelectedAccount.AccNo, 0, 8);
            }

            StateHasChanged();
            await Refreshlist(bshowinactive);
            StateHasChanged();
        }

    }
}
