using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security.AccessControl;
using System.Threading.Tasks;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;

namespace Impulse.Pages.Accounts
{
    public partial class ChqBookDetail : ComponentBase, IDisposable
    {

        private List<BankAccountInfo> Accounts = new List<BankAccountInfo>();
        private BankAccountInfo? SelectedAccount = null;
        private BankAccountInfo? SelectedBankAccount = null;
        private ChqBookDetailViewModel? SelectedChqBook = null;
        private ChqBookDetailViewModel ChqBookModel = new ChqBookDetailViewModel();
        private ChqBookDetailViewModel SelectedChqNo = new ChqBookDetailViewModel();

        [Inject]
        private IAuditService _auditService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; }
        [Inject]
        private IChqBookDetail _chqbookService { get; set; }
        [Inject]
        private IDBHelperService _idbHelperService { get; set; }

        private string validationMessage = string.Empty;
        private string successMessage = string.Empty;

        private Timer _timer;
        private bool isSaving = false;

        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private bool IsLoading = false;
        private List<ChqBookDetailViewModel> getDatafromdb_BankAccounts { get; set; } = new List<ChqBookDetailViewModel>();

        protected override async Task OnInitializedAsync()
        {
            await InitializeData();
            strComputerName = _auditService.GetClientIpAddress();
            strUserName = _auditService.GetCurrentUserName();
            //Accounts = await _chqbookService.GetTransactionalAccounts("WHERE Active=1 AND Parent=0");
            Accounts = await _chqbookService.GetBankAccounts();

            //getDatafromdb.Clear();
            //var getdatafromDb = await _chqbookService.GetBankChqBooks(BankAccNo);
            //getDatafromdb = getdatafromDb.ToList();
        }

        private async Task InitializeData()
        {
            isSaving = false;
            RefreshChqBooks("");
        }

        public void Dispose()
        {
            _timer?.Dispose();
        }

        private void HideSuccessMessage(object state)
        {
            InvokeAsync(() =>
            {
                successMessage = string.Empty;
                StateHasChanged();
            });
            _timer?.Dispose();
        }

        private async Task<IEnumerable<BankAccountInfo>> GetTransactionAccounts(string searchText)
        {
            return await Task.FromResult(Accounts.Where(x => x.AccTitle.ToLower().Contains(searchText.ToLower())).ToList());
        }

        private async Task SelectedResultChanged_Account(BankAccountInfo? selectedAccount)
        {
            SelectedAccount = selectedAccount;
            if (selectedAccount != null && selectedAccount.AccTitle != null)
            {
                IsLoading = true;
                RefreshChqBooks(SelectedAccount.AccNo);
            }
        }

        private List<ChqBookDetailViewModel> getDatafromdb { get; set; } = new List<ChqBookDetailViewModel>();
        private async Task RefreshChqBooks(string? BankAccNo)
        {
            IsLoading= true;
            try
            {
                StateHasChanged();

                    getDatafromdb.Clear();
                    var getdatafromDb = await _chqbookService.GetBankChqBooks(BankAccNo);
                    getDatafromdb = getdatafromDb.ToList();
                    
                IsLoading = false;
                StateHasChanged();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
                NotificationServiceManager.ShowError("Error Message",ex.Message);
            }
            finally
            {
                StateHasChanged();
                IsLoading = false;
            }
        }

        private List<BankAccountInfo> getDatafromdb_bankaccounts { get; set; } = new List<BankAccountInfo>();
        private async Task RefreshBankAccounts()
        {
            IsLoading = true;
            try
            {
                StateHasChanged();

                getDatafromdb_bankaccounts.Clear();
                var getdatafromDb = await _chqbookService.GetBankAccountsList();
                getDatafromdb_bankaccounts = getdatafromDb.ToList();

                IsLoading = false;
                StateHasChanged();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
                NotificationServiceManager.ShowError("Error Message", ex.Message);
            }
            finally
            {
                StateHasChanged();
                IsLoading = false;
            }
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private bool IsAdding = false;
        private bool ShowModal = false;
        private bool IsEdit = false;
        private bool isError= false;

        private void ShowAddModal(int OnlineChqBook=0)
        {

            if (string.IsNullOrEmpty(SelectedAccount?.AccNo))
            {
                NotificationServiceManager.ShowWarning("Warning", "Please Select Bank Account For Cheque Book");
                return;
            }

            IsAdding = true;
            ShowModal = true;
            IsEdit = false;
            ChqBookModel.Chq_Type = OnlineChqBook;
        }
        private void HideModal()
        {
            IsAdding = false;
            ShowModal = false;
            IsEdit = false;
            ChqBookModel.Chq_Type = 0;

            ChqBookModel.ManualNo = string.Empty;
            ChqBookModel.StartingFrom = string.Empty;
            ChqBookModel.Chqs = 0;
            ChqBookModel.ChqBookDetail = string.Empty;

        }

        private bool ShowBankModal = false;
        private bool IsEdit_BankAccount = false;

        private void ShowBankAccountModal()
        {
            ShowBankModal = true;
            IsEdit_BankAccount = false;
            RefreshBankAccounts();
        }

        private void HideBankAccountModal()
        {
            ShowBankModal = false;
            IsEdit_BankAccount = false;
        }

        private bool ShowODModel = false;

        private void ShowOverDraftModel(ItemClickEventArgs e)   
        {
            var chqbook = e.Data as BankAccountInfo;
            SelectedBankAccount = chqbook;
            if (SelectedBankAccount != null)
            {
                ShowODModel = true;
            }
        }
        private void HideODModel()
        {
            ShowODModel = false;
        }
        private bool ShowChqsData = false;
        private void ShowChqModel(ItemClickEventArgs e)
        {
            var chqbook = e.Data as ChqBookDetailViewModel;
            SelectedChqBook = chqbook;
            if (SelectedChqBook != null)
            {
                ShowChqsData = true;
                RefreshChqsData(SelectedChqBook.ChqBookNo);
            }
        }
        private void HideChqModel()
        {
            ShowChqsData = false;
            SelectedChqBook = null;
        }

        private List<ChqListModel> getDatafromdb_chqsdata { get; set; } = new List<ChqListModel>();
        private async Task RefreshChqsData(long chqbookno)
        {
            try
            {
                StateHasChanged();
                getDatafromdb_chqsdata.Clear();

                var data = await _chqbookService.GetChqBookChequesList(chqbookno);
                getDatafromdb_chqsdata = data.ToList();

                StateHasChanged();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
                NotificationServiceManager.ShowError("Error Message", ex.Message);
            }
            finally
            {
                StateHasChanged();
            }
        }

        private void MakeBankAccountDefault(ItemClickEventArgs e)
        {
            var bankaccount = e.Data as BankAccountInfo;
            SelectedBankAccount = bankaccount;
            if (SelectedBankAccount != null) 
            { 
                SelectedBankAccount.DefaultBank = true;
                SaveOverDraftData();
                RefreshBankAccounts();
            }
        }

        private void SaveOverDraftData()
        {
            if (isSaving)
                return;
            try
            {
                isSaving = true;

                _chqbookService.SaveOverDraftData(SelectedBankAccount, SelectedBankAccount.AccNo, SelectedBankAccount.DefaultBank);

                isError = false;
                NotificationServiceManager.ShowSuccess("Saved", "Updated Successfully.");
                isSaving = false;
                HideODModel();
            }
            catch (Exception ex)
            {
                isError = true;
                NotificationServiceManager.ShowError("Error", $"{ex.Message}");
                isSaving = false;
            }
            finally
            {
                isSaving = false;
            }
        }
        private async Task SaveChqBook()
        {
        if (isSaving)
            return;
            try
            {
                isSaving = true;

                if (ChqBookModel == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Bank Account.");
                    return;
                }
                if (string.IsNullOrEmpty(ChqBookModel.StartingFrom))
                {
                    NotificationServiceManager.ShowWarning("Error", "Please Starting From.");
                    return;
                }
                if ((ChqBookModel.Chqs==0))
                {   
                    NotificationServiceManager.ShowWarning("Error", "Please Enter No Of Cheques");
                    return;
                }
                if (string.IsNullOrEmpty(ChqBookModel.ManualNo))
                {
                    NotificationServiceManager.ShowWarning("Error", "Please Enter Manual Book No.");
                    return;
                }

                ChqBookModel.AccNo = SelectedAccount.AccNo;
                Console.WriteLine($"Error fetching Data: {ChqBookModel.Chqs}");

                await _chqbookService.SaveChqBook(ChqBookModel);

                isError = false;
                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");
                isSaving = false;
                RefreshChqBooks(ChqBookModel.AccNo);
                HideModal();
            }
            catch (Exception ex)
            {
                isError = true;
                NotificationServiceManager.ShowError("Error", $"{ex.Message}");
                isSaving = false;
            }
            finally
            {
                isSaving = false;
            }
        }
        private async Task DeleteChqBook(ItemClickEventArgs e)
        {
            try
            {

                var chqbook = e.Data  as ChqBookDetailViewModel;
                SelectedChqBook = chqbook;
                if (SelectedChqBook != null)
                {

                    int iChqUsed = await _idbHelperService.GetSingleValueAsync<int>("ISNULL(Chqs,0)-ISNULL(ChqsLeft,0)", "VChqBooks", " WHERE ChqBookNo='" + SelectedChqBook.ChqBookNo + "'");

                    if (iChqUsed > 0) {
                        NotificationServiceManager.ShowWarning("Warning", $"Can't Delete. Cheque(s) issued.");
                        return;
                    }   

                    bool deleted = await _chqbookService.DeleteChqBook(SelectedChqBook.ChqBookNo);
                    if (deleted)
                    {
                        NotificationServiceManager.ShowSuccess("Cheque Book Deleted.", $"Cheque Book Deleted Successfully.");
                        if (SelectedAccount != null)
                        {
                            RefreshChqBooks(SelectedAccount.AccNo);
                        }
                        else
                        {
                            RefreshChqBooks("");
                        }
                    }
                    else
                    {
                        NotificationServiceManager.ShowError("Error", $"Can't Delete Cheque Book Is Using");
                    }
                }
            }
            catch (System.Exception ex)
            {
                NotificationServiceManager.ShowError("Error", $"{ex.Message}");
            }
        }
        private void OpenLedger(ItemClickEventArgs e)
        {
            var bankaccno = e.Data as BankAccountInfo;
            SelectedBankAccount = bankaccno;

            var dtTo = DateTime.Today;
            var dtFrom = dtTo.AddDays(-30);

            Navigation.NavigateTo($"/accounts/accountsledger?para_AccNo={SelectedBankAccount.AccNo}&para_DTFrom={dtFrom:yyyy-MM-dd}&para_DTTo={dtTo:yyyy-MM-dd}");
        }

        private bool bPrintCheque = false;
        private ChqListModel? SelectedCheque;
        private async Task ShowPrintCheque(ItemClickEventArgs e)
        {
            SelectedCheque = e.Data as ChqListModel;
            _cachedMenuEvent = e;
            if (SelectedCheque == null)
                return;

            SelectedCheque.chqprintdate = DateTime.Now;

            int iprinted = _idbHelperService.getSingleIntValue ("ChqPrintingDone", "Cheque", " WHERE SNo=" + SelectedCheque.CSNo + "");
            if (iprinted == 1)
            {
                bool printAgain = await NotificationServiceManager.ShowQuestionNotification("Confirmation", "Cheque Already Printed. Do you want to print again?");
                if (!printAgain)
                {
                    bPrintCheque = false;
                    return;
                }
                else 
                {
                    bPrintCheque = true;
                }
            }
            else
            {
                bPrintCheque = true;
            }
        }

        private void HidePrintCheque()
        {
            bPrintCheque = false;
        }

        private bool bShowCancelChqModel = false;
        private ItemClickEventArgs _cachedMenuEvent;
        private void ShowCancelChqModel(ItemClickEventArgs e)
        {
            var chqno = e.Data as ChqListModel;
            SelectedCheque = chqno;
            _cachedMenuEvent = e;
            if (SelectedCheque != null)
            {
                bShowCancelChqModel = true;
                StateHasChanged();
            }
        }

        private void HideCancelChqModel()
        {
            bShowCancelChqModel =false;
            StateHasChanged();
        }

        private async Task MakeChequeCancel()
        {
            await ChangeChequeStatus(1, _cachedMenuEvent);
            bShowCancelChqModel = false;
            StateHasChanged();
        }

        private async Task PrintCheque()
        {
            string rptName = await _idbHelperService.getSingleStringValue("ChqFormat", "VChqLedger", " WHERE SNo=" + SelectedCheque.CSNo + "");
            if (rptName == "") 
            {
                rptName = "rptChq.rpt";
            }
            else
            {
                rptName = rptName.Replace(".rpt", "", StringComparison.OrdinalIgnoreCase) + ".rpt";
            }
            string strcond= "{VChqLedger.SNo}="+ SelectedCheque.CSNo +"";

            //ReportNavigationService.PrintChequeReportAsync(rptName, strcond);
            ReportNavigationService.PrintReportAsync_Old(rptName, strcond);

            await ChangeChequeStatus(2, _cachedMenuEvent);
            
            StateHasChanged();
            
            bShowCancelChqModel = false;
            bPrintCheque = false;

            StateHasChanged();
        }

        private async Task ChangeChequeStatus(int itype, ItemClickEventArgs e)
        {
            {
                try
                {
                    var chqno = e.Data as ChqListModel;
                    SelectedCheque = chqno;
                    if (SelectedCheque != null)
                    {

                        SelectedCheque.UserName = strUserName;
                        SelectedCheque.MachineName = strComputerName;

                        bool bChanged = await _chqbookService.ChangeChqStatus(SelectedCheque, SelectedCheque.ChqNo, SelectedCheque.CSNo, itype, 0);
                        if (bChanged)
                        {
                            if (itype == 0)
                            {
                                NotificationServiceManager.ShowSuccess("Cheque Status.", $"Cheque marked uncancel.");
                            }
                            else if (itype == 0)
                            {
                                NotificationServiceManager.ShowSuccess("Cheque Status.", $"Cheque marked as cancel.");
                            }

                            RefreshChqsData(SelectedCheque.ChqBookNo);
                        }
                        else
                        {
                            NotificationServiceManager.ShowError("Error", $"Can't Delete Cheque Book Is Using");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    NotificationServiceManager.ShowError("Error", $"{ex.Message}");
                }

            }
        }

        private async Task PrintChqBooks()
        {

            string rptName = "ChqBooksSummary.rpt";
            string strcond = "";

            //ReportNavigationService.PrintChequeReportAsync(rptName, strcond);
            ReportNavigationService.PrintReportAsync_Old (rptName, strcond);

            StateHasChanged();
        }

    }
}