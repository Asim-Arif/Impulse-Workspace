using BlazorContextMenu;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts
{
    public partial class ChqBookDetail : ComponentBase, IDisposable
    {
        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<BankAccountInfo> Accounts = new List<BankAccountInfo>();
        private BankAccountInfo? SelectedAccount = null;
        private BankAccountInfo? SelectedBankAccount = null;
        private ChqBookDetailViewModel? SelectedChqBook = null;
        private ChqBookDetailViewModel ChqBookModel = new ChqBookDetailViewModel();
        private ChqListModel? SelectedCheque = null;

        [Inject]
        private IAuditService _auditService { get; set; } = default!;

        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = default!;

        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; } = default!;

        [Inject]
        private IChqBookDetail _chqbookService { get; set; } = default!;

        [Inject]
        private IDBHelperService _idbHelperService { get; set; } = default!;

        [Inject]
        private NavigationManager Navigation { get; set; } = default!;

        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private bool IsLoading = false;
        private bool isSaving = false;

        // Modals visibility
        private bool ShowModal = false;
        private bool ShowODModel = false;
        private bool bShowCancelChqModel = false;
        private bool bPrintCheque = false;
        private ItemClickEventArgs? _cachedMenuEvent;

        // Cheque Books & Cheques data
        private List<ChqBookDetailViewModel> getDatafromdb { get; set; } = new List<ChqBookDetailViewModel>();
        private List<ChqListModel> getDatafromdb_chqsdata { get; set; } = new List<ChqListModel>();

        // Tree Search and Expansion state
        private string BankSearchText = string.Empty;
        private HashSet<string> ExpandedBanks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> ExpandedBranches = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public class BankTreeGroup
        {
            public string BankName { get; set; } = string.Empty;
            public List<BranchTreeGroup> Branches { get; set; } = new List<BranchTreeGroup>();
        }

        public class BranchTreeGroup
        {
            public string BranchName { get; set; } = string.Empty;
            public string BranchKey { get; set; } = string.Empty;
            public List<BankAccountInfo> Accounts { get; set; } = new List<BankAccountInfo>();
        }

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                strComputerName = _auditService.GetClientIpAddress();
                strUserName = _auditService.GetCurrentUserName();

                await RefreshBankAccounts();

                // Auto-select Default Bank Account (or first account if none marked default)
                var defaultAccount = Accounts.FirstOrDefault(a => a.DefaultBank) ?? Accounts.FirstOrDefault();
                if (defaultAccount != null)
                {
                    await SelectBankAccount(defaultAccount);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Initializing ChqBookDetail: {ex.Message}");
                NotificationServiceManager.ShowError("Error Initializing", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void Dispose()
        {
        }

        private async Task RefreshBankAccounts()
        {
            Accounts = await _chqbookService.GetBankAccountsList();

            // Expand all banks and branches by default so user immediately sees accounts
            foreach (var acc in Accounts)
            {
                if (!string.IsNullOrEmpty(acc.Bank))
                {
                    ExpandedBanks.Add(acc.Bank);
                    ExpandedBranches.Add($"{acc.Bank}_{acc.Branch ?? "Main Branch"}");
                }
            }
        }

        private List<BankTreeGroup> GetTreeGroups()
        {
            var query = Accounts.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(BankSearchText))
            {
                string search = BankSearchText.Trim();
                query = query.Where(a =>
                    (a.Bank != null && a.Bank.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (a.Branch != null && a.Branch.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (a.AccTitle != null && a.AccTitle.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (a.AccNo != null && a.AccNo.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (a.BankAccNo != null && a.BankAccNo.Contains(search, StringComparison.OrdinalIgnoreCase))
                );
            }

            var list = new List<BankTreeGroup>();
            foreach (var bankGroup in query.GroupBy(a => a.Bank ?? "Other Banks").OrderBy(g => g.Key))
            {
                var bg = new BankTreeGroup { BankName = bankGroup.Key };
                foreach (var branchGroup in bankGroup.GroupBy(a => a.Branch ?? "Main Branch").OrderBy(g => g.Key))
                {
                    var br = new BranchTreeGroup
                    {
                        BranchName = branchGroup.Key,
                        BranchKey = $"{bankGroup.Key}_{branchGroup.Key}",
                        Accounts = branchGroup.OrderByDescending(a => a.DefaultBank).ThenBy(a => a.AccTitle).ToList()
                    };
                    bg.Branches.Add(br);
                }
                list.Add(bg);
            }
            return list;
        }

        private void ToggleBankExpand(string bank)
        {
            if (ExpandedBanks.Contains(bank))
                ExpandedBanks.Remove(bank);
            else
                ExpandedBanks.Add(bank);
        }

        private void ToggleBranchExpand(string branchKey)
        {
            if (ExpandedBranches.Contains(branchKey))
                ExpandedBranches.Remove(branchKey);
            else
                ExpandedBranches.Add(branchKey);
        }

        private async Task OnBankHeaderClick(BankTreeGroup bankGroup)
        {
            ToggleBankExpand(bankGroup.BankName);

            // If the currently selected account does not belong to this bank, auto-select this bank's first account
            if (SelectedAccount == null || SelectedAccount.Bank != bankGroup.BankName)
            {
                var firstAcc = bankGroup.Branches.SelectMany(b => b.Accounts).FirstOrDefault();
                if (firstAcc != null)
                {
                    await SelectBankAccount(firstAcc);
                }
            }
        }

        private async Task OnBranchHeaderClick(BranchTreeGroup branchGroup)
        {
            ToggleBranchExpand(branchGroup.BranchKey);

            if (SelectedAccount == null || SelectedAccount.Branch != branchGroup.BranchName)
            {
                var firstAcc = branchGroup.Accounts.FirstOrDefault();
                if (firstAcc != null)
                {
                    await SelectBankAccount(firstAcc);
                }
            }
        }

        private async Task SelectBankAccount(BankAccountInfo acc)
        {
            SelectedAccount = acc;

            // Ensure parent bank and branch remain expanded
            if (!string.IsNullOrEmpty(acc.Bank))
            {
                ExpandedBanks.Add(acc.Bank);
                ExpandedBranches.Add($"{acc.Bank}_{acc.Branch ?? "Main Branch"}");
            }

            await RefreshChqBooks(acc.AccNo);

            SelectedChqBook = getDatafromdb.FirstOrDefault();
            if (SelectedChqBook != null)
            {
                await RefreshChqsData(SelectedChqBook.ChqBookNo);
            }
            else
            {
                getDatafromdb_chqsdata.Clear();
            }
            StateHasChanged();
        }

        private async Task RefreshChqBooks(string? BankAccNo)
        {
            try
            {
                getDatafromdb.Clear();
                var getdatafromDb = await _chqbookService.GetBankChqBooks(BankAccNo ?? "");
                getDatafromdb = getdatafromDb.ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Cheque Books: {ex.Message}");
                NotificationServiceManager.ShowError("Error Message", ex.Message);
            }
            finally
            {
                StateHasChanged();
            }
        }

        private async Task OnChequeBookSelected(ChqBookDetailViewModel chqbook)
        {
            SelectedChqBook = chqbook;
            if (SelectedChqBook != null)
            {
                await RefreshChqsData(SelectedChqBook.ChqBookNo);
            }
            else
            {
                getDatafromdb_chqsdata.Clear();
            }
        }

        private async Task RefreshChqsData(long chqbookno)
        {
            try
            {
                getDatafromdb_chqsdata.Clear();
                var data = await _chqbookService.GetChqBookChequesList(chqbookno);
                getDatafromdb_chqsdata = data.ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Cheques: {ex.Message}");
                NotificationServiceManager.ShowError("Error Message", ex.Message);
            }
            finally
            {
                StateHasChanged();
            }
        }

        private string GetDateInputValue(string? dtStr)
        {
            if (!string.IsNullOrEmpty(dtStr) && DateTime.TryParse(dtStr, out var d))
            {
                return d.ToString("yyyy-MM-dd");
            }
            return "";
        }

        private async Task OnClearanceDateChanged(ChqListModel cheque, ChangeEventArgs e)
        {
            string val = e.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(val))
            {
                cheque.ClearanceDT = "";
                bool success = await _chqbookService.UpdateChequeClearanceDate((long)cheque.CSNo, null);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", "Clearance Date cleared.");
                }
            }
            else if (DateTime.TryParse(val, out DateTime newDate))
            {
                cheque.ClearanceDT = newDate.ToString("dd-MMM-yyyy");
                bool success = await _chqbookService.UpdateChequeClearanceDate((long)cheque.CSNo, newDate);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", "Clearance Date updated successfully.");
                }
                else
                {
                    NotificationServiceManager.ShowError("Error", "Failed to update Clearance Date.");
                }
            }
        }

        private void ShowAddModal()
        {
            if (string.IsNullOrEmpty(SelectedAccount?.AccNo))
            {
                NotificationServiceManager.ShowWarning("Warning", "Please Select a Bank Account for the Cheque Book.");
                return;
            }

            ChqBookModel = new ChqBookDetailViewModel
            {
                AccNo = SelectedAccount.AccNo,
                Chqs = 25,
                ManualNo = string.Empty,
                StartingFrom = string.Empty,
                ChqBookDetail = string.Empty
            };
            ShowModal = true;
        }

        private void HideModal()
        {
            ShowModal = false;
            ChqBookModel = new ChqBookDetailViewModel();
        }

        private async Task SaveChqBook()
        {
            if (isSaving)
                return;
            try
            {
                isSaving = true;

                if (SelectedAccount == null || string.IsNullOrEmpty(SelectedAccount.AccNo))
                {
                    NotificationServiceManager.ShowWarning("Validation", "Please select a Bank Account.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(ChqBookModel.StartingFrom))
                {
                    NotificationServiceManager.ShowWarning("Validation", "Please enter Starting From.");
                    return;
                }
                if (ChqBookModel.Chqs <= 0)
                {
                    NotificationServiceManager.ShowWarning("Validation", "Please enter valid number of cheques.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(ChqBookModel.ManualNo))
                {
                    NotificationServiceManager.ShowWarning("Validation", "Please enter Manual Book No.");
                    return;
                }

                ChqBookModel.AccNo = SelectedAccount.AccNo;

                await _chqbookService.SaveChqBook(ChqBookModel);

                NotificationServiceManager.ShowSuccess("Saved", "Cheque Book saved successfully.");
                HideModal();

                await RefreshChqBooks(SelectedAccount.AccNo);
                SelectedChqBook = getDatafromdb.FirstOrDefault(b => b.ManualNo == ChqBookModel.ManualNo) ?? getDatafromdb.FirstOrDefault();
                if (SelectedChqBook != null)
                {
                    await RefreshChqsData(SelectedChqBook.ChqBookNo);
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                isSaving = false;
                StateHasChanged();
            }
        }

        private async Task DeleteChqBook(ItemClickEventArgs e)
        {
            try
            {
                var chqbook = e.Data as ChqBookDetailViewModel ?? SelectedChqBook;
                if (chqbook != null)
                {
                    int iChqUsed = await _idbHelperService.GetSingleValueAsync<int>("ISNULL(Chqs,0)-ISNULL(ChqsLeft,0)", "VChqBooks", " WHERE ChqBookNo='" + chqbook.ChqBookNo + "'");
                    if (iChqUsed > 0)
                    {
                        NotificationServiceManager.ShowWarning("Warning", "Can't Delete. Cheque(s) already issued.");
                        return;
                    }

                    bool deleted = await _chqbookService.DeleteChqBook(chqbook.ChqBookNo);
                    if (deleted)
                    {
                        NotificationServiceManager.ShowSuccess("Deleted", "Cheque Book deleted successfully.");
                        await RefreshChqBooks(SelectedAccount?.AccNo ?? "");
                        SelectedChqBook = getDatafromdb.FirstOrDefault();
                        if (SelectedChqBook != null)
                        {
                            await RefreshChqsData(SelectedChqBook.ChqBookNo);
                        }
                        else
                        {
                            getDatafromdb_chqsdata.Clear();
                        }
                    }
                    else
                    {
                        NotificationServiceManager.ShowError("Error", "Could not delete Cheque Book.");
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        private async Task MakeBankAccountDefault(ItemClickEventArgs e)
        {
            var bankaccount = e.Data as BankAccountInfo ?? SelectedAccount;
            if (bankaccount != null)
            {
                bankaccount.DefaultBank = true;
                bool saved = await _chqbookService.SaveOverDraftData(bankaccount, bankaccount.AccNo, true);
                if (saved)
                {
                    NotificationServiceManager.ShowSuccess("Success", $"[{bankaccount.AccNo}] set as default bank account.");
                    await RefreshBankAccounts();
                    StateHasChanged();
                }
            }
        }

        private void ShowOverDraftModel(ItemClickEventArgs e)
        {
            var bankaccount = e.Data as BankAccountInfo ?? SelectedAccount;
            SelectedBankAccount = bankaccount;
            if (SelectedBankAccount != null)
            {
                ShowODModel = true;
            }
        }

        private void HideODModel()
        {
            ShowODModel = false;
        }

        private async Task SaveOverDraftData()
        {
            if (isSaving || SelectedBankAccount == null)
                return;
            try
            {
                isSaving = true;
                bool saved = await _chqbookService.SaveOverDraftData(SelectedBankAccount, SelectedBankAccount.AccNo, false);
                if (saved)
                {
                    NotificationServiceManager.ShowSuccess("Saved", "Overdraft settings updated successfully.");
                    HideODModel();
                    await RefreshBankAccounts();
                }
                else
                {
                    NotificationServiceManager.ShowError("Error", "Could not save overdraft settings.");
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                isSaving = false;
                StateHasChanged();
            }
        }

        private void OpenLedger(ItemClickEventArgs e)
        {
            var bankacc = e.Data as BankAccountInfo ?? SelectedAccount;
            if (bankacc != null)
            {
                var dtTo = DateTime.Today;
                var dtFrom = dtTo.AddDays(-30);
                Navigation.NavigateTo($"/accounts/accountsledger?para_AccNo={bankacc.AccNo}&para_DTFrom={dtFrom:yyyy-MM-dd}&para_DTTo={dtTo:yyyy-MM-dd}");
            }
        }

        private async Task PrintChqBooks(ItemClickEventArgs? e = null)
        {
            var request = new ReportRequest
            {
                ReportName = "ChqBooksSummary.rpt",
                SelectionFormula = ""
            };
            await ReportNavigationService.PrintReportAsync(request);
        }

        private async Task PrintChqBookDetail(ItemClickEventArgs e)
        {
            var chqbook = e.Data as ChqBookDetailViewModel ?? SelectedChqBook;
            if (chqbook != null)
            {
                var request = new ReportRequest
                {
                    ReportName = "ChqBook.rpt",
                    SelectionFormula = $"{{VChqBookLedger.ChqbookNo}}={chqbook.ChqBookNo}"
                };
                await ReportNavigationService.PrintReportAsync(request);
            }
        }

        private async Task PrintSelectedChqBookDetail()
        {
            if (SelectedChqBook != null)
            {
                var request = new ReportRequest
                {
                    ReportName = "ChqBook.rpt",
                    SelectionFormula = $"{{VChqBookLedger.ChqbookNo}}={SelectedChqBook.ChqBookNo}"
                };
                await ReportNavigationService.PrintReportAsync(request);
            }
        }

        private async Task ShowPrintCheque(ItemClickEventArgs e)
        {
            SelectedCheque = e.Data as ChqListModel;
            _cachedMenuEvent = e;
            if (SelectedCheque == null)
                return;

            SelectedCheque.chqprintdate = DateTime.Now;

            int iprinted = _idbHelperService.getSingleIntValue("ChqPrintingDone", "Cheque", " WHERE SNo=" + SelectedCheque.CSNo + "");
            if (iprinted == 1)
            {
                bool printAgain = await NotificationServiceManager.ShowQuestionNotification("Confirmation", "Cheque Already Printed. Do you want to print again?");
                if (!printAgain)
                {
                    bPrintCheque = false;
                    return;
                }
            }
            bPrintCheque = true;
        }

        private void HidePrintCheque()
        {
            bPrintCheque = false;
        }

        private async Task PrintCheque()
        {
            if (SelectedCheque == null)
                return;

            string rptName = await _idbHelperService.getSingleStringValue("ChqFormat", "VChqLedger", " WHERE SNo=" + SelectedCheque.CSNo + "");
            if (string.IsNullOrWhiteSpace(rptName))
            {
                rptName = "rptChq.rpt";
            }
            else
            {
                rptName = rptName.Replace(".rpt", "", StringComparison.OrdinalIgnoreCase) + ".rpt";
            }

            var request = new ReportRequest
            {
                ReportName = rptName,
                SelectionFormula = "{VChqLedger.SNo}=" + SelectedCheque.CSNo
            };
            await ReportNavigationService.PrintReportAsync(request);

            if (_cachedMenuEvent != null)
            {
                await ChangeChequeStatus(2, _cachedMenuEvent);
            }

            bPrintCheque = false;
            StateHasChanged();
        }

        private void ShowCancelChqModel(ItemClickEventArgs e)
        {
            SelectedCheque = e.Data as ChqListModel;
            _cachedMenuEvent = e;
            if (SelectedCheque != null)
            {
                bShowCancelChqModel = true;
                StateHasChanged();
            }
        }

        private void HideCancelChqModel()
        {
            bShowCancelChqModel = false;
            StateHasChanged();
        }

        private async Task MakeChequeCancel()
        {
            if (_cachedMenuEvent != null)
            {
                await ChangeChequeStatus(1, _cachedMenuEvent);
            }
            bShowCancelChqModel = false;
            StateHasChanged();
        }

        private async Task ChangeChequeStatus(int itype, ItemClickEventArgs e)
        {
            try
            {
                var chqno = e.Data as ChqListModel ?? SelectedCheque;
                if (chqno != null)
                {
                    chqno.UserName = strUserName;
                    chqno.MachineName = strComputerName;

                    bool bChanged = await _chqbookService.ChangeChqStatus(chqno, chqno.ChqNo, chqno.CSNo, itype, 0);
                    if (bChanged)
                    {
                        if (itype == 0)
                        {
                            NotificationServiceManager.ShowSuccess("Cheque Status", "Cheque marked as uncanceled.");
                        }
                        else if (itype == 1)
                        {
                            NotificationServiceManager.ShowSuccess("Cheque Status", "Cheque marked as canceled.");
                        }

                        if (SelectedChqBook != null)
                        {
                            await RefreshChqsData(SelectedChqBook.ChqBookNo);
                        }
                    }
                    else
                    {
                        NotificationServiceManager.ShowError("Error", "Could not change cheque status.");
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                Navigation.NavigateTo(ReturnUrl);
            }
            else
            {
                Navigation.NavigateTo("/");
            }
        }
    }
}