using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using Impulse.Services.Accounts;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts.MakerLongTermLoan
{
    public partial class MakerLongTermLoan : ComponentBase
    {
        [Inject] private IMakerLongTermLoanService MakerLoanService { get; set; } = default!;
        [Inject] private IVoucherService VoucherService { get; set; } = default!;
        [Inject] private IDBHelperService DBHelperService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<MakerAccountModel> Makers { get; set; } = new();
        private List<MakerMasterPoModel> MasterPOs { get; set; } = new();

        private MakerAccountModel? _selectedMaker;
        private MakerAccountModel? SelectedMaker
        {
            get => _selectedMaker;
            set
            {
                if (_selectedMaker != value)
                {
                    _selectedMaker = value;
                    _ = HandleMakerSelectionAsync(value);
                }
            }
        }

        private MakerMasterPoModel? _selectedMasterPO;
        private MakerMasterPoModel? SelectedMasterPO
        {
            get => _selectedMasterPO;
            set
            {
                if (_selectedMasterPO != value)
                {
                    _selectedMasterPO = value;
                    _ = HandleMasterPOSelectionAsync(value);
                }
            }
        }

        private decimal PreviousBalance { get; set; } = 0m;
        private DateTime IssueDate { get; set; } = DateTime.Today;
        private decimal Amount { get; set; } = 0m;
        private decimal DAmount { get; set; } = 0m;
        private string Description { get; set; } = string.Empty;

        // ── Posting State ──
        private MakerLoanPaymentMode PaymentMode { get; set; } = MakerLoanPaymentMode.Cash;

        private ChartOfAccountsModel? SelectedCashAccount { get; set; }
        private ChartOfAccountsModel? SelectedOtherAccount { get; set; }

        private BankAccountInfo? _selectedBankAccount;
        private BankAccountInfo? SelectedBankAccount
        {
            get => _selectedBankAccount;
            set
            {
                if (_selectedBankAccount != value)
                {
                    _selectedBankAccount = value;
                    _ = HandleBankAccountSelectionAsync(value);
                }
            }
        }

        private string ChequeNo { get; set; } = string.Empty;
        private string ChequeType { get; set; } = "Open";
        private DateTime ChequeDate { get; set; } = DateTime.Today;
        private bool DontPostVoucher { get; set; } = false;

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        private List<ChartOfAccountsModel> CashAccounts { get; set; } = new();
        private List<ChartOfAccountsModel> AllAccounts { get; set; } = new();
        private List<BankAccountInfo> BankAccounts { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Makers = await MakerLoanService.GetMakerAccountsAsync();
                await LoadAccountsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to initialize form: {ex.Message}"
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task LoadAccountsAsync()
        {
            try
            {
                AllAccounts = (await VoucherService.GetTransactionalAccounts()) ?? new List<ChartOfAccountsModel>();
                BankAccounts = (await VoucherService.GetBankAccounts()) ?? new List<BankAccountInfo>();

                string cashInHandParent = await DBHelperService.getSingleStringValue(
                    "DataValue", "GeneralData", "WHERE DataName='CashInHandParent'");

                if (!string.IsNullOrEmpty(cashInHandParent))
                {
                    CashAccounts = (await VoucherService.GetTransactionalAccounts($"WHERE SubAccOf='{cashInHandParent}'"))
                                   ?? new List<ChartOfAccountsModel>();
                }
                else
                {
                    CashAccounts = AllAccounts.Where(a => a.AccTitle != null && a.AccTitle.Contains("Cash", StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!CashAccounts.Any()) CashAccounts = new List<ChartOfAccountsModel>(AllAccounts);
                if (CashAccounts.Any())
                {
                    SelectedCashAccount = CashAccounts.FirstOrDefault(a => a.AccTitle != null && a.AccTitle.Contains("Petty", StringComparison.OrdinalIgnoreCase)) ?? CashAccounts.First();
                }
            }
            catch { }
        }

        private async Task HandleMakerSelectionAsync(MakerAccountModel? maker)
        {
            if (maker != null && !string.IsNullOrWhiteSpace(maker.AccNo))
            {
                PreviousBalance = await MakerLoanService.GetMakerBalanceAsync(maker.AccNo);
                MasterPOs = await MakerLoanService.GetMakerMasterPOsAsync(maker.VendID, maker.AccNo);
            }
            else
            {
                PreviousBalance = 0m;
                MasterPOs = new List<MakerMasterPoModel>();
                SelectedMasterPO = null;
            }
        }

        private async Task HandleMasterPOSelectionAsync(MakerMasterPoModel? po)
        {
            if (po != null && !string.IsNullOrWhiteSpace(po.MasterPONo))
            {
                bool isVendor = SelectedMaker != null && SelectedMaker.AccNo != null && SelectedMaker.AccNo.StartsWith("24-008");
                var details = await MakerLoanService.GetMasterPoDetailsAsync(po.MasterPONo, isVendor);
                if (details != null)
                {
                    po.POValue = details.POValue;
                    po.LTTaken = details.LTTaken;
                    po.STTaken = details.STTaken;
                    StateHasChanged();
                }
            }
        }

        private async Task HandleBankAccountSelectionAsync(BankAccountInfo? bank)
        {
            if (bank != null && !string.IsNullOrWhiteSpace(bank.AccNo))
            {
                try
                {
                    var chequeInfo = await VoucherService.GetNewChequeNumbersByBankAccount(bank.AccNo);
                    if (chequeInfo != null)
                    {
                        ChequeNo = chequeInfo.ChqNo ?? string.Empty;
                    }
                }
                catch { }
            }
            else
            {
                ChequeNo = string.Empty;
            }
        }

        private void SetPaymentMode(MakerLoanPaymentMode mode)
        {
            PaymentMode = mode;
            if (mode != MakerLoanPaymentMode.Bank)
            {
                ChequeNo = string.Empty;
                _selectedBankAccount = null;
            }
        }

        private async Task<IEnumerable<MakerAccountModel>> SearchMakersAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(Makers);
            return await Task.FromResult(Makers.Where(m =>
                (m.MakerName != null && m.MakerName.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.MakerNo != null && m.MakerNo.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                (m.AccNo != null && m.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<IEnumerable<MakerMasterPoModel>> SearchMasterPOsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(MasterPOs);
            return await Task.FromResult(MasterPOs.Where(p => p.MasterPONo != null && p.MasterPONo.Contains(text, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchCashAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(CashAccounts);
            return await Task.FromResult(CashAccounts.Where(a => (a.AccTitle != null && a.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                 (a.AccNo != null && a.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchAllAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(AllAccounts);
            return await Task.FromResult(AllAccounts.Where(a => (a.AccTitle != null && a.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                (a.AccNo != null && a.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<IEnumerable<BankAccountInfo>> SearchBankAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(BankAccounts);
            return await Task.FromResult(BankAccounts.Where(b => (b.Bank != null && b.Bank.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                 (b.AccTitle != null && b.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<bool> SaveAsync()
        {
            if (SelectedMaker == null || string.IsNullOrWhiteSpace(SelectedMaker.AccNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select a valid Maker account."
                });
                return false;
            }

            if (Amount <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter a valid loan amount greater than zero."
                });
                return false;
            }

            string paymentAccNo = string.Empty;
            if (PaymentMode == MakerLoanPaymentMode.Cash)
            {
                paymentAccNo = SelectedCashAccount?.AccNo ?? string.Empty;
            }
            else if (PaymentMode == MakerLoanPaymentMode.OtherAccount)
            {
                paymentAccNo = SelectedOtherAccount?.AccNo ?? string.Empty;
            }
            else if (PaymentMode == MakerLoanPaymentMode.Bank)
            {
                paymentAccNo = SelectedBankAccount?.AccNo ?? string.Empty;
                if (string.IsNullOrWhiteSpace(ChequeNo) && !DontPostVoucher)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Validation Warning",
                        Detail = "Please enter a valid Cheque Number for Bank Payment."
                    });
                    return false;
                }
            }

            if (string.IsNullOrWhiteSpace(paymentAccNo) && !DontPostVoucher)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select a valid payment account."
                });
                return false;
            }

            IsSaving = true;
            StateHasChanged();

            try
            {
                var dto = new MakerLongTermLoanDto
                {
                    AccNo = SelectedMaker.AccNo,
                    MakerNo = SelectedMaker.MakerNo,
                    MakerName = SelectedMaker.MakerName,
                    IssueDate = IssueDate,
                    Amount = Amount,
                    DAmount = DAmount,
                    Description = Description,
                    MasterPONo = SelectedMasterPO?.MasterPONo,
                    POValue = SelectedMasterPO?.POValue ?? 0m,
                    LTTaken = SelectedMasterPO?.LTTaken ?? 0m,
                    STTaken = SelectedMasterPO?.STTaken ?? 0m,
                    PaymentMode = PaymentMode,
                    PaymentAccNo = paymentAccNo,
                    BankAccNo = SelectedBankAccount?.AccNo ?? string.Empty,
                    BankId = SelectedBankAccount?.AccNo ?? string.Empty,
                    ChequeNo = ChequeNo,
                    ChequeType = ChequeType,
                    ChequeDate = ChequeDate,

                    DontPostVoucher = DontPostVoucher
                };

                string vchrNo = await MakerLoanService.SaveLoanAsync(dto);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved Successfully",
                    Detail = string.IsNullOrWhiteSpace(vchrNo)
                        ? "Maker Long Term Loan recorded successfully."
                        : $"Voucher {vchrNo} and Maker Loan recorded successfully.",
                    Duration = 4000
                });

                return true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
                return false;
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private async Task SaveAndNewAsync()
        {
            if (await SaveAsync())
            {
                Amount = 0m;
                DAmount = 0m;
                Description = string.Empty;
                SelectedMasterPO = null;
                ChequeNo = string.Empty;
                if (SelectedMaker != null)
                {
                    PreviousBalance = await MakerLoanService.GetMakerBalanceAsync(SelectedMaker.AccNo);
                }
            }
        }

        private async Task SaveAndCloseAsync()
        {
            if (await SaveAsync())
            {
                NavigateBack();
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/financial");
            }
        }
    }
}
