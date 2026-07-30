using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.AdvanceLong
{
    public partial class AdvanceLong : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IAdvanceLongService     _advanceLongService  { get; set; } = null!;
        [Inject] private IEmployeeService         _employeeService     { get; set; } = null!;
        [Inject] private IVoucherService          _voucherService      { get; set; } = null!;
        [Inject] private IDBHelperService         _dbHelperService     { get; set; } = null!;
        [Inject] private INotificationService    _notificationService { get; set; } = null!;
        [Inject] private IReportNavigationService _reportNavService   { get; set; } = null!;
        [Inject] private NavigationManager       _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl    { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors  { get; set; } = false;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading    { get; set; } = true;
        private bool IsSaving     { get; set; } = false;
        private bool IsContractor { get; set; } = false;

        // ── Employee / Department ─────────────────────────────────────────────
        private List<DepartmentListItemModel>  AllDepartments { get; set; } = new();
        private List<EmployeeListItemModel>    AllEmployees   { get; set; } = new();
        private DepartmentListItemModel?       SelectedDepartment { get; set; }
        private EmployeeListItemModel?         SelectedEmployee   { get; set; }
        private AdvanceLongSummaryModel?       EmployeeSummary    { get; set; }

        // ── Advance Fields ────────────────────────────────────────────────────
        private DateTime IssueDate        { get; set; } = DateTime.Today;
        private DateTime DeductionStartDT { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        private decimal  Amount           { get; set; } = 0;
        private decimal  MonthlyDeduction { get; set; } = 0;
        private string   Description      { get; set; } = "Long Term Loan";

        // ── Payment Mode ──────────────────────────────────────────────────────
        private AdvanceLongPaymentMode PaymentMode { get; set; } = AdvanceLongPaymentMode.Cash;

        // ── Account Lookups ───────────────────────────────────────────────────
        private List<ChartOfAccountsModel> CashAccounts  { get; set; } = new();
        private List<ChartOfAccountsModel> AllAccounts   { get; set; } = new();
        private List<BankAccountInfo>       BankAccounts  { get; set; } = new();

        private ChartOfAccountsModel? SelectedCashAccount  { get; set; }
        private ChartOfAccountsModel? SelectedOtherAccount { get; set; }
        private BankAccountInfo?      SelectedBankAccount  { get; set; }

        // ── Cheque Fields (BPV) ───────────────────────────────────────────────
        private string   ChequeNo   { get; set; } = string.Empty;
        private string   ChequeType { get; set; } = "Open";
        private DateTime ChequeDate { get; set; } = DateTime.Today;
        private bool     IsChequeLoading { get; set; } = false;

        // ── Voucher Number Preview ────────────────────────────────────────────
        private string PreviewVchrNo   { get; set; } = string.Empty;
        private bool   IsVchrNoLoading { get; set; } = false;

        // ── EmpLongTermAccNo (validated on init) ──────────────────────────────
        private string EmpLongTermAccNo { get; set; } = string.Empty;

        // ── Computed Properties ───────────────────────────────────────────────
        private string CurrentVchrType => PaymentMode switch
        {
            AdvanceLongPaymentMode.Cash         => "CPV",
            AdvanceLongPaymentMode.OtherAccount => "JV",
            AdvanceLongPaymentMode.Bank         => "BPV",
            _                                   => "CPV"
        };

        private string VchrBadgeClass => PaymentMode switch
        {
            AdvanceLongPaymentMode.Cash         => "badge bg-success",
            AdvanceLongPaymentMode.OtherAccount => "badge bg-secondary", // purple-ish secondary
            AdvanceLongPaymentMode.Bank         => "badge bg-primary",
            _                                   => "badge bg-success"
        };

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsContractor = contractors;
            IsLoading    = true;

            try
            {
                // Load departments
                AllDepartments = await _employeeService.GetDepartmentsAsync(false);

                // Load employees (all depts initially)
                AllEmployees = await _employeeService.GetEmployeesAsync("0", false, IsContractor);

                // Load accounts
                AllAccounts  = await _voucherService.GetTransactionalAccounts();
                BankAccounts = await _voucherService.GetBankAccounts();

                // Filter cash accounts (sub-accounts of CashInHandParent from GeneralData)
                string cashInHandParent = await _dbHelperService.getSingleStringValue(
                    "DataValue", "GeneralData", "WHERE DataName='CashInHandParent'");

                if (!string.IsNullOrEmpty(cashInHandParent))
                {
                    CashAccounts = await _voucherService.GetTransactionalAccounts(
                        $"WHERE SubAccOf='{cashInHandParent}'");
                }
                else
                {
                    CashAccounts = new List<ChartOfAccountsModel>(AllAccounts);
                }

                // Validate EmpLongTermAccNo is configured
                EmpLongTermAccNo = await _advanceLongService.GetEmpLongTermAccNoAsync();
                if (string.IsNullOrWhiteSpace(EmpLongTermAccNo))
                {
                    _notificationService.ShowWarning(
                        "Configuration Missing",
                        "EmpLongTermAccNo is not set in General Data. Saving will not work until an administrator configures it.");
                }

                // Set default cash account (Petty Cash if exists)
                if (CashAccounts.Any())
                {
                    SelectedCashAccount = CashAccounts.FirstOrDefault(a => a.AccTitle.Contains("Petty")) 
                                          ?? CashAccounts.First();
                }

                // Fetch initial voucher number preview
                await RefreshVchrNoPreviewAsync();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Load Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Event Handlers ────────────────────────────────────────────────────

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            SelectedEmployee   = null;
            EmployeeSummary    = null;

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            if (emp != null)
            {
                EmployeeSummary = await _advanceLongService.GetEmployeeSummaryAsync(emp.EmpID);
            }
            else
            {
                EmployeeSummary = null;
            }
        }

        private async Task OnIssueDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                IssueDate        = d;
                // Default deduction start to 1st of same month as issue date
                DeductionStartDT = new DateTime(d.Year, d.Month, 1);
                await RefreshVchrNoPreviewAsync();
            }
        }

        private void OnDeductionStartChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString() + "-01", out DateTime d))
            {
                DeductionStartDT = new DateTime(d.Year, d.Month, 1);
            }
        }

        private async Task SetPaymentMode(AdvanceLongPaymentMode mode)
        {
            PaymentMode = mode;
            // Clear cheque fields when switching away from Bank
            if (mode != AdvanceLongPaymentMode.Bank)
            {
                ChequeNo = string.Empty;
                ChequeType = "Open";
                SelectedBankAccount = null;
            }
            await RefreshVchrNoPreviewAsync();
        }

        private void OnCashAccountChanged(ChartOfAccountsModel? acc)
        {
            SelectedCashAccount = acc;
        }

        private void OnOtherAccountChanged(ChartOfAccountsModel? acc)
        {
            SelectedOtherAccount = acc;
        }

        private async Task OnBankAccountChanged(BankAccountInfo? bank)
        {
            SelectedBankAccount = bank;
            if (bank != null)
            {
                IsChequeLoading = true;
                try
                {
                    var chequeInfo = await _voucherService.GetNewChequeNumbersByBankAccount(bank.AccNo);
                    if (chequeInfo != null)
                    {
                        ChequeNo  = chequeInfo.ChqNo ?? string.Empty;
                    }
                }
                finally
                {
                    IsChequeLoading = false;
                }
            }
            else
            {
                ChequeNo = string.Empty;
            }
        }

        // ── Voucher Preview ───────────────────────────────────────────────────

        private async Task RefreshVchrNoPreviewAsync()
        {
            IsVchrNoLoading = true;
            try
            {
                PreviewVchrNo = await _advanceLongService.GetPreviewVchrNoAsync(IssueDate, PaymentMode);
            }
            catch
            {
                PreviewVchrNo = "—";
            }
            finally
            {
                IsVchrNoLoading = false;
            }
        }

        // ── Search Methods (BlazoredTypeahead) ────────────────────────────────

        private Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string search)
        {
            var results = AllDepartments
                .Where(d => string.IsNullOrWhiteSpace(search) ||
                            d.Name.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string search)
        {
            var results = AllEmployees
                .Where(e => string.IsNullOrWhiteSpace(search) ||
                            e.EmpID.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            e.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.FName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (e.Designation ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<ChartOfAccountsModel>> SearchCashAccounts(string search)
        {
            var results = CashAccounts
                .Where(a => string.IsNullOrWhiteSpace(search) ||
                            a.AccNo.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            a.AccTitle.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<ChartOfAccountsModel>> SearchAllAccounts(string search)
        {
            var results = AllAccounts
                .Where(a => string.IsNullOrWhiteSpace(search) ||
                            a.AccNo.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            a.AccTitle.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        private Task<IEnumerable<BankAccountInfo>> SearchBankAccounts(string search)
        {
            var results = BankAccounts
                .Where(b => string.IsNullOrWhiteSpace(search) ||
                            b.Bank.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            b.AccTitle.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            b.Branch.Contains(search, StringComparison.OrdinalIgnoreCase))
                .Take(100);
            return Task.FromResult(results);
        }

        // ── Save ──────────────────────────────────────────────────────────────

        private async Task SaveAsync()
        {
            if (IsSaving) return;

            // ── Validation ────────────────────────────────────────────────────
            if (SelectedEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select an employee.");
                return;
            }
            if (Amount <= 0)
            {
                _notificationService.ShowWarning("Validation", "Loan Amount must be greater than zero.");
                return;
            }
            if (MonthlyDeduction > Amount)
            {
                _notificationService.ShowWarning("Validation", "Monthly Deduction cannot exceed the Loan Amount.");
                return;
            }
            if (string.IsNullOrWhiteSpace(EmpLongTermAccNo))
            {
                _notificationService.ShowError("Configuration Error",
                    "EmpLongTermAccNo is not configured in General Data. Cannot save.");
                return;
            }

            string paymentAccNo = string.Empty;
            int    bankId       = 0;

            switch (PaymentMode)
            {
                case AdvanceLongPaymentMode.Cash:
                    if (SelectedCashAccount == null)
                    {
                        _notificationService.ShowWarning("Validation", "Please select a Cash In Hand account.");
                        return;
                    }
                    paymentAccNo = SelectedCashAccount.AccNo;
                    break;

                case AdvanceLongPaymentMode.OtherAccount:
                    if (SelectedOtherAccount == null)
                    {
                        _notificationService.ShowWarning("Validation", "Please select an Account.");
                        return;
                    }
                    paymentAccNo = SelectedOtherAccount.AccNo;
                    break;

                case AdvanceLongPaymentMode.Bank:
                    if (SelectedBankAccount == null)
                    {
                        _notificationService.ShowWarning("Validation", "Please select a Bank Account.");
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(ChequeNo))
                    {
                        _notificationService.ShowWarning("Validation", "Please enter the Cheque Number.");
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(ChequeType))
                    {
                        _notificationService.ShowWarning("Validation", "Please select the Cheque Type.");
                        return;
                    }
                    paymentAccNo = SelectedBankAccount.AccNo;
                    bankId       = SelectedBankAccount.BankID;
                    break;
            }

            // ── Build DTO ─────────────────────────────────────────────────────
            var dto = new AdvanceLongDto
            {
                EmpID            = SelectedEmployee.EmpID,
                IssueDate        = IssueDate,
                DeductionStartDT = DeductionStartDT,
                Amount           = Amount,
                DAmount          = MonthlyDeduction,
                Description      = string.IsNullOrWhiteSpace(Description) ? "Long Term Loan" : Description,
                PaymentMode      = PaymentMode,
                PaymentAccNo     = paymentAccNo,
                IsContractor     = IsContractor,
                BankID           = bankId,
                ChequeNo         = ChequeNo,
                ChequeType       = ChequeType,
                ChequeDate       = ChequeDate
            };

            // ── Save ──────────────────────────────────────────────────────────
            IsSaving = true;
            try
            {
                string postedVchrNo = await _advanceLongService.SaveAsync(dto);

                _notificationService.ShowSuccess(
                    "Saved Successfully",
                    $"Long Term Advance recorded. Voucher: {postedVchrNo}");

                // Print voucher
                await _reportNavService.PrintVoucher(postedVchrNo);

                // Reset form
                ResetForm();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ── Reset Form ────────────────────────────────────────────────────────

        private void ResetForm()
        {
            SelectedEmployee   = null;
            EmployeeSummary    = null;
            Amount             = 0;
            MonthlyDeduction   = 0;
            Description        = "Long Term Loan";
            IssueDate          = DateTime.Today;
            DeductionStartDT   = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            SelectedCashAccount  = null;
            SelectedOtherAccount = null;
            SelectedBankAccount  = null;
            ChequeNo   = string.Empty;
            ChequeType = "Open";
            ChequeDate = DateTime.Today;
            PaymentMode = AdvanceLongPaymentMode.Cash;
            _ = RefreshVchrNoPreviewAsync();
        }

        // ── Navigation ────────────────────────────────────────────────────────

        private void NavigateBack()
        {
            if (!string.IsNullOrEmpty(returnUrl))
                _navManager.NavigateTo(returnUrl);
            else
                _navManager.NavigateTo("/payroll");
        }
    }
}
