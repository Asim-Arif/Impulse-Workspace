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

namespace Impulse.Pages.Payroll.ClearLongTermLoan
{
    public partial class ClearLongTermLoan : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IClearAdvanceService    _clearAdvanceService { get; set; } = null!;
        [Inject] private IEmployeeService        _employeeService     { get; set; } = null!;
        [Inject] private IVoucherService         _voucherService      { get; set; } = null!;
        [Inject] private IDBHelperService        _dbHelperService     { get; set; } = null!;
        [Inject] private INotificationService    _notificationService { get; set; } = null!;
        [Inject] private IReportNavigationService _reportNavService   { get; set; } = null!;
        [Inject] private NavigationManager       _navManager          { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl { get; set; }

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading { get; set; } = true;
        private bool IsSaving  { get; set; } = false;

        // ── Employee / Department ─────────────────────────────────────────────
        private List<DepartmentListItemModel>  AllDepartments { get; set; } = new();
        private List<EmployeeListItemModel>    AllEmployees   { get; set; } = new();
        private DepartmentListItemModel?       SelectedDepartment { get; set; }
        private EmployeeListItemModel?         SelectedEmployee   { get; set; }
        private ClearAdvanceSummaryModel?      EmployeeSummary    { get; set; }
        private string                         EmployeePicBase64  { get; set; } = string.Empty;

        // ── Clearance Fields ──────────────────────────────────────────────────
        private DateTime ClearDate      { get; set; } = DateTime.Today;
        private decimal  AmountReceived { get; set; } = 0;
        private decimal  NewDeduction   { get; set; } = 0;
        private string   Description    { get; set; } = string.Empty;

        // ── Receipt Mode ──────────────────────────────────────────────────────
        private ClearAdvanceReceiptMode ReceiptMode     { get; set; } = ClearAdvanceReceiptMode.Cash;
        private bool                    DontPostVoucher { get; set; } = false;

        // ── Account Lookups ───────────────────────────────────────────────────
        private List<ChartOfAccountsModel> CashAccounts { get; set; } = new();
        private List<ChartOfAccountsModel> AllAccounts  { get; set; } = new();

        private ChartOfAccountsModel? SelectedCashAccount  { get; set; }
        private ChartOfAccountsModel? SelectedOtherAccount { get; set; }

        // ── EmpLongTermAccNo (validated on init) ──────────────────────────────
        private string EmpLongTermAccNo { get; set; } = string.Empty;

        // ── Computed Properties ───────────────────────────────────────────────
        private string CurrentVchrType => DontPostVoucher
            ? "NO VCHR"
            : ReceiptMode == ClearAdvanceReceiptMode.Cash ? "CRV" : "JV";

        private string VchrBadgeClass => DontPostVoucher
            ? "badge bg-dark"
            : ReceiptMode == ClearAdvanceReceiptMode.Cash ? "badge bg-success" : "badge bg-secondary";

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;

            try
            {
                // Load departments
                AllDepartments = await _employeeService.GetDepartmentsAsync(false);

                // Load employees (all depts initially)
                AllEmployees = await _employeeService.GetEmployeesAsync("0", false, false);

                // Load accounts
                AllAccounts = await _voucherService.GetTransactionalAccounts();

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
                EmpLongTermAccNo = await _clearAdvanceService.GetEmpLongTermAccNoAsync();
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
            EmployeePicBase64  = string.Empty;

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, false);
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee  = emp;
            EmployeePicBase64 = string.Empty;

            if (emp != null)
            {
                EmployeeSummary = await _clearAdvanceService.GetSummaryAsync(emp.EmpID);

                var picBytes = await _employeeService.GetEmployeePictureAsync(emp.EmpID);
                if (picBytes != null && picBytes.Length > 0)
                {
                    EmployeePicBase64 = Convert.ToBase64String(picBytes);
                }
            }
            else
            {
                EmployeeSummary = null;
            }
        }

        private void OnClearDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                ClearDate = d;
            }
        }

        private void SetReceiptMode(ClearAdvanceReceiptMode mode)
        {
            ReceiptMode = mode;
        }

        private void OnCashAccountChanged(ChartOfAccountsModel? acc)
        {
            SelectedCashAccount = acc;
        }

        private void OnOtherAccountChanged(ChartOfAccountsModel? acc)
        {
            SelectedOtherAccount = acc;
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

        // ── Save ──────────────────────────────────────────────────────────────

        private Task SaveAndNewAsync()   => SaveAsync(closeAfterSave: false);
        private Task SaveAndCloseAsync() => SaveAsync(closeAfterSave: true);

        private async Task SaveAsync(bool closeAfterSave)
        {
            if (IsSaving) return;

            // ── Validation (order mirrors the legacy frmClearAmt) ─────────────
            if (SelectedEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select an employee.");
                return;
            }
            if (EmployeeSummary == null || EmployeeSummary.OutstandingBalance == 0)
            {
                _notificationService.ShowWarning("Validation", "No Amount To Receive!");
                return;
            }
            if (AmountReceived <= 0)
            {
                _notificationService.ShowWarning("Validation", "Receiving amount could not be zero.");
                return;
            }
            if (AmountReceived > EmployeeSummary.OutstandingBalance)
            {
                _notificationService.ShowWarning("Validation",
                    "Receive amount could not be greater than the total outstanding amount.");
                return;
            }

            string receiptAccNo = string.Empty;

            if (!DontPostVoucher)
            {
                if (ReceiptMode == ClearAdvanceReceiptMode.Cash)
                {
                    if (SelectedCashAccount == null)
                    {
                        _notificationService.ShowWarning("Validation", "Please select a Cash In Hand account.");
                        return;
                    }
                    receiptAccNo = SelectedCashAccount.AccNo;
                }
                else
                {
                    if (SelectedOtherAccount == null)
                    {
                        _notificationService.ShowWarning("Validation", "Please select an Account.");
                        return;
                    }
                    receiptAccNo = SelectedOtherAccount.AccNo;
                }

                if (string.IsNullOrWhiteSpace(EmpLongTermAccNo))
                {
                    _notificationService.ShowError("Configuration Error",
                        "EmpLongTermAccNo is not configured in General Data. Cannot save.");
                    return;
                }
            }

            // ── Build DTO ─────────────────────────────────────────────────────
            var dto = new ClearAdvanceDto
            {
                EmpID                = SelectedEmployee.EmpID,
                EmpDisplayName       = BuildEmployeeDisplayName(SelectedEmployee),
                ClearDate            = ClearDate,
                AmountReceived       = AmountReceived,
                NewDeduction         = NewDeduction,
                Description          = Description ?? string.Empty,
                DontPostVoucher      = DontPostVoucher,
                ReceiptMode          = ReceiptMode,
                ReceiptAccNo         = receiptAccNo,
                LatestAdvanceEntryID = EmployeeSummary.LatestAdvanceEntryID
            };

            // ── Save ──────────────────────────────────────────────────────────
            IsSaving = true;
            try
            {
                string postedVchrNo = await _clearAdvanceService.SaveAsync(dto);

                _notificationService.ShowSuccess(
                    "Saved Successfully",
                    string.IsNullOrEmpty(postedVchrNo)
                        ? "Long Term Loan clearance recorded (no voucher posted)."
                        : $"Long Term Loan clearance recorded. Voucher: {postedVchrNo}");

                // Print voucher (only when one was posted)
                if (!string.IsNullOrEmpty(postedVchrNo))
                {
                    await _reportNavService.PrintVoucher(postedVchrNo);
                }

                if (closeAfterSave)
                {
                    NavigateBack();
                }
                else
                {
                    ResetForm();
                }
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
            SelectedEmployee     = null;
            EmployeeSummary      = null;
            EmployeePicBase64    = string.Empty;
            ClearDate            = DateTime.Today;
            AmountReceived       = 0;
            NewDeduction         = 0;
            Description          = string.Empty;
            DontPostVoucher      = false;
            ReceiptMode          = ClearAdvanceReceiptMode.Cash;
            SelectedOtherAccount = null;

            if (CashAccounts.Any())
            {
                SelectedCashAccount = CashAccounts.FirstOrDefault(a => a.AccTitle.Contains("Petty"))
                                      ?? CashAccounts.First();
            }
            else
            {
                SelectedCashAccount = null;
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string BuildEmployeeDisplayName(EmployeeListItemModel emp)
        {
            string display = emp.Name;
            if (!string.IsNullOrWhiteSpace(emp.FName))       display += $" - {emp.FName}";
            if (!string.IsNullOrWhiteSpace(emp.Designation)) display += $" {emp.Designation}";
            return display;
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
