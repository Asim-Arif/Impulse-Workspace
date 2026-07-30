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

namespace Impulse.Pages.Payroll.ClearShortTermLoan
{
    public partial class ClearShortTermLoan : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private IClearShortTermService  _clearShortTermService { get; set; } = null!;
        [Inject] private IEmployeeService        _employeeService       { get; set; } = null!;
        [Inject] private IVoucherService         _voucherService        { get; set; } = null!;
        [Inject] private IDBHelperService        _dbHelperService       { get; set; } = null!;
        [Inject] private INotificationService    _notificationService   { get; set; } = null!;
        [Inject] private IReportNavigationService _reportNavService     { get; set; } = null!;
        [Inject] private NavigationManager       _navManager            { get; set; } = null!;

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
        private string                         EmployeePicBase64  { get; set; } = string.Empty;

        // ── Open Advances Table ───────────────────────────────────────────────
        private List<ShortTermAdvanceItem> OpenAdvances { get; set; } = new();

        // ── Clearance Fields ──────────────────────────────────────────────────
        private DateTime ClearDate { get; set; } = DateTime.Today;
        private string   Remarks   { get; set; } = string.Empty;

        // ── Receipt Mode ──────────────────────────────────────────────────────
        private ClearAdvanceReceiptMode ReceiptMode     { get; set; } = ClearAdvanceReceiptMode.Cash;
        private bool                    DontPostVoucher { get; set; } = false;

        // ── Account Lookups ───────────────────────────────────────────────────
        private List<ChartOfAccountsModel> CashAccounts { get; set; } = new();
        private List<ChartOfAccountsModel> AllAccounts  { get; set; } = new();

        private ChartOfAccountsModel? SelectedCashAccount  { get; set; }
        private ChartOfAccountsModel? SelectedOtherAccount { get; set; }

        // ── EmpShortTermAccNo (validated on init) ─────────────────────────────
        private string EmpShortTermAccNo { get; set; } = string.Empty;

        // ── Computed Properties ───────────────────────────────────────────────
        private string CurrentVchrType => DontPostVoucher
            ? "NO VCHR"
            : ReceiptMode == ClearAdvanceReceiptMode.Cash ? "CRV" : "JV";

        private string VchrBadgeClass => DontPostVoucher
            ? "badge bg-dark"
            : ReceiptMode == ClearAdvanceReceiptMode.Cash ? "badge bg-success" : "badge bg-secondary";

        private decimal OutstandingBalance => OpenAdvances.Sum(a => a.Amount);
        private decimal SelectedAmount     => OpenAdvances.Where(a => a.IsSelected).Sum(a => a.Amount);
        private bool    AreAllSelected     => OpenAdvances.Count > 0 && OpenAdvances.All(a => a.IsSelected);

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

                // Validate EmpShortTermAccNo is configured
                EmpShortTermAccNo = await _clearShortTermService.GetEmpShortTermAccNoAsync();
                if (string.IsNullOrWhiteSpace(EmpShortTermAccNo))
                {
                    _notificationService.ShowWarning(
                        "Configuration Missing",
                        "EmpShortTermAccNo is not set in General Data. Saving will not work until an administrator configures it.");
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
            OpenAdvances       = new List<ShortTermAdvanceItem>();
            EmployeePicBase64  = string.Empty;

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, false);
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee  = emp;
            EmployeePicBase64 = string.Empty;
            OpenAdvances      = new List<ShortTermAdvanceItem>();

            if (emp != null)
            {
                OpenAdvances = await _clearShortTermService.GetOpenAdvancesAsync(emp.EmpID);

                var picBytes = await _employeeService.GetEmployeePictureAsync(emp.EmpID);
                if (picBytes != null && picBytes.Length > 0)
                {
                    EmployeePicBase64 = Convert.ToBase64String(picBytes);
                }
            }
        }

        private void OnClearDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime d))
            {
                ClearDate = d;
            }
        }

        private void ToggleSelectAll(ChangeEventArgs e)
        {
            bool select = e?.Value is bool b && b;
            foreach (var adv in OpenAdvances)
            {
                adv.IsSelected = select;
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

            // ── Validation (order mirrors the legacy frmClearShortTerm) ──────
            if (SelectedEmployee == null)
            {
                _notificationService.ShowWarning("Validation", "Please select an employee.");
                return;
            }

            var selected = OpenAdvances.Where(a => a.IsSelected).ToList();
            if (selected.Count == 0)
            {
                _notificationService.ShowWarning("Validation", "No Entry to Clear.");
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

                if (string.IsNullOrWhiteSpace(EmpShortTermAccNo))
                {
                    _notificationService.ShowError("Configuration Error",
                        "EmpShortTermAccNo is not configured in General Data. Cannot save.");
                    return;
                }
            }

            // ── Build DTO ─────────────────────────────────────────────────────
            var dto = new ClearShortTermDto
            {
                EmpID            = SelectedEmployee.EmpID,
                EmpDisplayName   = BuildEmployeeDisplayName(SelectedEmployee),
                ClearDate        = ClearDate,
                Remarks          = Remarks ?? string.Empty,
                DontPostVoucher  = DontPostVoucher,
                ReceiptMode      = ReceiptMode,
                ReceiptAccNo     = receiptAccNo,
                SelectedAdvances = selected
            };

            // ── Save ──────────────────────────────────────────────────────────
            IsSaving = true;
            try
            {
                string postedVchrNo = await _clearShortTermService.SaveAsync(dto);

                _notificationService.ShowSuccess(
                    "Saved Successfully",
                    string.IsNullOrEmpty(postedVchrNo)
                        ? "Short Term Loan clearance recorded (no voucher posted)."
                        : $"Short Term Loan clearance recorded. Voucher: {postedVchrNo}");

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
            OpenAdvances         = new List<ShortTermAdvanceItem>();
            EmployeePicBase64    = string.Empty;
            ClearDate            = DateTime.Today;
            Remarks              = string.Empty;
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
