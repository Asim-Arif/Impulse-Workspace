using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.SalarySheet
{
    public partial class SalarySheet : ComponentBase
    {
        // ── Injected Services ─────────────────────────────────────────────────
        [Inject] private ISalarySheetService        _salarySheetService       { get; set; } = null!;
        [Inject] private IEmployeeService           _employeeService          { get; set; } = null!;
        [Inject] private INotificationService       _notificationService      { get; set; } = null!;
        [Inject] private NavigationManager          _navManager               { get; set; } = null!;
        [Inject] private IReportNavigationService   _reportNavigation         { get; set; } = null!;
        [Inject] private IBlazorContextMenuService  _blazorContextMenuService { get; set; } = null!;

        // ── Query Parameters ──────────────────────────────────────────────────
        [Parameter][SupplyParameterFromQuery] public string? returnUrl   { get; set; }
        [Parameter][SupplyParameterFromQuery] public bool   contractors { get; set; } = false;

        private string companyName = string.Empty;

        // ── UI State ──────────────────────────────────────────────────────────
        private bool IsLoading    { get; set; } = true;
        private bool IsSaving     { get; set; } = false;
        private bool IsFinalized  { get; set; } = false;
        private bool IsContractor { get; set; } = false;

        // ── Filter State ──────────────────────────────────────────────────────
        private SalarySheetFilterDto Filter { get; set; } = new();

        // ── Lookups & Selections ──────────────────────────────────────────────
        private List<DepartmentListItemModel> AllDepartments     { get; set; } = new();
        private List<EmployeeListItemModel>   AllEmployees       { get; set; } = new();
        private DepartmentListItemModel?      SelectedDepartment { get; set; }
        private EmployeeListItemModel?        SelectedEmployee   { get; set; }

        // ── Data & Metrics ────────────────────────────────────────────────────
        private List<SalarySheetRowModel> SalaryRows { get; set; } = new();
        private SalarySheetKpiSummaryModel KpiSummary { get; set; } = new();

        private List<SalarySheetRowModel> FilteredSalaryRows
        {
            get
            {
                var query = SalaryRows.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(Filter.SearchText))
                {
                    string q = Filter.SearchText.Trim();
                    query = query.Where(r =>
                        r.EmpID.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        r.EmpName.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        (r.DeptName ?? "").Contains(q, StringComparison.OrdinalIgnoreCase) ||
                        (r.Designation ?? "").Contains(q, StringComparison.OrdinalIgnoreCase)
                    );
                }

                return query.ToList();
            }
        }

        // ── Lifecycle ─────────────────────────────────────────────────────────

        protected override async Task OnInitializedAsync()
        {
            IsContractor = contractors;
            Filter.IsContractor = IsContractor;

            // Legacy (frmSalarySheet) runs this company with chkExternal = vbChecked, so all
            // finalized rows are stored with IsForSA = 1. Without this, the finalized check
            // looks for IsForSA = 0, never finds the saved sheet, and silently falls back to
            // the live calculation (where e.g. cleared short-term advances come back as 0).
            Filter.IsExternal = true;

            try
            {
                companyName = await _employeeService.GetCompanyNameAsync();
                AllDepartments = await _employeeService.GetDepartmentsAsync(false);
                AllEmployees   = await _employeeService.GetEmployeesAsync("0", false, IsContractor);
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

        // ── Data Fetch ────────────────────────────────────────────────────────

        private async Task LoadSalarySheetAsync()
        {
            IsLoading = true;
            try
            {
                var (rows, kpi, finalized) = await _salarySheetService.GetSalarySheetDataAsync(Filter);
                SalaryRows  = rows;
                KpiSummary  = kpi;
                IsFinalized = finalized;
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Fetch Error", ex.Message);
                SalaryRows = new List<SalarySheetRowModel>();
                KpiSummary = new SalarySheetKpiSummaryModel();
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task RefreshAsync()
        {
            await LoadSalarySheetAsync();
        }

        // ── Filter Handlers ───────────────────────────────────────────────────

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            SelectedEmployee   = null;
            Filter.DeptID      = dept?.DeptID ?? "0";
            Filter.EmpID       = "0";

            string deptId = dept?.DeptID ?? "0";
            AllEmployees = await _employeeService.GetEmployeesAsync(deptId, false, IsContractor);
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            SelectedEmployee = emp;
            Filter.EmpID     = emp?.EmpID ?? "0";
        }

        // ── Typeahead Search ──────────────────────────────────────────────────

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

        // ── Finalize Salary Sheet Action ──────────────────────────────────────

        private async Task FinalizeSalarySheetAsync()
        {
            if (IsSaving) return;

            if (!SalaryRows.Any())
            {
                _notificationService.ShowWarning("Validation", "No salary rows to finalize.");
                return;
            }

            var negativeSalaries = SalaryRows.Where(r => r.NetPayable < 0).ToList();
            if (negativeSalaries.Any())
            {
                bool forceFinalize = await _notificationService.ShowQuestionNotification(
                    "Negative Salary Warning",
                    $"There are {negativeSalaries.Count} employees with negative net payable salaries.\nDo you want to Force Salary Finalization?");
                if (!forceFinalize)
                {
                    return;
                }
            }
            else
            {
                bool confirm = await _notificationService.ShowQuestionNotification(
                    "Confirm Finalize",
                    $"Are you sure you want to finalize the salary sheet for {Filter.Month}/{Filter.Year}?");
                if (!confirm)
                {
                    return;
                }
            }

            IsSaving = true;
            try
            {
                await _salarySheetService.FinalizeSalarySheetAsync(Filter.Year, Filter.Month, Filter.IsExternal, SalaryRows);

                _notificationService.ShowSuccess(
                    "Salary Sheet Finalized",
                    $"Salary Sheet for {Filter.Month}/{Filter.Year} has been saved & finalized successfully!");

                await LoadSalarySheetAsync();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Finalization Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ── Context Menu Actions ──────────────────────────────────────────────

        private async Task TogglePaidStatusClicked(ItemClickEventArgs e)
        {
            if (e.Data is SalarySheetRowModel row)
            {
                try
                {
                    bool newPaidState = !row.IsPaid;
                    await _salarySheetService.UpdateSalaryPaidStatusAsync(row.EmpID, Filter.Year, Filter.Month, newPaidState);
                    row.IsPaid = newPaidState;

                    _notificationService.ShowSuccess(
                        "Status Updated",
                        $"Salary status for [{row.EmpName}] marked as {(newPaidState ? "PAID" : "UNPAID")}.");
                }
                catch (Exception ex)
                {
                    _notificationService.ShowError("Update Failed", ex.Message);
                }
            }
        }

        private bool ShowChangeLTDeductionModal = false;
        private bool IsSavingLTDeduction = false;
        private SalarySheetRowModel? SelectedRowForLTDeduction;
        private decimal NewLTDeductionAmount = 0;

        private decimal MaxAllowedLTDeduction => SelectedRowForLTDeduction != null
            ? Math.Max(SelectedRowForLTDeduction.PrevLTLoanBalance, SelectedRowForLTDeduction.LongTermLoanDeduction)
            : 0;

        private void OpenChangeLTDeductionModal(ItemClickEventArgs e)
        {
            if (e.Data is SalarySheetRowModel row)
            {
                if (IsFinalized)
                {
                    _notificationService.ShowWarning("Finalized", "Cannot modify long term deduction on a finalized salary sheet.");
                    return;
                }

                if (row.PrevLTLoanBalance <= 0 && row.LongTermLoanDeduction <= 0)
                {
                    _notificationService.ShowWarning("No Loan Balance", $"Employee [{row.EmpName}] has no outstanding Long Term loan balance.");
                    return;
                }

                SelectedRowForLTDeduction = row;
                NewLTDeductionAmount = row.LongTermLoanDeduction;
                ShowChangeLTDeductionModal = true;
                StateHasChanged();
            }
        }

        private void CloseChangeLTDeductionModal()
        {
            ShowChangeLTDeductionModal = false;
            SelectedRowForLTDeduction = null;
            StateHasChanged();
        }

        private async Task SaveChangeLTDeductionAsync()
        {
            if (SelectedRowForLTDeduction == null)
                return;

            if (NewLTDeductionAmount < 0 || NewLTDeductionAmount > MaxAllowedLTDeduction)
            {
                _notificationService.ShowWarning("Invalid Amount", $"Deduction amount must be between 0 and {MaxAllowedLTDeduction:N2}.");
                return;
            }

            IsSavingLTDeduction = true;
            StateHasChanged();
            try
            {
                DateTime dtMonthEnd = new DateTime(Filter.Year, Filter.Month, DateTime.DaysInMonth(Filter.Year, Filter.Month));
                await _salarySheetService.UpdateLongTermDeductionOverrideAsync(SelectedRowForLTDeduction.EmpID, dtMonthEnd, NewLTDeductionAmount);
                _notificationService.ShowSuccess("Deduction Updated", $"Long term loan deduction updated for [{SelectedRowForLTDeduction.EmpName}] to {NewLTDeductionAmount:N2}.");
                ShowChangeLTDeductionModal = false;
                await LoadSalarySheetAsync();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Update Failed", ex.Message);
            }
            finally
            {
                IsSavingLTDeduction = false;
                StateHasChanged();
            }
        }

        private async Task StopAllLTDeductionClicked()
        {
            if (IsFinalized)
            {
                _notificationService.ShowWarning("Finalized", "Cannot modify long term deductions on a finalized salary sheet.");
                return;
            }

            if (FilteredSalaryRows == null || !FilteredSalaryRows.Any())
                return;

            bool proceed = await _notificationService.ShowQuestionNotification(
                "Confirm Stop Deductions",
                $"Are you sure you want to stop all Long Term Loan deductions (set deduction to 0) for all {FilteredSalaryRows.Count} employee(s) in the current list?");
            if (!proceed)
                return;

            IsSaving = true;
            StateHasChanged();
            try
            {
                DateTime dtMonthEnd = new DateTime(Filter.Year, Filter.Month, DateTime.DaysInMonth(Filter.Year, Filter.Month));
                int count = 0;
                foreach (var row in FilteredSalaryRows)
                {
                    await _salarySheetService.UpdateLongTermDeductionOverrideAsync(row.EmpID, dtMonthEnd, 0);
                    count++;
                }

                _notificationService.ShowSuccess("Deductions Stopped", $"Long term loan deductions stopped (set to 0) for {count} employee(s).");
                await LoadSalarySheetAsync();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Stop Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private bool HasNegativeSalaries => FilteredSalaryRows != null && FilteredSalaryRows.Any(x => x.NetPayable < 0);

        private async Task TransferNegativeSalariesAsync()
        {
            if (SelectedDepartment != null && !string.IsNullOrEmpty(SelectedDepartment.DeptID) && SelectedDepartment.DeptID != "0" ||
                SelectedEmployee != null && !string.IsNullOrEmpty(SelectedEmployee.EmpID) && SelectedEmployee.EmpID != "0")
            {
                _notificationService.ShowWarning(
                    "Selection Required",
                    "Please Select All Departments and Employees !!!");
                return;
            }

            DateTime monthEnd = new DateTime(Filter.Year, Filter.Month, 1).AddMonths(1).AddDays(-1);
            if (monthEnd >= DateTime.Today)
            {
                bool proceed = await _notificationService.ShowQuestionNotification(
                    "Confirm Transfer",
                    "This is normally allowed after month end. Do you want to proceed with transferring negative salaries?");
                if (!proceed)
                    return;
            }
            else
            {
                bool proceed = await _notificationService.ShowQuestionNotification(
                    "Confirm Transfer",
                    "Are you sure you want to transfer negative salaries to Short Term Advances?");
                if (!proceed)
                    return;
            }

            IsSaving = true;
            StateHasChanged();

            try
            {
                int count = 0;
                DateTime monthDate = new DateTime(Filter.Year, Filter.Month, 1);
                foreach (var row in FilteredSalaryRows)
                {
                    if (row.NetPayable < 0 && row.ShortTermAdvanceDeduction > 0 && row.ShortTermAdvanceDeduction > Math.Abs(row.NetPayable))
                    {
                        await _salarySheetService.TransferNegativeSalaryAsync(row.EmpID, Math.Abs(row.NetPayable), monthDate);
                        count++;
                    }
                }

                _notificationService.ShowInformation(
                    "Adjusted...",
                    count > 0 ? $"Adjusted negative salaries for {count} employee(s)." : "No eligible negative salaries found to adjust.");

                await LoadSalarySheetAsync();
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Transfer Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ── Printing & Reports (Options Menu) ────────────────────────────────

        private async Task OpenPrintMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await _blazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private string GetSalarySheetSelectionFormula(bool isNegative = false)
        {
            var conditions = new List<string>();

            // Filter by selected period
            conditions.Add($"{{VSalarySheet.Month}}={Filter.Month} AND {{VSalarySheet.Year}}={Filter.Year}");

            if (SelectedDepartment != null && !string.IsNullOrEmpty(SelectedDepartment.DeptID))
            {
                conditions.Add($"{{VSalarySheet.DeptID}}='{SelectedDepartment.DeptID}'");
            }

            if (SelectedEmployee != null && !string.IsNullOrEmpty(SelectedEmployee.EmpID))
            {
                conditions.Add($"{{VSalarySheet.EmpID}}='{SelectedEmployee.EmpID}'");
            }

            if (isNegative)
            {
                conditions.Add("{@Balance}<0");
            }

            return string.Join(" AND ", conditions);
        }

        private Dictionary<string, object> GetSalaryReportFormulaValues(bool isSummary = false)
        {
            int year = Filter.Year;
            int month = Filter.Month;
            int totalDays = DateTime.DaysInMonth(year, month);
            
            int totalSundays = 0;
            for (int d = 1; d <= totalDays; d++)
            {
                if (new DateTime(year, month, d).DayOfWeek == DayOfWeek.Sunday)
                    totalSundays++;
            }

            string monthName = new DateTime(year, month, 1).ToString("MMM");
            string fromTo = $" For {monthName}-{year}";

            return new Dictionary<string, object>
            {
                { "Company", companyName },
                { "FromTo", $"'{fromTo}'" },
                { "MonthDays", totalDays },
                { "TotalMonthDays", totalDays },
                { "TotalSundays", totalSundays },
                { "TotalHolidays", 0 }
            };
        }

        private async Task PrintSalarySheetDetail(bool isNegative = false)
        {
            string reportName;
            if (companyName.Equals("Dr-Frgz", StringComparison.OrdinalIgnoreCase) ||
                companyName.Equals("Banzai", StringComparison.OrdinalIgnoreCase) ||
                companyName.Equals("IAA", StringComparison.OrdinalIgnoreCase) ||
                companyName.Equals("Kami", StringComparison.OrdinalIgnoreCase) ||
                companyName.Equals("Towne", StringComparison.OrdinalIgnoreCase))
            {
                reportName = ReportNames.Payroll.SalarySheet;
            }
            else if (companyName.Equals("Tecno", StringComparison.OrdinalIgnoreCase))
            {
                reportName = ReportNames.Payroll.SalarySheetDetail;
            }
            else
            {
                reportName = IsFinalized ? ReportNames.Payroll.SalarySheetDetail : ReportNames.Payroll.SalarySheetUnFinal;
            }

            string selectionFormula = GetSalarySheetSelectionFormula(isNegative);
            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintNegativeEmpList()
        {
            string period = $"{new DateTime(Filter.Year, Filter.Month, 1):MMMM} - {Filter.Year}";
            DateTime myDT = new DateTime(Filter.Year, Filter.Month, 1).AddMonths(1);
            string selectionFormula = $"{{Advances.Description}}='Negative Salary Adjustment Entry' AND {{Advances.DT}}=#{myDT:yyyy-MM-dd}#";

            var formulaValues = new Dictionary<string, object>
            {
                { "Company", companyName },
                { "Period", $"'{period}'" }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpListNegativeSalaries,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintSalarySheetSummary()
        {
            string selectionFormula = GetSalarySheetSelectionFormula(false);
            var formulaValues = GetSalaryReportFormulaValues(true);

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.SalarySheetSummary,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintSalarySheetInternal()
        {
            string selectionFormula = GetSalarySheetSelectionFormula(false);
            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.SalarySheetInternal,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintLateEmpList()
        {
            string selectionFormula = GetSalarySheetSelectionFormula(false);
            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.LateEmpList,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private Dictionary<string, object> CalculateCurrencyDenominations()
        {
            long thousands = 0;
            long fHundreds = 0;
            long hundreds = 0;
            long fifties = 0;
            long tens = 0;
            long fives = 0;
            long twos = 0;
            long ones = 0;
            decimal totalSal = 0;

            foreach (var row in FilteredSalaryRows)
            {
                if (row.NetPayable > 0)
                {
                    decimal rem = Math.Round(row.NetPayable, 0);
                    totalSal += rem;

                    long count = (long)(rem / 1000m);
                    thousands += count;
                    rem %= 1000m;

                    count = (long)(rem / 500m);
                    fHundreds += count;
                    rem %= 500m;

                    count = (long)(rem / 100m);
                    hundreds += count;
                    rem %= 100m;

                    count = (long)(rem / 50m);
                    fifties += count;
                    rem %= 50m;

                    count = (long)(rem / 10m);
                    tens += count;
                    rem %= 10m;

                    count = (long)(rem / 5m);
                    fives += count;
                    rem %= 5m;

                    count = (long)(rem / 2m);
                    twos += count;
                    rem %= 2m;

                    ones += (long)rem;
                }
            }

            return new Dictionary<string, object>
            {
                { "Company", companyName },
                { "Thousands", thousands },
                { "FHundreds", fHundreds },
                { "Hundreds", hundreds },
                { "Fifties", fifties },
                { "Tens", tens },
                { "Fives", fives },
                { "Tows", twos },
                { "Twos", twos },
                { "Ones", ones },
                { "TotalSal", $"'{totalSal:N0}'" }
            };
        }

        private async Task PrintCurrencyDenomination()
        {
            var formulaValues = CalculateCurrencyDenominations();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.RupeeStructure,
                SelectionFormula = "",
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintPaySlips(int reportType)
        {
            string reportName = reportType switch
            {
                1 => ReportNames.Payroll.PaySlipRule9,
                2 => ReportNames.Payroll.PaySlipMini,
                3 => ReportNames.Payroll.PaySlipFull,
                _ => ReportNames.Payroll.PaySlip
            };

            string selectionFormula = GetSalarySheetSelectionFormula(false);
            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintSalaryPacket()
        {
            string cond = "{VPaySlip.BankPymt}=False";
            string baseFilter = GetSalarySheetSelectionFormula(false);
            if (!string.IsNullOrEmpty(baseFilter))
                cond = $"({cond}) AND ({baseFilter})";

            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.SalaryPacket,
                SelectionFormula = cond,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintPayslipsInternal()
        {
            string selectionFormula = GetSalarySheetSelectionFormula(false);
            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.PaySlipInternal,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task PrintOTPaySheet()
        {
            string cond = "{@OHAmt}>0";
            string baseFilter = GetSalarySheetSelectionFormula(false);
            if (!string.IsNullOrEmpty(baseFilter))
                cond = $"({cond}) AND ({baseFilter})";

            var formulaValues = GetSalaryReportFormulaValues();

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.OverTimePaymentSheet,
                SelectionFormula = cond,
                FormulaValues = formulaValues
            };

            await _reportNavigation.PrintReportAsync(request);
        }

        private async Task RePostSalaryVouchers()
        {
            if (!IsFinalized)
            {
                _notificationService.ShowWarning(
                    "Cannot Re-Post",
                    "Salary Vouchers can only be re-posted for a finalized salary sheet.");
                return;
            }

            bool confirm = await _notificationService.ShowQuestionNotification(
                "Confirm Re-Post",
                "Are you sure you want to re-post salary vouchers?");
            if (!confirm)
                return;

            try
            {
                int count = await _salarySheetService.AddSalaryVouchersAsync(Filter.Year, Filter.Month);
                _notificationService.ShowInformation(
                    "Re-Post Salary Vouchers",
                    $"Salary Vouchers reposted successfully ({count} department JVs created).");
            }
            catch (Exception ex)
            {
                _notificationService.ShowError("Error Re-Posting Vouchers", ex.Message);
            }
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
