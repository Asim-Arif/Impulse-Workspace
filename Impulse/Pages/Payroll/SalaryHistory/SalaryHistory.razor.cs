using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.SalaryHistory
{
    public partial class SalaryHistory : ComponentBase
    {
        private bool IsLoading = false;
        private string CurrentUserName = "System";
        private string CompanyName = string.Empty;

        private DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        private DateTime ToDate { get; set; } = DateTime.Today;

        private List<DepartmentListItemModel> Departments = new();
        private DepartmentListItemModel _selectedDepartment;
        public DepartmentListItemModel SelectedDepartment
        {
            get => _selectedDepartment;
            set
            {
                if (_selectedDepartment != value)
                {
                    _selectedDepartment = value;
                    _ = OnDepartmentChanged();
                }
            }
        }

        private List<EmployeeListItemModel> Employees = new();
        private EmployeeListItemModel SelectedEmployee { get; set; }

        private List<SalaryHistoryRowModel> SalaryRows = new();

        // Change Salary Modal Properties
        private bool ShowChangeSalaryModal = false;
        private bool IsSavingModal = false;
        private EmployeeListItemModel ModalSelectedEmployee { get; set; }
        private DateTime ModalEffectiveDate { get; set; } = DateTime.Today;
        private double ModalNewSalary { get; set; } = 0;
        private double ModalNewSalaryExt { get; set; } = 0;
        private string ModalDescription { get; set; } = "Salary Adjustment";

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                }

                CompanyName = await SalaryHistoryService.GetCompanyNameAsync();
                Departments = await SalaryHistoryService.GetDepartmentsAsync();
                Employees = await SalaryHistoryService.GetEmployeesAsync("0");

                await LoadSalaryHistoryAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Departments;

            return await Task.FromResult(Departments.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Employees;

            return await Task.FromResult(Employees.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || x.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployeesModal(string searchText)
        {
            var validEmps = Employees.Where(x => x.EmpID != "0").ToList();
            if (string.IsNullOrWhiteSpace(searchText))
                return validEmps;

            return await Task.FromResult(validEmps.Where(x => x.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || x.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnDepartmentChanged()
        {
            SelectedEmployee = null;
            string deptId = SelectedDepartment?.DeptID ?? "0";

            try
            {
                Employees = await SalaryHistoryService.GetEmployeesAsync(deptId);
                await LoadSalaryHistoryAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadSalaryHistoryAsync()
        {
            IsLoading = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var filter = new SalaryHistoryFilterModel
                {
                    DeptID = SelectedDepartment?.DeptID ?? "0",
                    EmpID = SelectedEmployee?.EmpID ?? "0",
                    FromDate = FromDate,
                    ToDate = ToDate
                };

                SalaryRows = await SalaryHistoryService.GetSalaryHistoryAsync(filter);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading salary history", ex.Message);
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void OpenChangeSalaryModal()
        {
            ModalSelectedEmployee = null;
            ModalEffectiveDate = DateTime.Today;
            ModalNewSalary = 0;
            ModalNewSalaryExt = 0;
            ModalDescription = "Salary Adjustment";
            ShowChangeSalaryModal = true;
        }

        private void CloseChangeSalaryModal()
        {
            ShowChangeSalaryModal = false;
        }

        private async Task SaveSalaryChangeAsync()
        {
            if (ModalSelectedEmployee == null || ModalSelectedEmployee.EmpID == "0")
            {
                NotificationService.ShowWarning("Warning", "Please select an employee.");
                return;
            }

            if (ModalNewSalary <= 0)
            {
                NotificationService.ShowWarning("Warning", "New salary must be greater than 0.");
                return;
            }

            IsSavingModal = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                var req = new ChangeSalaryRequestModel
                {
                    EmpID = ModalSelectedEmployee.EmpID,
                    EffectiveDate = ModalEffectiveDate,
                    NewSalary = ModalNewSalary,
                    NewSalaryExt = ModalNewSalaryExt,
                    Description = string.IsNullOrWhiteSpace(ModalDescription) ? "Salary Adjustment" : ModalDescription,
                    UserName = CurrentUserName
                };

                bool success = await SalaryHistoryService.SaveSalaryChangeAsync(req);
                if (success)
                {
                    NotificationService.ShowSuccess("Success", $"Salary updated successfully for employee {ModalSelectedEmployee.Name}.");
                    ShowChangeSalaryModal = false;

                    // Reload dropdowns and history grid
                    Employees = await SalaryHistoryService.GetEmployeesAsync(SelectedDepartment?.DeptID ?? "0");
                    await LoadSalaryHistoryAsync();
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to save salary change.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsSavingModal = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task PrintSalaryHistoryReport()
        {
            if (!SalaryRows.Any()) return;

            try
            {
                // Run SP_PrintEmpAbsentList first as required by legacy Crystal Report
                await SalaryHistoryService.PrepareAbsentListReportProcAsync(FromDate, ToDate);

                string formula = $"{{VSalaryLedger.DT}} = Date({FromDate:yyyy,MM,dd}) to Date({ToDate:yyyy,MM,dd})";

                string deptId = SelectedDepartment?.DeptID ?? "0";
                string empId = SelectedEmployee?.EmpID ?? "0";

                if (deptId == "0")
                {
                    if (empId != "0")
                        formula += $" and {{VSalaryLedger.EmpID}} = '{empId}'";
                }
                else
                {
                    if (empId == "0")
                        formula += $" and {{VSalaryLedger.DeptID}} = '{deptId}'";
                    else
                        formula += $" and {{VSalaryLedger.EmpID}} = '{empId}'";
                }

                formula += " and {VSalaryLedger.TempDept} = False and {VSalaryLedger.Active} = True";

                var req = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.EmpSalHistory,
                    SelectionFormula = formula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Company", CompanyName },
                        { "From", FromDate.ToString("ddd dd-MMM-yyyy") },
                        { "To", ToDate.ToString("ddd dd-MMM-yyyy") }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }
    }
}
