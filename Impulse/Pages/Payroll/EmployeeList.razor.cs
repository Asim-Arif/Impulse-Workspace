using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;

namespace Impulse.Pages.Payroll
{
    public partial class EmployeeList : ComponentBase
    {
        [Inject] private IEmployeeService EmployeeService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;

        [Parameter]
        [SupplyParameterFromQuery(Name = "contractorsOnly")]
        public bool ContractorsOnly { get; set; } = false;

        private bool isLoading = true;
        private bool showInactive = false;
        private bool showPicture = false;
        
        private string searchText = string.Empty;

        private List<DepartmentListItemModel> departments = new();
        private List<EmployeeListItemModel> employeesForSelect = new();
        private List<EmployeeListItemModel> allEmployees = new();

        private DepartmentListItemModel? selectedDepartment = null;
        private EmployeeListItemModel? selectedEmployeeFilter = null;
        private EmployeeListItemModel? highlightedEmployee = null;

        private byte[]? selectedEmployeePic = null;
        private string selectedEmployeePicBase64 = string.Empty;

        // Modal State for making an employee inactive
        private bool showInactiveModal = false;
        private EmployeeListItemModel? employeeToInactivate = null;
        private DateTime discontinuationDate = DateTime.Today;
        private string discontinuationRemarks = string.Empty;

        // Company Name from general settings for reports
        private string companyName = string.Empty;

        // Search fields dropdown options removed — search now covers all fields

        protected override async Task OnInitializedAsync()
        {
            await LoadData();
        }

        protected override async Task OnParametersSetAsync()
        {
            // Re-load data if contractorsOnly parameter changes
            await LoadData();
        }

        private async Task LoadData()
        {
            isLoading = true;
            try
            {
                companyName = await EmployeeService.GetCompanyNameAsync();
                
                // Load departments
                departments = await EmployeeService.GetDepartmentsAsync(showInactive);
                
                // Clear selection if not in new list
                if (selectedDepartment != null && !departments.Any(d => d.DeptID == selectedDepartment.DeptID))
                {
                    selectedDepartment = null;
                }

                // Load all employees for departments
                allEmployees.Clear();
                foreach (var dept in departments)
                {
                    var deptEmps = await EmployeeService.GetEmployeesAsync(dept.DeptID, showInactive, ContractorsOnly);
                    allEmployees.AddRange(deptEmps);
                }

                // Populate employee filter dropdown
                await UpdateEmployeeFilterDropdown();
                
                highlightedEmployee = null;
                selectedEmployeePic = null;
                selectedEmployeePicBase64 = string.Empty;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Data", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task UpdateEmployeeFilterDropdown()
        {
            if (selectedDepartment == null)
            {
                employeesForSelect = allEmployees;
            }
            else
            {
                employeesForSelect = allEmployees.Where(e => e.DeptID == selectedDepartment.DeptID).ToList();
            }

            if (selectedEmployeeFilter != null && !employeesForSelect.Any(e => e.EmpID == selectedEmployeeFilter.EmpID))
            {
                selectedEmployeeFilter = null;
            }

            await Task.CompletedTask;
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return departments;

            return await Task.FromResult(departments
                .Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                            d.DeptID.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList());
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            selectedDepartment = dept;
            selectedEmployeeFilter = null;
            await UpdateEmployeeFilterDropdown();
            highlightedEmployee = null;
            selectedEmployeePic = null;
            selectedEmployeePicBase64 = string.Empty;
        }

        private async Task<IEnumerable<EmployeeListItemModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return employeesForSelect;

            return await Task.FromResult(employeesForSelect
                .Where(e => e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                            e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList());
        }

        private async Task OnEmployeeChanged(EmployeeListItemModel? emp)
        {
            selectedEmployeeFilter = emp;
            if (emp != null)
            {
                await HighlightEmployee(emp);
            }
        }

        private async Task OnShowInactiveChanged(ChangeEventArgs e)
        {
            showInactive = (bool)(e.Value ?? false);
            await LoadData();
        }

        private async Task RefreshList()
        {
            await LoadData();
            NotificationServiceManager.ShowSuccess("Success", "List refreshed successfully.");
        }

        private async Task HighlightEmployee(EmployeeListItemModel emp)
        {
            highlightedEmployee = emp;
            selectedEmployeePic = null;
            selectedEmployeePicBase64 = string.Empty;

            if (showPicture && highlightedEmployee != null)
            {
                await LoadEmployeePicture(highlightedEmployee.EmpID);
            }
        }

        private async Task OnShowPictureChanged(ChangeEventArgs e)
        {
            showPicture = (bool)(e.Value ?? false);
            if (showPicture && highlightedEmployee != null)
            {
                await LoadEmployeePicture(highlightedEmployee.EmpID);
            }
            else
            {
                selectedEmployeePic = null;
                selectedEmployeePicBase64 = string.Empty;
            }
        }

        private async Task LoadEmployeePicture(string empId)
        {
            try
            {
                selectedEmployeePic = await EmployeeService.GetEmployeePictureAsync(empId);
                if (selectedEmployeePic != null && selectedEmployeePic.Length > 0)
                {
                    selectedEmployeePicBase64 = Convert.ToBase64String(selectedEmployeePic);
                }
                else
                {
                    selectedEmployeePicBase64 = string.Empty;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading picture: {ex.Message}");
                selectedEmployeePicBase64 = string.Empty;
            }
        }

        // Toggles active state or triggers discontinuation remarks dialog
        private async Task ToggleActiveState(EmployeeListItemModel emp)
        {
            if (emp.Active)
            {
                // Prompt for date & remarks
                employeeToInactivate = emp;
                discontinuationDate = DateTime.Today;
                discontinuationRemarks = string.Empty;
                showInactiveModal = true;
            }
            else
            {
                // Reactivate immediately
                try
                {
                    await EmployeeService.SetEmployeeActiveStateAsync(emp.EmpID, true);
                    NotificationServiceManager.ShowSuccess("Re-activated", $"Employee '{emp.Name}' has been activated.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Activation Error", ex.Message);
                }
            }
        }

        private async Task SaveInactivation()
        {
            if (employeeToInactivate == null) return;

            try
            {
                await EmployeeService.SetEmployeeActiveStateAsync(
                    employeeToInactivate.EmpID, 
                    false, 
                    discontinuationDate, 
                    discontinuationRemarks);

                NotificationServiceManager.ShowSuccess("Deactivated", $"Employee '{employeeToInactivate.Name}' discontinued.");
                showInactiveModal = false;
                employeeToInactivate = null;
                await LoadData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Deactivation Error", ex.Message);
            }
        }

        private void CancelInactivation()
        {
            showInactiveModal = false;
            employeeToInactivate = null;
        }

        // Handles contractor permanent transfer
        private async Task TransferToPermanent(EmployeeListItemModel emp)
        {
            if (!emp.Active)
            {
                NotificationServiceManager.ShowWarning("Warning", "Only active employees can be transferred.");
                return;
            }

            // Mimics the VB6 code: REPLACE 'T' or 'TMP' with permanent prefixes like SLI, WIS, PSS
            string strTemp = "EMP";
            if (companyName.Equals("PAKSMITH GROUP", StringComparison.OrdinalIgnoreCase))
                strTemp = "PSS";
            else if (companyName.Equals("Weldon Instruments", StringComparison.OrdinalIgnoreCase))
                strTemp = "WIS";
            else if (companyName.Equals("Sunlike", StringComparison.OrdinalIgnoreCase))
                strTemp = "SLI";
            else if (companyName.Equals("QEL", StringComparison.OrdinalIgnoreCase))
                strTemp = "QEL";
            else if (companyName.Equals("Dr-Frgz", StringComparison.OrdinalIgnoreCase))
                strTemp = "DFG";
            else if (companyName.Equals("Tecno", StringComparison.OrdinalIgnoreCase))
                strTemp = "TCN";
            else if (companyName.Equals("Towne", StringComparison.OrdinalIgnoreCase))
                strTemp = "TWN";

            string newEmpId = emp.EmpID;
            if (companyName.Equals("Tecno", StringComparison.OrdinalIgnoreCase))
            {
                newEmpId = emp.EmpID.Replace("TMP", strTemp, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                // Replace 'T' at start or in code
                if (emp.EmpID.StartsWith("T", StringComparison.OrdinalIgnoreCase))
                {
                    newEmpId = strTemp + emp.EmpID.Substring(1);
                }
                else
                {
                    newEmpId = emp.EmpID.Replace("T", strTemp, StringComparison.OrdinalIgnoreCase);
                }
            }

            try
            {
                // Direct SQL update matching VB6 con.Execute
                string query = "";
                if (companyName.Equals("Tecno", StringComparison.OrdinalIgnoreCase))
                {
                    query = "UPDATE Employees SET TempEmp=0, EmpID=REPLACE(EmpID, 'TMP', @StrTemp) WHERE EmpID=@OldEmpID";
                }
                else
                {
                    query = "UPDATE Employees SET TempEmp=0, EmpID=REPLACE(EmpID, 'T', @StrTemp) WHERE EmpID=@OldEmpID";
                }

                await DbHelper.ExecuteAsync(query, new { StrTemp = strTemp, OldEmpID = emp.EmpID });
                NotificationServiceManager.ShowSuccess("Transferred", $"Employee '{emp.Name}' transferred to permanent successfully.");
                await LoadData();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Transfer Error", ex.Message);
            }
        }

        // Search filtering logic — matches against all visible columns
        private IEnumerable<EmployeeListItemModel> GetFilteredEmployees(string deptId)
        {
            var deptEmps = allEmployees.Where(e => e.DeptID == deptId);

            // Filter by department selection dropdown
            if (selectedDepartment != null && deptId != selectedDepartment.DeptID)
            {
                return Enumerable.Empty<EmployeeListItemModel>();
            }

            // Filter by employee selection dropdown
            if (selectedEmployeeFilter != null)
            {
                return deptEmps.Where(e => e.EmpID == selectedEmployeeFilter.EmpID);
            }

            // All-field text search
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                var q = searchText.Trim();
                return deptEmps.Where(e =>
                    e.EmpID.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    (e.FName != null && e.FName.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (e.Designation != null && e.Designation.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (e.Phone1 != null && e.Phone1.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    e.StartingSalary.ToString().Contains(q, StringComparison.OrdinalIgnoreCase)
                );
            }

            return deptEmps;
        }

        private string ConvertWin32ColorToHex(int? win32Color)
        {
            if (win32Color == null || win32Color == 0)
                return "#000000";

            int color = win32Color.Value;
            int r = color & 0xFF;
            int g = (color >> 8) & 0xFF;
            int b = (color >> 16) & 0xFF;

            return $"#{r:X2}{g:X2}{b:X2}";
        }

        // ==========================================
        // REPORT GENERATION (PRINTING)
        // ==========================================

        private string GetSelectionFormula(bool transport = false, bool tempEmpList = false)
        {
            string cond = showInactive ? "" : " {VEmp.Active}=True ";
            
            if (transport)
            {
                cond = string.IsNullOrEmpty(cond) ? " {VEmp.Transport}=true " : cond + " AND {VEmp.Transport}=true ";
            }

            if (ContractorsOnly)
            {
                cond = string.IsNullOrEmpty(cond) ? " {VEmp.TempDept}=True " : cond + " AND {VEmp.TempDept}=True ";
            }
            else
            {
                cond = string.IsNullOrEmpty(cond) ? " {VEmp.TempDept}=False " : cond + " AND {VEmp.TempDept}=False ";
            }

            if (tempEmpList)
            {
                cond += " AND {@TempEmployee}=True ";
            }

            return cond;
        }

        private async Task PrintEmpList(bool withExternalSalary = false, bool transport = false, bool oldID = false, bool designationwise = false)
        {
            string reportName = oldID ? ReportNames.Payroll.EmpListWithOldID
                              : designationwise ? ReportNames.Payroll.EmpListDesignationwise
                              : ReportNames.Payroll.EmpList;
            string selectionFormula = GetSelectionFormula(transport);

            var formulaValues = new Dictionary<string, object>
            {
                { "Company", companyName }
            };

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpListWithoutSal(bool withDOB = false, string reportCaption = "Employee List", bool tempEmpList = false)
        {
            string selectionFormula = GetSelectionFormula(false, tempEmpList);

            var formulaValues = new Dictionary<string, object>
            {
                { "Company", companyName },
                { "BirthDate", withDOB ? "True" : "False" },
                { "RptCaption", reportCaption }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpListWithoutSal,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpCard(string? empId = null, string? deptId = null)
        {
            string cond = "{Employees.Active}=True";
            
            if (ContractorsOnly)
                cond += " AND {Departments.TempDept}=True";
            else
                cond += " AND {Departments.TempDept}=False";

            if (!string.IsNullOrEmpty(empId))
            {
                cond = $"( {cond} ) AND {{Employees.EmpID}}='{empId}'";
            }
            else if (!string.IsNullOrEmpty(deptId))
            {
                cond = $"( {cond} ) AND {{Employees.DeptID}}='{deptId}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpCards,
                SelectionFormula = cond
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpDetail(string empId)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpDetail,
                SelectionFormula = $"{{Employees.empid}}='{empId}'"
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintAppForm()
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.AppForm,
                SelectionFormula = highlightedEmployee != null ? $"{{Employees.empid}}='{highlightedEmployee.EmpID}'" : ""
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpListCF()
        {
            string selectionFormula = GetSelectionFormula();
            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpCardFields,
                SelectionFormula = selectionFormula
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpListWCardID()
        {
            string selectionFormula = GetSelectionFormula();
            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpListWCardID,
                SelectionFormula = selectionFormula
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpListSalaryPercentage()
        {
            // Prompt for percentage could be a simple 75% default or dialog. Let's default to 75% or use a simple modal.
            int iPercentage = 75; 
            string selectionFormula = "{VEmp.Active}=True AND {VEmp.TempDept}=False";

            var formulaValues = new Dictionary<string, object>
            {
                { "Company", companyName },
                { "SalaryPercentageHead", $"Salary @ {iPercentage}%" },
                { "iPercentage", iPercentage }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpListPercentage,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintBloodGroupReport(string bloodGroup)
        {
            // Helper to print blood group report
            string selectionFormula = $"{GetSelectionFormula()} AND {{VEmp.BloodGroup}}='{bloodGroup}'";
            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpList,
                SelectionFormula = selectionFormula
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintEmpListDeptSelection()
        {
            // Normally prompts for departments, let's print for the current selected department
            string selectionFormula = "{VEmp.Active}=TRUE";
            if (selectedDepartment != null)
            {
                selectionFormula += $" AND {{VEmp.DeptID}}='{selectedDepartment.DeptID}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.Payroll.EmpList,
                SelectionFormula = selectionFormula
            };

            await ReportNavigation.PrintReportAsync(request);
        }
    }
}
