using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using Impulse.Services.Payroll;

namespace Impulse.Pages.Payroll
{
    public partial class NewDepartment : ComponentBase
    {
        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = null!;

        [Inject]
        private IDepartmentService DepartmentService { get; set; } = null!;

        [Inject]
        private IVoucherService VoucherService { get; set; } = null!;

        [Parameter]
        public bool IsVisible { get; set; }

        [Parameter]
        public EventCallback OnSaved { get; set; }

        [Parameter]
        public EventCallback OnCancelled { get; set; }

        private DepartmentModel department = new();
        private List<ChartOfAccountsModel> Accounts = new();
        private bool isAdd = true;
        private bool isSaving = false;
        private string hexColor = "#000000";

        // Mapped models for typeaheads
        private ChartOfAccountsModel? selectedSalaryAcc;
        private ChartOfAccountsModel? selectedAccruedAcc;
        private ChartOfAccountsModel? selectedFoodAcc;
        private ChartOfAccountsModel? selectedUnionAcc;
        private ChartOfAccountsModel? selectedSchAcc;
        private ChartOfAccountsModel? selectedEOBIAcc;
        private ChartOfAccountsModel? selectedTaxAcc;
        private ChartOfAccountsModel? selectedOnePercentAcc;

        protected override async Task OnInitializedAsync()
        {
            Accounts = await VoucherService.GetTransactionalAccounts("Where Parent=0 AND Active=1");
        }

        public async Task OpenAsync(string? deptId = null)
        {
            isSaving = false;
            if (string.IsNullOrEmpty(deptId))
            {
                department = new DepartmentModel
                {
                    Active = true,
                    TempDept = false
                };
                department.DeptID = await DepartmentService.GetNextDeptIdAsync();
                isAdd = true;
                hexColor = "#000000";
                ClearAccountSelections();
            }
            else
            {
                var existing = await DepartmentService.GetDepartmentByIdAsync(deptId);
                if (existing != null)
                {
                    department = existing;
                    isAdd = false;
                    hexColor = ConvertWin32ColorToHex(department.Color);
                    MapAccountSelections();
                }
            }

            IsVisible = true;
            StateHasChanged();
        }

        private void MapAccountSelections()
        {
            selectedSalaryAcc = Accounts.FirstOrDefault(a => a.AccNo == department.SalaryAccNo);
            selectedAccruedAcc = Accounts.FirstOrDefault(a => a.AccNo == department.AccruedAccNo);
            selectedFoodAcc = Accounts.FirstOrDefault(a => a.AccNo == department.AccruedFoodAccNo);
            selectedUnionAcc = Accounts.FirstOrDefault(a => a.AccNo == department.AccruedUnionAccNo);
            selectedSchAcc = Accounts.FirstOrDefault(a => a.AccNo == department.SchAccNo);
            selectedEOBIAcc = Accounts.FirstOrDefault(a => a.AccNo == department.AccruedEOBIAccNo);
            selectedTaxAcc = Accounts.FirstOrDefault(a => a.AccNo == department.TaxAccNo);
            selectedOnePercentAcc = Accounts.FirstOrDefault(a => a.AccNo == department.OnePercentAccNo);
        }

        private void ClearAccountSelections()
        {
            selectedSalaryAcc = null;
            selectedAccruedAcc = null;
            selectedFoodAcc = null;
            selectedUnionAcc = null;
            selectedSchAcc = null;
            selectedEOBIAcc = null;
            selectedTaxAcc = null;
            selectedOnePercentAcc = null;
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchAccounts(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Accounts;

            return await Task.FromResult(Accounts
                .Where(x => x.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase) 
                         || x.AccNo.Contains(searchText))
                .ToList());
        }

        private void OnSalaryAccChanged(ChartOfAccountsModel? acc)
        {
            selectedSalaryAcc = acc;
            department.SalaryAccNo = acc?.AccNo;
        }

        private void OnAccruedAccChanged(ChartOfAccountsModel? acc)
        {
            selectedAccruedAcc = acc;
            department.AccruedAccNo = acc?.AccNo;
        }

        private void OnFoodAccChanged(ChartOfAccountsModel? acc)
        {
            selectedFoodAcc = acc;
            department.AccruedFoodAccNo = acc?.AccNo;
        }

        private void OnUnionAccChanged(ChartOfAccountsModel? acc)
        {
            selectedUnionAcc = acc;
            department.AccruedUnionAccNo = acc?.AccNo;
        }

        private void OnSchAccChanged(ChartOfAccountsModel? acc)
        {
            selectedSchAcc = acc;
            department.SchAccNo = acc?.AccNo;
        }

        private void OnEOBIAccChanged(ChartOfAccountsModel? acc)
        {
            selectedEOBIAcc = acc;
            department.AccruedEOBIAccNo = acc?.AccNo;
        }

        private void OnTaxAccChanged(ChartOfAccountsModel? acc)
        {
            selectedTaxAcc = acc;
            department.TaxAccNo = acc?.AccNo;
        }

        private void OnOnePercentAccChanged(ChartOfAccountsModel? acc)
        {
            selectedOnePercentAcc = acc;
            department.OnePercentAccNo = acc?.AccNo;
        }

        private void OnColorChanged(ChangeEventArgs e)
        {
            hexColor = e.Value?.ToString() ?? "#000000";
            department.Color = ConvertHexToWin32Color(hexColor);
        }

        private async Task<bool> SaveAsync()
        {
            if (string.IsNullOrWhiteSpace(department.Name))
            {
                NotificationServiceManager.ShowError("Validation Error", "Invalid Department Name");
                return false;
            }

            try
            {
                isSaving = true;
                if (department.TempDept == true)
                {
                    // Clear out account mapping values for temp departments (per VB6 code)
                    department.SalaryAccNo = null;
                    department.AccruedAccNo = null;
                    department.AccruedFoodAccNo = null;
                    department.AccruedUnionAccNo = null;
                    department.SchAccNo = null;
                    department.AccruedEOBIAccNo = null;
                    department.TaxAccNo = null;
                    department.OnePercentAccNo = null;
                }

                await DepartmentService.SaveDepartmentAsync(department, isAdd);
                NotificationServiceManager.ShowSuccess("Success", "Department saved successfully.");
                return true;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Save Error", ex.Message);
                return false;
            }
            finally
            {
                isSaving = false;
            }
        }

        private async Task SaveAndClose()
        {
            if (await SaveAsync())
            {
                IsVisible = false;
                await OnSaved.InvokeAsync();
            }
        }

        private async Task SaveAndNew()
        {
            if (await SaveAsync())
            {
                await OnSaved.InvokeAsync();
                await OpenAsync(null); // reopen as add/new
            }
        }

        private async Task Cancel()
        {
            IsVisible = false;
            await OnCancelled.InvokeAsync();
        }

        // Color conversion helpers
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

        private int ConvertHexToWin32Color(string hexColor)
        {
            if (string.IsNullOrEmpty(hexColor) || !hexColor.StartsWith("#") || hexColor.Length < 7)
                return 0;

            try
            {
                int r = Convert.ToInt32(hexColor.Substring(1, 2), 16);
                int g = Convert.ToInt32(hexColor.Substring(3, 2), 16);
                int b = Convert.ToInt32(hexColor.Substring(5, 2), 16);

                return r + (g << 8) + (b << 16);
            }
            catch
            {
                return 0;
            }
        }
    }
}
