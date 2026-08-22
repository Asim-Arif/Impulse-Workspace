using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Stock.RMIssuanceList
{
    public partial class RMIssuanceList : ComponentBase
    {
        [Inject] private IRMIssuanceListService RMIssuanceListService { get; set; }
        [Inject] private IRMIssuanceService RMIssuanceService { get; set; }
        [Inject] private Impulse.Services.IReportNavigationService ReportNavigation { get; set; }
        [Inject] private Impulse.Services.INotificationService NotificationService { get; set; }
        [Inject] private NavigationManager NavigationManager { get; set; }
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; }

        // State
        private bool IsLoading { get; set; } = false;
        private DateTime DateFrom { get; set; } = DateTime.Today.AddDays(-30);
        private DateTime DateTo { get; set; } = DateTime.Today;

        // Data
        private List<RMIssuanceListRowModel> Issuances = new();
        
        // Lookups
        private List<RMGroupLookupModel> AllGroups = new();
        private List<DepartmentLookupModel> AllDepartments = new();
        private List<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel> AllMaterials = new();
        private List<EmployeeLookupModel> AllEmployees = new();
        private List<MakerLookupModel> AllVendors = new();

        // Selections
        private RMGroupLookupModel? SelectedGroup;
        private DepartmentLookupModel? SelectedDepartment;
        private DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel? SelectedMaterial;
        private EmployeeLookupModel? SelectedEmployee;
        private MakerLookupModel? SelectedVendor;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                var groupsTask = RMIssuanceListService.GetRMGroupsAsync();
                var deptsTask = RMIssuanceListService.GetDepartmentsAsync();
                var empTask = RMIssuanceService.GetEmployeesAsync();
                var makerTask = RMIssuanceService.GetMakersAsync();
                var matTask = RMIssuanceService.GetMaterialsAsync();

                await Task.WhenAll(groupsTask, deptsTask, empTask, makerTask, matTask);

                AllGroups = groupsTask.Result.ToList();
                AllDepartments = deptsTask.Result.ToList();
                AllEmployees = empTask.Result.ToList();
                AllVendors = makerTask.Result.ToList();
                AllMaterials = matTask.Result.ToList();
                
                await RefreshData();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task RefreshData()
        {
            IsLoading = true;
            try
            {
                var result = await RMIssuanceListService.GetRMIssuanceListAsync(
                    DateFrom, 
                    DateTo, 
                    SelectedMaterial?.RMID1, 
                    SelectedGroup?.ID, 
                    SelectedEmployee?.EmpID, 
                    SelectedVendor?.VendID, 
                    SelectedDepartment?.DeptID);
                    
                Issuances = result.ToList();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading data", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void NavigateToNewIssuance()
        {
            NavigationManager.NavigateTo("/stock/rm-issuance");
        }

        private async Task ShowOptionsMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void PrintIssuanceReport()
        {
            string dateRange = $"'{DateFrom:dd-MMM-yyyy} to {DateTo:dd-MMM-yyyy}'";
            string strSelection = $"{{IssItemsSimple.DT}} in Date({DateFrom.Year}, {DateFrom.Month}, {DateFrom.Day}) to Date({DateTo.Year}, {DateTo.Month}, {DateTo.Day})";
            
            if (SelectedMaterial != null)
                strSelection += $" AND {{IssItemsSimpleDetail.RMID1}}='{SelectedMaterial.RMID1}'";
            if (SelectedGroup != null)
                strSelection += $" AND {{RMGroups.ID}}={SelectedGroup.ID}";
            if (SelectedEmployee != null)
                strSelection += $" AND {{IssItemsSimple.IssdToEmpID}}='{SelectedEmployee.EmpID}'";
            if (SelectedVendor != null)
                strSelection += $" AND {{VendIssued.VendID}}={SelectedVendor.VendID}";
            if (SelectedDepartment != null)
                strSelection += $" AND {{VEmp.DeptID}}='{SelectedDepartment.DeptID}'";

            ReportNavigation.PrintReportAsync(new Impulse.Services.ReportRequest
            {
                ReportName = "RMIssuanceList.rpt",
                SelectionFormula = strSelection,
                FormulaValues = new Dictionary<string, object> { { "DateRange", dateRange } }
            });
        }

        private void PrintSingleEntry(ItemClickEventArgs e)
        {
            if (e.Data is RMIssuanceListRowModel row)
            {
                ReportNavigation.PrintReportAsync(new Impulse.Services.ReportRequest
                {
                    ReportName = "IssuanceNote.rpt",
                    SelectionFormula = $"{{IssItemsSimple.IssNo}}='{row.IssNo}'"
                });
            }
        }

        private async Task DeleteIssuance(ItemClickEventArgs e)
        {
            if (e.Data is RMIssuanceListRowModel row)
            {
                bool confirm = await NotificationService.ShowQuestionNotification("Delete Issuance", $"Are you sure you want to delete issuance {row.IssNo}?");
                if (confirm)
                {
                    try
                    {
                        bool success = await RMIssuanceListService.DeleteIssuanceAsync(row.IssNo);
                        if (success)
                        {
                            NotificationService.ShowSuccess("Deleted", "Issuance deleted successfully.");
                            await RefreshData();
                        }
                        else
                        {
                            NotificationService.ShowError("Error", "Failed to delete issuance.");
                        }
                    }
                    catch (Exception ex)
                    {
                        NotificationService.ShowError("Error", ex.Message);
                    }
                }
            }
        }

        // Search Methods
        private Task<IEnumerable<RMGroupLookupModel>> SearchGroups(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Task.FromResult(AllGroups.Take(50));
            return Task.FromResult(AllGroups.Where(x => x.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).Take(50));
        }

        private Task<IEnumerable<DepartmentLookupModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Task.FromResult(AllDepartments.Take(50));
            return Task.FromResult(AllDepartments.Where(x => x.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).Take(50));
        }

        private Task<IEnumerable<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Task.FromResult(AllMaterials.Take(50));
            return Task.FromResult(AllMaterials.Where(m => 
                (m.RMName != null && m.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) || 
                (m.RMID1 != null && m.RMID1.Contains(searchText, StringComparison.OrdinalIgnoreCase)))
                .Take(50));
        }

        private Task<IEnumerable<EmployeeLookupModel>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Task.FromResult(AllEmployees.Take(50));
            return Task.FromResult(AllEmployees.Where(x => x.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).Take(50));
        }

        private Task<IEnumerable<MakerLookupModel>> SearchVendors(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return Task.FromResult(AllVendors.Take(50));
            return Task.FromResult(AllVendors.Where(x => x.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).Take(50));
        }
    }
}
