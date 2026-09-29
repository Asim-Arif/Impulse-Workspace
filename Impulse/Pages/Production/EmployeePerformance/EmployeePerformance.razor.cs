using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.EmployeePerformance
{
    public partial class EmployeePerformance : ComponentBase
    {
        [Inject] public IEmployeePerformanceService PerformanceService { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;

        public EmployeePerformanceFilter Filter { get; set; } = new EmployeePerformanceFilter();
        public EmployeePerformanceDashboardDto DashboardData { get; set; } = new EmployeePerformanceDashboardDto();

        public List<DepartmentLookupItem> Departments { get; set; } = new();
        public List<ProcessLookupItem> Processes { get; set; } = new();
        public List<PerformanceEmployeeLookupItem> Employees { get; set; } = new();

        public bool IsLoading { get; set; } = false;
        public string ClientSearchTerm { get; set; } = string.Empty;

        // Modal Drill-down state
        public bool ShowLotsModal { get; set; } = false;
        public bool IsLoadingLotsModal { get; set; } = false;
        public EmployeePerformanceSummaryDto? SelectedEmployeeForModal { get; set; }
        public List<EmployeeLotDetailDto> SelectedEmployeeLots { get; set; } = new();
        public string LotsSearchTerm { get; set; } = string.Empty;

        public List<EmployeeLotDetailDto> FilteredEmployeeLots
        {
            get
            {
                if (string.IsNullOrWhiteSpace(LotsSearchTerm))
                    return SelectedEmployeeLots;

                string term = LotsSearchTerm.Trim().ToLowerInvariant();
                return SelectedEmployeeLots.Where(l =>
                    (l.LotNo != null && l.LotNo.ToLowerInvariant().Contains(term)) ||
                    (l.OrderNo != null && l.OrderNo.ToLowerInvariant().Contains(term)) ||
                    (l.ItemCode != null && l.ItemCode.ToLowerInvariant().Contains(term)) ||
                    (l.ItemName != null && l.ItemName.ToLowerInvariant().Contains(term)) ||
                    (l.ProcessName != null && l.ProcessName.ToLowerInvariant().Contains(term)) ||
                    (l.Status != null && l.Status.ToLowerInvariant().Contains(term))
                ).ToList();
            }
        }

        public List<EmployeePerformanceSummaryDto> FilteredEmployees
        {
            get
            {
                var list = DashboardData.Employees ?? new List<EmployeePerformanceSummaryDto>();
                if (string.IsNullOrWhiteSpace(ClientSearchTerm))
                    return list;

                string term = ClientSearchTerm.Trim().ToLowerInvariant();
                return list.Where(e =>
                    (e.EmpID != null && e.EmpID.ToLowerInvariant().Contains(term)) ||
                    (e.EmployeeName != null && e.EmployeeName.ToLowerInvariant().Contains(term)) ||
                    (e.Designation != null && e.Designation.ToLowerInvariant().Contains(term)) ||
                    (e.DepartmentName != null && e.DepartmentName.ToLowerInvariant().Contains(term)) ||
                    (e.PrimaryProcessName != null && e.PrimaryProcessName.ToLowerInvariant().Contains(term)) ||
                    (e.PerformanceStatus != null && e.PerformanceStatus.ToLowerInvariant().Contains(term))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;
                ApplyDatePreset(Filter.DateRangeType);

                var deptTask = PerformanceService.GetDepartmentsLookupAsync();
                var procTask = PerformanceService.GetProcessesLookupAsync();
                var empTask = PerformanceService.GetEmployeesLookupAsync();

                await Task.WhenAll(deptTask, procTask, empTask);

                Departments = await deptTask;
                Processes = await procTask;
                Employees = await empTask;

                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Init Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                DashboardData = await PerformanceService.GetPerformanceDashboardAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Data Error", $"Failed to load performance data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task OnDatePresetChanged()
        {
            ApplyDatePreset(Filter.DateRangeType);
            await LoadDataAsync();
        }

        private void ApplyDatePreset(int preset)
        {
            DateTime now = DateTime.Today;
            switch (preset)
            {
                case 0: // Today
                    Filter.DtFrom = now;
                    Filter.DtTo = now;
                    break;
                case 1: // This Week
                    int diff = (7 + (now.DayOfWeek - DayOfWeek.Monday)) % 7;
                    Filter.DtFrom = now.AddDays(-1 * diff).Date;
                    Filter.DtTo = now;
                    break;
                case 2: // Last 15 Days
                    Filter.DtFrom = now.AddDays(-15);
                    Filter.DtTo = now;
                    break;
                case 3: // This Month
                    Filter.DtFrom = new DateTime(now.Year, now.Month, 1);
                    Filter.DtTo = now;
                    break;
                case 4: // Last 30 Days
                    Filter.DtFrom = now.AddDays(-30);
                    Filter.DtTo = now;
                    break;
                case 5: // Custom
                    break;
            }
        }

        public async Task OnFilterChanged()
        {
            await LoadDataAsync();
        }

        public async Task OnDepartmentChanged(int? deptId)
        {
            Filter.DepartmentId = deptId;
            Employees = await PerformanceService.GetEmployeesLookupAsync(deptId);
            await LoadDataAsync();
        }

        public async Task ResetFiltersAsync()
        {
            Filter = new EmployeePerformanceFilter
            {
                DateRangeType = 3,
                StatusFilter = 0
            };
            ApplyDatePreset(Filter.DateRangeType);
            ClientSearchTerm = string.Empty;

            Employees = await PerformanceService.GetEmployeesLookupAsync();
            await LoadDataAsync();
        }

        public async Task OpenLotsModalAsync(EmployeePerformanceSummaryDto emp)
        {
            SelectedEmployeeForModal = emp;
            ShowLotsModal = true;
            IsLoadingLotsModal = true;
            LotsSearchTerm = string.Empty;

            try
            {
                SelectedEmployeeLots = await PerformanceService.GetEmployeeLotsDetailAsync(emp.EmpID, Filter.DtFrom, Filter.DtTo);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Lots Error", ex.Message);
            }
            finally
            {
                IsLoadingLotsModal = false;
            }
        }

        public void CloseLotsModal()
        {
            ShowLotsModal = false;
            SelectedEmployeeForModal = null;
            SelectedEmployeeLots.Clear();
            LotsSearchTerm = string.Empty;
        }

        public string GetCapacityBarClass(double pct) => pct switch
        {
            > 100 => "bg-danger",
            >= 75 => "bg-success",
            >= 50 => "bg-primary",
            _ => "bg-warning"
        };
    }
}
