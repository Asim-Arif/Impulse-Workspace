using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.Setup;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Setup;
using Impulse.Services.WorkflowTasks;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Production.ApprovalHistory
{
    public partial class ApprovalHistory : ComponentBase
    {
        [Inject] private IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;
        [Inject] private IProductionDeletionDataAccess ProdDeletionDataAccess { get; set; } = default!;
        [Inject] private IUserPermissionService UserPermissionService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        protected bool IsLoading { get; set; } = false;
        protected string? ErrorMessage { get; set; }

        protected List<WorkflowConfigurationModel> WorkflowConfigs { get; set; } = new();
        protected List<WorkflowApprovalHistoryModel> AllHistoryItems { get; set; } = new();
        protected WorkflowApprovalHistoryModel? SelectedItem { get; set; }
        protected bool ShowDetailsModal { get; set; } = false;

        // Filter Controls
        protected string SelectedWorkflowCode { get; set; } = string.Empty;
        protected string SelectedStatus { get; set; } = string.Empty;
        protected DateTime? DateFrom { get; set; }
        protected DateTime? DateTo { get; set; }
        protected string SearchText { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            if (!UserPermissionService.HasOptionAccess("PrdApprovalHistory"))
            {
                NavigationManager.NavigateTo("/access-denied");
                return;
            }

            // Default date range: Last 30 days
            DateFrom = DateTime.Today.AddDays(-30);
            DateTo = DateTime.Today;

            await LoadWorkflowConfigurationsAsync();
            await LoadHistoryAsync();
        }

        protected async Task LoadWorkflowConfigurationsAsync()
        {
            try
            {
                WorkflowConfigs = await WorkflowTaskEngine.GetAllWorkflowConfigsAsync();
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to load workflow configurations: {ex.Message}";
            }
        }

        protected async Task LoadHistoryAsync()
        {
            IsLoading = true;
            ErrorMessage = null;

            try
            {
                AllHistoryItems = await WorkflowTaskEngine.GetApprovalHistoryAsync(
                    SelectedWorkflowCode,
                    DateFrom,
                    DateTo,
                    SelectedStatus);
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Error loading approval history: {ex.Message}";
                AllHistoryItems = new();
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task ResetFiltersAsync()
        {
            SelectedWorkflowCode = string.Empty;
            SelectedStatus = string.Empty;
            SearchText = string.Empty;
            DateFrom = DateTime.Today.AddDays(-30);
            DateTo = DateTime.Today;

            await LoadHistoryAsync();
        }

        protected IEnumerable<WorkflowApprovalHistoryModel> FilteredItems
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                {
                    return AllHistoryItems;
                }

                var query = SearchText.Trim();
                return AllHistoryItems.Where(x =>
                    (!string.IsNullOrEmpty(x.LotNo) && x.LotNo.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.OrderNo) && x.OrderNo.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.ItemCode) && x.ItemCode.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.ItemName) && x.ItemName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.ProcessName) && x.ProcessName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.MakerName) && x.MakerName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.TargetProcessName) && x.TargetProcessName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.RequestedBy) && x.RequestedBy.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.ActionBy) && x.ActionBy.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.WorkflowName) && x.WorkflowName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.RequestReason) && x.RequestReason.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(x.ActionRemarks) && x.ActionRemarks.Contains(query, StringComparison.OrdinalIgnoreCase)));
            }
        }

        protected void ViewDetails(WorkflowApprovalHistoryModel item)
        {
            SelectedItem = item;
            ShowDetailsModal = true;
        }

        protected void CloseDetails()
        {
            ShowDetailsModal = false;
            SelectedItem = null;
        }

        protected static string GetStatusBadgeClass(string? status) => (status?.ToLowerInvariant()) switch
        {
            "approved" => "bg-success text-white",
            "rejected" => "bg-danger text-white",
            "pending" => "bg-warning text-dark",
            _ => "bg-secondary text-white"
        };

        protected static string GetStatusIcon(string? status) => (status?.ToLowerInvariant()) switch
        {
            "approved" => "fas fa-check-circle",
            "rejected" => "fas fa-times-circle",
            "pending" => "fas fa-clock",
            _ => "fas fa-info-circle"
        };
    }
}
