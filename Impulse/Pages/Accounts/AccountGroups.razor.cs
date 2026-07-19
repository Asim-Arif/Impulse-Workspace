using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Interface.Accounts;
using Impulse.Services;
using BlazorContextMenu;

namespace Impulse.Pages.Accounts
{
    public partial class AccountGroups : ComponentBase
    {
        [Inject]
        protected IChartOfAccountsDataAccess ChartOfAccountsService { get; set; }

        [Inject]
        protected NavigationManager Navigation { get; set; }

        [Inject]
        private INotificationService NotificationServiceManager { get; set; }

        private List<ReportingGroupModel> GroupsList { get; set; } = new List<ReportingGroupModel>();
        private ReportingGroupModel SelectedGroup { get; set; }
        private ReportingGroupModel CurrentGroup { get; set; } = new ReportingGroupModel();
        
        private bool IsLoading;
        private bool ShowModal;
        private bool IsEdit;

        protected override async Task OnInitializedAsync()
        {
            await LoadData();
        }

        private async Task LoadData()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                GroupsList = await ChartOfAccountsService.GetReportingGroupsAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
                NotificationServiceManager.ShowError("Error", "Failed to load Account Groups.");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private void ShowAddModal()
        {
            int nextSortNo = GroupsList != null && GroupsList.Any() ? GroupsList.Max(g => g.SortNo) + 1 : 1;
            CurrentGroup = new ReportingGroupModel { Type = 0, SortNo = nextSortNo }; // Default to Expense
            IsEdit = false;
            ShowModal = true;
        }

        private void ShowEditModal(ItemClickEventArgs e)
        {
            var group = e.Data as ReportingGroupModel;
            if (group != null)
            {
                CurrentGroup = new ReportingGroupModel
                {
                    EntryID = group.EntryID,
                    GroupName = group.GroupName,
                    Type = group.Type,
                    SortNo = group.SortNo
                };
                IsEdit = true;
                ShowModal = true;
            }
        }

        private void HideModal()
        {
            ShowModal = false;
            CurrentGroup = new ReportingGroupModel();
        }

        private async Task SaveGroup()
        {
            if (string.IsNullOrWhiteSpace(CurrentGroup.GroupName))
            {
                NotificationServiceManager.ShowWarning("Validation", "Group Name is required.");
                return;
            }

            try
            {
                if (IsEdit)
                {
                    await ChartOfAccountsService.UpdateReportingGroupDetailsAsync(CurrentGroup);
                    NotificationServiceManager.ShowSuccess("Success", "Account Group updated successfully.");
                }
                else
                {
                    await ChartOfAccountsService.AddReportingGroupAsync(CurrentGroup);
                    NotificationServiceManager.ShowSuccess("Success", "Account Group added successfully.");
                }

                HideModal();
                await LoadData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving Data: {ex.Message}");
                NotificationServiceManager.ShowError("Error", "Failed to save Account Group.");
            }
        }

        private async Task DeleteSelectedGroup(ItemClickEventArgs e)
        {
            var group = e.Data as ReportingGroupModel;
            if (group != null)
            {
                try
                {
                    await ChartOfAccountsService.DeleteReportingGroupAsync(group.EntryID);
                    NotificationServiceManager.ShowSuccess("Success", "Account Group deleted successfully.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting Data: {ex.Message}");
                    NotificationServiceManager.ShowError("Error", "Failed to delete Account Group. It might be in use.");
                }
            }
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/financial", true);
        }
    }
}
