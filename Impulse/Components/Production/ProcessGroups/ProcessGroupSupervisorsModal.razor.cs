using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Components.Production.ProcessGroups
{
    public partial class ProcessGroupSupervisorsModal : ComponentBase
    {
        [Inject] public IProcessGroupService ProcessGroupService { get; set; } = default!;
        [Inject] public IUserDataAccess UserDataAccess { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public ProcessGroupListItem? Group { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }
        [Parameter] public EventCallback OnSaved { get; set; }

        public List<ProcessGroupHubOverviewDto> HubsList { get; set; } = new();
        public List<UserModel> AvailableUsers { get; set; } = new();
        public UserModel? SelectedTempUser { get; set; }

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen && Group != null)
            {
                await LoadDataAsync();
            }
        }

        public async Task LoadDataAsync()
        {
            if (Group == null) return;

            try
            {
                IsLoading = true;
                var hubsTask = ProcessGroupService.GetGroupHubOverviewAsync(Group.EntryID);
                var usersTask = UserDataAccess.GetUsersAsync(activeOnly: true);

                await Task.WhenAll(hubsTask, usersTask);

                HubsList = await hubsTask;
                AvailableUsers = await usersTask;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Hubs",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task<IEnumerable<UserModel>> SearchUsers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AvailableUsers;

            return await Task.FromResult(AvailableUsers.Where(u =>
                (!string.IsNullOrEmpty(u.EmpID) && u.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(u.UserName) && u.UserName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(u.EmployeeName) && u.EmployeeName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(u.Designation) && u.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(u.DepartmentName) && u.DepartmentName.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            ).ToList());
        }

        public void OnUserSelected(ProcessGroupHubOverviewDto hub, UserModel? user)
        {
            if (user == null || Group == null) return;

            // Check if user is already assigned to this hub
            if (hub.Supervisors.Any(s => s.UserID == user.UserID))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Already Assigned",
                    Detail = $"{user.EmployeeName ?? user.UserName} is already assigned to Hub {hub.Hub_Name}.",
                    Duration = 3000
                });
                SelectedTempUser = null;
                return;
            }

            hub.Supervisors.Add(new HubSupervisorDto
            {
                GroupID = Group.EntryID,
                Hub_Name = hub.Hub_Name,
                UserID = user.UserID,
                UserName = user.UserName,
                EmpID = user.EmpID,
                EmployeeName = user.EmployeeName,
                Designation = user.Designation
            });

            SelectedTempUser = null;
        }

        public void RemoveSupervisor(ProcessGroupHubOverviewDto hub, HubSupervisorDto sup)
        {
            hub.Supervisors.Remove(sup);
        }

        public async Task SaveAllSupervisorsAsync()
        {
            if (Group == null) return;

            try
            {
                IsSaving = true;

                foreach (var hub in HubsList)
                {
                    var userIds = hub.Supervisors.Select(s => s.UserID).ToList();
                    await ProcessGroupService.SaveGroupHubSupervisorsAsync(Group.EntryID, hub.Hub_Name, userIds);
                }

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Supervisors Saved",
                    Detail = $"Hub supervisors updated successfully for group '{Group.GroupName}'.",
                    Duration = 4000
                });

                await OnSaved.InvokeAsync();
                await CloseModal();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Save Supervisors",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task CloseModal()
        {
            SelectedTempUser = null;
            await OnClosed.InvokeAsync();
        }
    }
}
