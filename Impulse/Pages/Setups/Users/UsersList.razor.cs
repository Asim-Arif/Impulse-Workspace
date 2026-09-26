using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.Setups.Users
{
    public partial class UsersList : ComponentBase
    {
        [Inject]
        protected IUserService UserService { get; set; } = default!;

        [Inject]
        protected IUserPermissionService PermissionService { get; set; } = default!;

        [Inject]
        protected NotificationService NotificationService { get; set; } = default!;

        [Inject]
        protected NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        protected DataAccessLibrary.Interface.Setup.IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        protected List<UserModel> AllUsers { get; set; } = new();
        protected List<UserModel> FilteredUsers { get; set; } = new();
        protected List<UserRoleModel> AvailableRoles { get; set; } = new();
        protected string SearchQuery { get; set; } = string.Empty;
        protected string StatusFilter { get; set; } = "all";
        protected string RoleFilter { get; set; } = "all";
        protected bool IsLoading { get; set; } = true;
        protected bool IsSyncing { get; set; } = false;

        // Manage Roles Modal State
        protected bool IsManageRolesOpen { get; set; } = false;

        // Modal States
        protected bool IsModalOpen { get; set; } = false;
        protected bool IsCreatingNew { get; set; } = false;
        protected UserModel EditingUser { get; set; } = new();

        // Password Reset Modal States
        protected bool IsResetPasswordOpen { get; set; } = false;
        protected UserModel? ResettingUser { get; set; }
        protected string NewPasswordInput { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            await LoadUsersAsync();
        }

        protected async Task LoadUsersAsync()
        {
            try
            {
                IsLoading = true;
                AvailableRoles = await UserRoleDataAccess.GetAllRolesAsync();
                AllUsers = await UserService.GetUsersAsync();
                ApplyFilter();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load users: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected void ApplyFilter()
        {
            var query = AllUsers.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(SearchQuery))
            {
                var q = SearchQuery.Trim().ToLowerInvariant();
                query = query.Where(u =>
                    (!string.IsNullOrEmpty(u.UserName) && u.UserName.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(u.FullUserName) && u.FullUserName.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(u.EmployeeName) && u.EmployeeName.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(u.EmpID) && u.EmpID.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(u.DepartmentName) && u.DepartmentName.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(u.Designation) && u.Designation.ToLowerInvariant().Contains(q))
                );
            }

            if (StatusFilter == "active")
            {
                query = query.Where(u => !u.InActive);
            }
            else if (StatusFilter == "inactive")
            {
                query = query.Where(u => u.InActive);
            }

            if (!string.IsNullOrEmpty(RoleFilter) && RoleFilter != "all")
            {
                query = query.Where(u => u.AssignedRoles != null && u.AssignedRoles.Any(r => string.Equals(r, RoleFilter, StringComparison.OrdinalIgnoreCase)));
            }

            FilteredUsers = query.ToList();
        }

        protected Task ApplyFilterAsync()
        {
            ApplyFilter();
            return Task.CompletedTask;
        }

        protected Task ResetFilterAsync()
        {
            SearchQuery = string.Empty;
            StatusFilter = "all";
            RoleFilter = "all";
            ApplyFilter();
            return Task.CompletedTask;
        }

        protected void OpenManageRolesModal()
        {
            IsManageRolesOpen = true;
        }

        protected void CloseManageRolesModal(bool changed)
        {
            IsManageRolesOpen = false;
        }

        protected async Task HandleRolesChanged()
        {
            await LoadUsersAsync();
        }

        protected string GetRoleBadgeClass(string role)
        {
            return (role?.ToLowerInvariant()) switch
            {
                "stock" => "bg-purple-subtle text-purple border",
                "ppc" => "bg-info-subtle text-info-emphasis border border-info-subtle",
                "purchaser" or "purchase" => "bg-warning-subtle text-warning-emphasis border border-warning-subtle",
                "production" or "factory" => "bg-danger-subtle text-danger border border-danger-subtle",
                "accounts" or "finance" => "bg-primary-subtle text-primary border border-primary-subtle",
                _ => "bg-secondary-subtle text-secondary border border-secondary-subtle"
            };
        }

        protected void HandleSearchKeyUp(KeyboardEventArgs e)
        {
            ApplyFilter();
        }

        protected void OpenCreateModal()
        {
            NavigationManager.NavigateTo("/setup/users/new?returnUrl=/setup/users");
        }

        protected void OpenEditModal(UserModel user)
        {
            NavigationManager.NavigateTo($"/setup/users/edit/{user.UserID}?returnUrl=/setup/users");
        }

        protected void CloseModal(bool changed)
        {
            IsModalOpen = false;
        }

        protected async Task HandleUserSaved(UserModel savedUser)
        {
            IsModalOpen = false;
            await LoadUsersAsync();
        }

        protected async Task ToggleStatusAsync(UserModel user)
        {
            var newStatus = !user.InActive;
            var success = await UserService.ToggleUserStatusAsync(user.UserID, newStatus);
            if (success)
            {
                user.InActive = newStatus;
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = newStatus ? NotificationSeverity.Warning : NotificationSeverity.Success,
                    Summary = newStatus ? "Account Suspended" : "Account Activated",
                    Detail = $"User '{user.UserName}' status updated.",
                    Duration = 3000
                });
                ApplyFilter();
            }
            else
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = "Failed to update account status.",
                    Duration = 3000
                });
            }
        }

        protected void OpenResetPasswordModal(UserModel user)
        {
            ResettingUser = user;
            NewPasswordInput = string.Empty;
            IsResetPasswordOpen = true;
        }

        protected async Task ConfirmResetPasswordAsync()
        {
            if (ResettingUser == null || string.IsNullOrWhiteSpace(NewPasswordInput))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a valid new password.",
                    Duration = 3000
                });
                return;
            }

            var (success, msg) = await UserService.ResetPasswordAsync(ResettingUser.UserID, NewPasswordInput);
            if (success)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Password Reset",
                    Detail = $"Password for '{ResettingUser.UserName}' updated successfully.",
                    Duration = 3000
                });
                IsResetPasswordOpen = false;
            }
            else
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = msg,
                    Duration = 4000
                });
            }
        }

        protected async Task SyncMenuOptionsAsync()
        {
            try
            {
                IsSyncing = true;
                var success = await PermissionService.ResetAllOptionsAsync();
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Options Synced",
                        Detail = "System screen menu options catalog updated successfully.",
                        Duration = 3500
                    });
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = "Failed to synchronize menu options catalog.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSyncing = false;
            }
        }

        protected bool IsCopyUserOpen { get; set; } = false;
        protected UserModel? CopyingSourceUser { get; set; }
        protected string NewCopyUserName { get; set; } = string.Empty;
        protected string NewCopyPassword { get; set; } = string.Empty;
        protected string NewCopyFullName { get; set; } = string.Empty;
        protected bool IsCopyingUser { get; set; } = false;

        protected void OpenCopyUserModal(UserModel user)
        {
            CopyingSourceUser = user;
            NewCopyUserName = $"Copy Of {user.UserName}";
            NewCopyPassword = string.Empty;
            NewCopyFullName = user.FullUserName ?? string.Empty;
            IsCopyUserOpen = true;
        }

        protected void CloseCopyUserModal()
        {
            IsCopyUserOpen = false;
            CopyingSourceUser = null;
        }

        protected async Task ExecuteCopyUserAsync()
        {
            if (CopyingSourceUser == null) return;

            if (string.IsNullOrWhiteSpace(NewCopyUserName))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "New Username is required.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                IsCopyingUser = true;
                var (success, msg, newId) = await UserService.CopyUserAsync(
                    CopyingSourceUser.UserID,
                    NewCopyUserName.Trim(),
                    string.IsNullOrWhiteSpace(NewCopyPassword) ? null : NewCopyPassword.Trim(),
                    string.IsNullOrWhiteSpace(NewCopyFullName) ? null : NewCopyFullName.Trim()
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "User Copied",
                        Detail = msg,
                        Duration = 3000
                    });
                    CloseCopyUserModal();
                    await LoadUsersAsync();

                    // Open the newly created user in edit page so admin can fine-tune
                    NavigationManager.NavigateTo($"/setup/users/edit/{newId}?returnUrl=/setup/users");
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Copy Failed",
                        Detail = msg,
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsCopyingUser = false;
            }
        }
    }
}
