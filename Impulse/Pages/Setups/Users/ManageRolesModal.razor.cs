using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.Setups.Users
{
    public partial class ManageRolesModal : ComponentBase
    {
        [Inject]
        public IUserRoleDataAccess UserRoleDataAccess { get; set; } = null!;

        [Inject]
        public NotificationService NotificationService { get; set; } = null!;

        [Parameter]
        public bool IsOpen { get; set; }

        [Parameter]
        public EventCallback<bool> OnClose { get; set; }

        [Parameter]
        public EventCallback OnRolesChanged { get; set; }

        protected List<UserRoleModel> Roles { get; set; } = new();
        protected string NewRoleName { get; set; } = string.Empty;
        protected string? ErrorMessage { get; set; }
        protected bool IsLoading { get; set; }
        protected bool IsProcessing { get; set; }

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen)
            {
                NewRoleName = string.Empty;
                ErrorMessage = null;
                await LoadRolesAsync();
            }
        }

        protected async Task LoadRolesAsync()
        {
            IsLoading = true;
            try
            {
                Roles = await UserRoleDataAccess.GetAllRolesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load roles: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task HandleKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await AddRoleAsync();
            }
        }

        protected async Task AddRoleAsync()
        {
            ErrorMessage = null;
            if (string.IsNullOrWhiteSpace(NewRoleName))
            {
                ErrorMessage = "Role name cannot be empty.";
                return;
            }

            var trimmed = NewRoleName.Trim();
            if (trimmed.Length > 50)
            {
                ErrorMessage = "Role name cannot exceed 50 characters.";
                return;
            }

            IsProcessing = true;
            try
            {
                bool exists = await UserRoleDataAccess.RoleExistsAsync(trimmed);
                if (exists)
                {
                    ErrorMessage = $"Role '{trimmed}' already exists.";
                    return;
                }

                bool success = await UserRoleDataAccess.AddRoleAsync(trimmed);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Role Created",
                        Detail = $"Role '{trimmed}' was added successfully.",
                        Duration = 3000
                    });

                    NewRoleName = string.Empty;
                    await LoadRolesAsync();
                    await OnRolesChanged.InvokeAsync();
                }
                else
                {
                    ErrorMessage = "Failed to create role.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
            }
            finally
            {
                IsProcessing = false;
            }
        }

        protected async Task ConfirmDeleteRoleAsync(UserRoleModel role)
        {
            try
            {
                bool success = await UserRoleDataAccess.DeleteRoleAsync(role.User_Role);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Role Deleted",
                        Detail = $"Role '{role.User_Role}' was removed.",
                        Duration = 3000
                    });

                    await LoadRolesAsync();
                    await OnRolesChanged.InvokeAsync();
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = "Failed to delete role.",
                        Duration = 3000
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

        protected async Task CloseAsync()
        {
            await OnClose.InvokeAsync(false);
        }
    }
}
