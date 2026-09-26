using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Setups.Users
{
    public partial class UserEdit : ComponentBase
    {
        [Inject]
        protected IUserService UserService { get; set; } = default!;

        [Inject]
        protected IEmployeeService EmployeeService { get; set; } = default!;

        [Inject]
        protected IUserPermissionService PermissionService { get; set; } = default!;

        [Inject]
        protected NotificationService NotificationService { get; set; } = default!;

        [Inject]
        protected NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        protected DataAccessLibrary.Interface.Setup.IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        [Parameter]
        public int? UserId { get; set; }

        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        protected UserModel User { get; set; } = new UserModel();
        protected bool IsNewUser { get; set; } = false;
        protected bool IsLoadingUser { get; set; } = true;
        protected int ActiveTab { get; set; } = 0;
        protected bool IsSaving { get; set; } = false;
        protected EmployeeListItemModel? SelectedEmployee { get; set; }
        private List<EmployeeListItemModel> _allEmployees = new();

        // Operational Roles State
        protected List<UserRoleModel> AvailableRoles { get; set; } = new();
        protected HashSet<string> SelectedRoles { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        protected bool IsManageRolesOpen { get; set; } = false;

        // Screen & Menu Options State
        protected List<string> AvailableModules { get; set; } = new();
        protected string SelectedModule { get; set; } = "Accounts";
        protected List<MenuOptionModel> CurrentModuleOptions { get; set; } = new();
        protected HashSet<string> UserSelectedOptionIds { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        protected string OptionSearchFilter { get; set; } = string.Empty;
        protected bool IsLoadingOptions { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoadingUser = true;

                await LoadEmployeesAsync();
                AvailableRoles = await UserRoleDataAccess.GetAllRolesAsync();

                if (UserId.HasValue && UserId.Value > 0)
                {
                    IsNewUser = false;
                    var existingUser = await UserService.GetUserByIdAsync(UserId.Value);
                    if (existingUser != null)
                    {
                        User = existingUser;
                        var userRoles = await UserRoleDataAccess.GetRolesByUserIdAsync(UserId.Value);
                        SelectedRoles = new HashSet<string>(userRoles, StringComparer.OrdinalIgnoreCase);
                    }
                    else
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"User ID #{UserId.Value} could not be located.",
                            Duration = 4000
                        });
                        NavigateBack();
                        return;
                    }
                }
                else
                {
                    IsNewUser = true;
                    User = new UserModel
                    {
                        InActive = false,
                        ChangePassword = true,
                        DashBoardMainLink = true,
                        OpenCommandCenter = true,
                        ShowTips = true
                    };
                }

                if (!string.IsNullOrEmpty(User.EmpID))
                {
                    SelectedEmployee = _allEmployees.FirstOrDefault(e => e.EmpID == User.EmpID);
                }

                await LoadPermissionsDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load user: {ex.Message}",
                    Duration = 4000
                });
            }
            finally
            {
                IsLoadingUser = false;
            }
        }

        private async Task LoadPermissionsDataAsync()
        {
            try
            {
                IsLoadingOptions = true;

                // 1. Fetch available modules
                AvailableModules = await PermissionService.GetDistinctModulesAsync();
                if (!AvailableModules.Any())
                {
                    await PermissionService.ResetAllOptionsAsync();
                    AvailableModules = await PermissionService.GetDistinctModulesAsync();
                }

                if (!string.IsNullOrEmpty(SelectedModule) && !AvailableModules.Contains(SelectedModule) && AvailableModules.Any())
                {
                    SelectedModule = AvailableModules.First();
                }
                else if (string.IsNullOrEmpty(SelectedModule) && AvailableModules.Any())
                {
                    SelectedModule = AvailableModules.First();
                }

                // 2. Fetch existing user options
                if (!IsNewUser && User.UserID > 0)
                {
                    UserSelectedOptionIds = await PermissionService.GetUserMenuOptionIdsAsync(User.UserID);
                }
                else
                {
                    UserSelectedOptionIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                // 3. Load options for current module
                await LoadModuleOptionsAsync(SelectedModule);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading permissions data: {ex.Message}");
            }
            finally
            {
                IsLoadingOptions = false;
            }
        }

        protected async Task OnModuleChangedAsync(string moduleName)
        {
            SelectedModule = moduleName;
            await LoadModuleOptionsAsync(moduleName);
        }

        private async Task LoadModuleOptionsAsync(string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName)) return;

            try
            {
                IsLoadingOptions = true;
                CurrentModuleOptions = await PermissionService.GetMenuOptionsByModuleAsync(moduleName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading module options: {ex.Message}");
            }
            finally
            {
                IsLoadingOptions = false;
            }
        }

        protected IEnumerable<MenuOptionModel> FilteredModuleOptions
        {
            get
            {
                if (string.IsNullOrWhiteSpace(OptionSearchFilter))
                    return CurrentModuleOptions;
                var q = OptionSearchFilter.Trim().ToLowerInvariant();
                return CurrentModuleOptions.Where(o =>
                    (!string.IsNullOrEmpty(o.OptionName) && o.OptionName.ToLowerInvariant().Contains(q)) ||
                    (!string.IsNullOrEmpty(o.OptionID) && o.OptionID.ToLowerInvariant().Contains(q))
                );
            }
        }

        protected void ToggleOption(string optionId, bool isChecked)
        {
            if (isChecked)
            {
                UserSelectedOptionIds.Add(optionId);
            }
            else
            {
                UserSelectedOptionIds.Remove(optionId);
            }
        }

        protected void SelectAllCurrentModule()
        {
            foreach (var opt in FilteredModuleOptions)
            {
                UserSelectedOptionIds.Add(opt.OptionID);
            }
        }

        protected void ClearAllCurrentModule()
        {
            foreach (var opt in FilteredModuleOptions)
            {
                UserSelectedOptionIds.Remove(opt.OptionID);
            }
        }

        protected int CurrentModuleSelectedCount => CurrentModuleOptions.Count(o => UserSelectedOptionIds.Contains(o.OptionID));
        protected int TotalSelectedOptionsCount => UserSelectedOptionIds.Count;

        private async Task LoadEmployeesAsync()
        {
            try
            {
                if (!_allEmployees.Any())
                {
                    _allEmployees = await EmployeeService.GetEmployeesAsync(string.Empty, true, false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading employees for user binding: {ex.Message}");
            }
        }

        protected async Task<IEnumerable<EmployeeListItemModel>> SearchEmployeesAsync(string searchText)
        {
            await LoadEmployeesAsync();

            if (string.IsNullOrWhiteSpace(searchText))
            {
                return _allEmployees.Take(50);
            }

            var query = searchText.Trim().ToLowerInvariant();
            return _allEmployees.Where(e =>
                (!string.IsNullOrEmpty(e.EmpID) && e.EmpID.ToLowerInvariant().Contains(query)) ||
                (!string.IsNullOrEmpty(e.Name) && e.Name.ToLowerInvariant().Contains(query)) ||
                (!string.IsNullOrEmpty(e.FName) && e.FName.ToLowerInvariant().Contains(query)) ||
                (!string.IsNullOrEmpty(e.Designation) && e.Designation.ToLowerInvariant().Contains(query)) ||
                (!string.IsNullOrEmpty(e.DeptName) && e.DeptName.ToLowerInvariant().Contains(query))
            ).ToList();
        }

        protected async Task SaveUserAsync()
        {
            if (string.IsNullOrWhiteSpace(User.UserName))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Username is required.",
                    Duration = 3000
                });
                return;
            }

            if (IsNewUser && string.IsNullOrWhiteSpace(User.Password))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Password is required for new user accounts.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                IsSaving = true;

                // Bind employee
                User.EmpID = SelectedEmployee?.EmpID;

                if (IsNewUser)
                {
                    var (success, msg, newId) = await UserService.CreateUserAsync(User);
                    if (success)
                    {
                        User.UserID = newId;
                        await PermissionService.SaveAllUserMenuOptionsAsync(newId, UserSelectedOptionIds);
                        await UserRoleDataAccess.SaveUserRolesAsync(newId, SelectedRoles);

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "User Created",
                            Detail = msg,
                            Duration = 3000
                        });
                        NavigateBack();
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
                else
                {
                    var (success, msg) = await UserService.UpdateUserAsync(User);
                    if (success)
                    {
                        await PermissionService.SaveAllUserMenuOptionsAsync(User.UserID, UserSelectedOptionIds);
                        await UserRoleDataAccess.SaveUserRolesAsync(User.UserID, SelectedRoles);

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "User Updated",
                            Detail = msg,
                            Duration = 3000
                        });
                        NavigateBack();
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
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        protected void ToggleRole(string role)
        {
            if (SelectedRoles.Contains(role))
            {
                SelectedRoles.Remove(role);
            }
            else
            {
                SelectedRoles.Add(role);
            }
        }

        protected bool IsRoleSelected(string role) => SelectedRoles.Contains(role);

        protected async Task HandleRolesChanged()
        {
            AvailableRoles = await UserRoleDataAccess.GetAllRolesAsync();
        }

        protected Task CancelAsync()
        {
            NavigateBack();
            return Task.CompletedTask;
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/setup/users");
            }
        }

        protected bool IsCopyModalOpen { get; set; } = false;
        protected string NewCopyUserName { get; set; } = string.Empty;
        protected string NewCopyPassword { get; set; } = string.Empty;
        protected string NewCopyFullName { get; set; } = string.Empty;
        protected bool IsCopying { get; set; } = false;

        protected void OpenCopyModal()
        {
            NewCopyUserName = $"Copy Of {User.UserName}";
            NewCopyPassword = string.Empty;
            NewCopyFullName = User.FullUserName ?? string.Empty;
            IsCopyModalOpen = true;
        }

        protected void CloseCopyModal()
        {
            IsCopyModalOpen = false;
        }

        protected async Task ExecuteCopyAsync()
        {
            if (User.UserID <= 0) return;

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
                IsCopying = true;
                var (success, msg, newId) = await UserService.CopyUserAsync(
                    User.UserID,
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
                    CloseCopyModal();

                    // Open the newly created copy in edit mode
                    NavigationManager.NavigateTo($"/setup/users/edit/{newId}?returnUrl=/setup/users", forceLoad: true);
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
                IsCopying = false;
            }
        }
    }
}
