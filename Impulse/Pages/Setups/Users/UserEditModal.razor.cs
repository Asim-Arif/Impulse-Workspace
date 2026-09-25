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
    public partial class UserEditModal : ComponentBase
    {
        [Inject]
        protected IUserService UserService { get; set; } = default!;

        [Inject]
        protected IEmployeeService EmployeeService { get; set; } = default!;

        [Inject]
        protected IUserPermissionService PermissionService { get; set; } = default!;

        [Inject]
        protected NotificationService NotificationService { get; set; } = default!;

        [Parameter]
        public bool IsOpen { get; set; }

        [Parameter]
        public UserModel User { get; set; } = new UserModel();

        [Parameter]
        public bool IsNewUser { get; set; }

        [Parameter]
        public EventCallback<bool> OnClose { get; set; }

        [Parameter]
        public EventCallback<UserModel> OnSaved { get; set; }

        protected int ActiveTab { get; set; } = 0;
        protected bool IsSaving { get; set; } = false;
        protected EmployeeListItemModel? SelectedEmployee { get; set; }
        private List<EmployeeListItemModel> _allEmployees = new();

        // Screen & Menu Options State
        protected List<string> AvailableModules { get; set; } = new();
        protected string SelectedModule { get; set; } = "Accounts";
        protected List<MenuOptionModel> CurrentModuleOptions { get; set; } = new();
        protected HashSet<string> UserSelectedOptionIds { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        protected string OptionSearchFilter { get; set; } = string.Empty;
        protected bool IsLoadingOptions { get; set; } = false;

        private int _lastLoadedUserId = -1;
        private bool _wasOpen = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen)
            {
                if (!_wasOpen || _lastLoadedUserId != User.UserID)
                {
                    _wasOpen = true;
                    _lastLoadedUserId = User.UserID;
                    ActiveTab = 0;
                    OptionSearchFilter = string.Empty;

                    await LoadEmployeesAsync();
                    await LoadPermissionsDataAsync();

                    if (!string.IsNullOrEmpty(User.EmpID))
                    {
                        SelectedEmployee = _allEmployees.FirstOrDefault(e => e.EmpID == User.EmpID);
                    }
                    else
                    {
                        SelectedEmployee = null;
                    }
                }
            }
            else
            {
                _wasOpen = false;
                _lastLoadedUserId = -1;
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
                    // If table was empty, seed with standard catalog
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

                // 3. Load options for initial module
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

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "User Created",
                            Detail = msg,
                            Duration = 3000
                        });
                        await OnSaved.InvokeAsync(User);
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

                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "User Updated",
                            Detail = msg,
                            Duration = 3000
                        });
                        await OnSaved.InvokeAsync(User);
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

        protected async Task CloseAsync()
        {
            await OnClose.InvokeAsync(false);
        }
    }
}
