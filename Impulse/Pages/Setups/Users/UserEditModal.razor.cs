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

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen)
            {
                ActiveTab = 0;
                await LoadEmployeesAsync();

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
