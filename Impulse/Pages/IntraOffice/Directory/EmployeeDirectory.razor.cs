using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.IntraOffice.Directory
{
    public partial class EmployeeDirectory : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected string SearchTerm { get; set; } = string.Empty;
        protected string SelectedDepartmentId { get; set; } = string.Empty;

        protected List<IntraUserProfile> UsersList { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();

        protected List<IntraUserProfile> FilteredUsers
        {
            get
            {
                var query = UsersList.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SearchTerm))
                {
                    query = query.Where(u =>
                        (u.EmployeeName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (u.FullUserName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        u.UserName.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        (u.Designation?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (u.Email?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false));
                }

                if (!string.IsNullOrWhiteSpace(SelectedDepartmentId))
                {
                    query = query.Where(u => u.DepartmentId == SelectedDepartmentId);
                }

                return query.ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                Departments = await IntraOfficeService.GetDepartmentsAsync();
                UsersList = await IntraOfficeService.GetActiveUsersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }
    }
}
