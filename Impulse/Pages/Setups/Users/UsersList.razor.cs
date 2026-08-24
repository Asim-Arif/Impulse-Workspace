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
        protected NotificationService NotificationService { get; set; } = default!;

        [Inject]
        protected NavigationManager NavigationManager { get; set; } = default!;

        protected List<UserModel> AllUsers { get; set; } = new();
        protected List<UserModel> FilteredUsers { get; set; } = new();
        protected string SearchQuery { get; set; } = string.Empty;
        protected string StatusFilter { get; set; } = "all";
        protected bool IsLoading { get; set; } = true;

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
            ApplyFilter();
            return Task.CompletedTask;
        }

        protected void HandleSearchKeyUp(KeyboardEventArgs e)
        {
            ApplyFilter();
        }

        protected void OpenCreateModal()
        {
            EditingUser = new UserModel
            {
                InActive = false,
                ChangePassword = true,
                DashBoardMainLink = true,
                OpenCommandCenter = true,
                ShowTips = true
            };
            IsCreatingNew = true;
            IsModalOpen = true;
        }

        protected void OpenEditModal(UserModel user)
        {
            // Clone user to avoid in-place mutations before saving
            EditingUser = new UserModel
            {
                UserID = user.UserID,
                UserName = user.UserName,
                Password = user.Password,
                FullUserName = user.FullUserName,
                EmpID = user.EmpID,
                InActive = user.InActive,
                EmployeeName = user.EmployeeName,
                Designation = user.Designation,
                DepartmentName = user.DepartmentName,
                CellNo = user.CellNo,
                UserManagement = user.UserManagement,
                ChangePassword = user.ChangePassword,
                CompanyMainLink = user.CompanyMainLink,
                FinancialMainLink = user.FinancialMainLink,
                PayrollMainLink = user.PayrollMainLink,
                ExportMainLink = user.ExportMainLink,
                StockMainLink = user.StockMainLink,
                ProductionMainLink = user.ProductionMainLink,
                DashBoardMainLink = user.DashBoardMainLink,
                QMSMainLink = user.QMSMainLink,
                FixedAssetsMainLink = user.FixedAssetsMainLink,
                SamplingMainLink = user.SamplingMainLink,
                HelpMainLink = user.HelpMainLink,
                OpenCommandCenter = user.OpenCommandCenter,
                RestrictedItemProfile = user.RestrictedItemProfile,
                GeneralInfoItemProfile = user.GeneralInfoItemProfile,
                ProcessesItemProfile = user.ProcessesItemProfile,
                WeightItemProfile = user.WeightItemProfile,
                PriceItemProfile = user.PriceItemProfile,
                ReferencesItemProfile = user.ReferencesItemProfile,
                PictureItemProfile = user.PictureItemProfile,
                RMItemProfile = user.RMItemProfile,
                ShipInfoItemProfile = user.ShipInfoItemProfile,
                RestrictCompanyCatalogEditing = user.RestrictCompanyCatalogEditing,
                AuthorizeVouchers = user.AuthorizeVouchers,
                AuthorizeIssuance = user.AuthorizeIssuance,
                PostMakerBill = user.PostMakerBill,
                ChangeRateonIssuance = user.ChangeRateonIssuance,
                ChangeRateMakerAssign = user.ChangeRateMakerAssign,
                MakerBill_EditRate = user.MakerBill_EditRate,
                HideRateMakerAssign = user.HideRateMakerAssign,
                AddEditCustomerComplaint = user.AddEditCustomerComplaint,
                EditFollowUp = user.EditFollowUp,
                CloseCAPA = user.CloseCAPA,
                BackupData = user.BackupData,
                RestoreData = user.RestoreData,
                HicoVisible = user.HicoVisible,
                AddProdPlan = user.AddProdPlan,
                ShowTips = user.ShowTips
            };
            IsCreatingNew = false;
            IsModalOpen = true;
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
    }
}
