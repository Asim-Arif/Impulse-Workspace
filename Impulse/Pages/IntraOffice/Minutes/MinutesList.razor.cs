using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.IntraOffice.Minutes
{
    public partial class MinutesList : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected string CurrentUserId { get; set; } = string.Empty;

        protected MinuteFilter Filter { get; set; } = new()
        {
            FromDate = DateTime.Today.AddDays(-90),
            ToDate = DateTime.Today.AddDays(1)
        };

        protected List<MinuteApproval> AllMinutes { get; set; } = new();
        protected List<MinuteType> MinuteTypes { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadInitialDataAsync();
        }

        protected async Task LoadInitialDataAsync()
        {
            try
            {
                MinuteTypes = await IntraOfficeService.GetMinuteTypesAsync();
                await ApplyFilterAsync();
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
        }

        protected async Task ApplyFilterAsync()
        {
            IsLoading = true;
            try
            {
                AllMinutes = await IntraOfficeService.GetMinutesAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Filter Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task ResetFilterAsync()
        {
            Filter = new MinuteFilter
            {
                FromDate = DateTime.Today.AddDays(-90),
                ToDate = DateTime.Today.AddDays(1)
            };
            await ApplyFilterAsync();
        }

        protected async Task QuickApproveAsync(int minuteId)
        {
            try
            {
                var success = await IntraOfficeService.UpdateMinuteStatusAsync(
                    minuteId,
                    "Approved",
                    CurrentUserId,
                    "Approved",
                    "Approved via quick action"
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Approved",
                        Detail = "Minute memo marked as Approved.",
                        Duration = 3000
                    });

                    await ApplyFilterAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Approval Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        protected string GetStatusBadgeClass(string status)
        {
            return status switch
            {
                "Approved" => "bg-success",
                "Rejected" => "bg-danger",
                "Closed" => "bg-secondary",
                _ => "bg-warning text-dark"
            };
        }
    }
}
