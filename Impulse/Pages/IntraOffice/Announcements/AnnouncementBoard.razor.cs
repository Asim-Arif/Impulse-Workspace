using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.IntraOffice.Announcements
{
    public partial class AnnouncementBoard : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected bool IsSaving { get; set; } = false;
        protected bool ShowCreateModalState { get; set; } = false;

        protected string CurrentUserId { get; set; } = string.Empty;
        protected string SearchTerm { get; set; } = string.Empty;
        protected string SelectedDepartmentId { get; set; } = string.Empty;
        protected int SelectedPriorityFilter { get; set; } = -1;

        protected List<Announcement> AnnouncementsList { get; set; } = new();
        protected List<(string DeptId, string DeptName)> Departments { get; set; } = new();

        // Create form model
        protected string NewTitle { get; set; } = string.Empty;
        protected string NewContent { get; set; } = string.Empty;
        protected string NewDepartmentId { get; set; } = string.Empty;
        protected int NewPriority { get; set; } = 0;
        protected bool NewIsPinned { get; set; } = false;

        protected List<Announcement> FilteredAnnouncements
        {
            get
            {
                var query = AnnouncementsList.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SearchTerm))
                {
                    query = query.Where(a =>
                        a.Title.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        a.Content.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        (a.CreatorName?.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase) ?? false));
                }

                if (!string.IsNullOrWhiteSpace(SelectedDepartmentId))
                {
                    query = query.Where(a => a.DepartmentId == SelectedDepartmentId || a.DepartmentId == null);
                }

                if (SelectedPriorityFilter >= 0)
                {
                    query = query.Where(a => (int)a.Priority == SelectedPriorityFilter);
                }

                return query.ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadDataAsync();
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                Departments = await IntraOfficeService.GetDepartmentsAsync();
                AnnouncementsList = await IntraOfficeService.GetAnnouncementsAsync(CurrentUserId);
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

        protected void OpenCreateModal()
        {
            NewTitle = string.Empty;
            NewContent = string.Empty;
            NewDepartmentId = string.Empty;
            NewPriority = 0;
            NewIsPinned = false;
            ShowCreateModalState = true;
        }

        protected void CloseCreateModal()
        {
            ShowCreateModalState = false;
        }

        protected async Task SaveAnnouncementAsync()
        {
            if (string.IsNullOrWhiteSpace(NewTitle))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter an announcement title.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(NewContent))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter announcement content.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                var announcement = new Announcement
                {
                    Title = NewTitle.Trim(),
                    Content = NewContent.Trim(),
                    DepartmentId = string.IsNullOrWhiteSpace(NewDepartmentId) ? null : NewDepartmentId,
                    Priority = (AnnouncementPriority)NewPriority,
                    IsPinned = NewIsPinned,
                    CreatedBy = CurrentUserId
                };

                var id = await IntraOfficeService.CreateAnnouncementAsync(announcement);
                if (id > 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Published",
                        Detail = "Announcement published successfully.",
                        Duration = 3000
                    });

                    ShowCreateModalState = false;
                    await LoadDataAsync();
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = "Failed to publish announcement.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        protected async Task AcknowledgeAsync(int announcementId)
        {
            try
            {
                var success = await IntraOfficeService.AcknowledgeAnnouncementAsync(announcementId, CurrentUserId);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Acknowledged",
                        Detail = "You have acknowledged this notice.",
                        Duration = 3000
                    });

                    var ann = AnnouncementsList.FirstOrDefault(a => a.Id == announcementId);
                    if (ann != null)
                    {
                        ann.IsAcknowledgedByCurrentUser = true;
                        ann.AcknowledgmentCount++;
                    }
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

        protected string GetPriorityBadgeClass(AnnouncementPriority priority)
        {
            return priority switch
            {
                AnnouncementPriority.Critical => "bg-danger",
                AnnouncementPriority.Important => "bg-warning text-dark",
                _ => "bg-secondary"
            };
        }

        protected string GetCardBorderClass(AnnouncementPriority priority)
        {
            return priority switch
            {
                AnnouncementPriority.Critical => "border-danger-subtle",
                AnnouncementPriority.Important => "border-warning-subtle",
                _ => "border-light-subtle"
            };
        }
    }
}
