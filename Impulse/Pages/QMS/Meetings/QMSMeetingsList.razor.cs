using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.Meetings
{
    public partial class QMSMeetingsList : ComponentBase
    {
        [Inject] private IQMSMeetingDataAccess MeetingDataAccess { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        private List<QMSMeetingListItemModel> meetings = new();
        private List<string> meetingTypes = new();
        private AttachmentModal? attachmentModalRef;
        private QMSMeetingListItemModel? meetingToDelete;

        private bool isLoading = true;
        private string searchText = string.Empty;
        private string selectedType = string.Empty;
        private string selectedPeriodPreset = "All";
        private DateTime? filterDateFrom;
        private DateTime? filterDateTo;

        protected override async Task OnInitializedAsync()
        {
            await LoadMeetingTypes();
            ApplyPeriodPreset("All");
            await LoadMeetings();
        }

        private async Task LoadMeetingTypes()
        {
            try
            {
                meetingTypes = await MeetingDataAccess.GetMeetingTypesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load meeting types: {ex.Message}");
            }
        }

        private async Task LoadMeetings()
        {
            try
            {
                isLoading = true;
                var filter = new QMSMeetingFilterDto
                {
                    SearchText = searchText,
                    MeetingType = selectedType,
                    DateFrom = filterDateFrom,
                    DateTo = filterDateTo
                };

                meetings = await MeetingDataAccess.GetMeetingsAsync(filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load meetings: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private void OnSearchInput(ChangeEventArgs e)
        {
            searchText = e.Value?.ToString() ?? string.Empty;
            _ = LoadMeetings();
        }

        private void OnTypeChanged(ChangeEventArgs e)
        {
            selectedType = e.Value?.ToString() ?? string.Empty;
            _ = LoadMeetings();
        }

        private void OnPeriodPresetChanged(ChangeEventArgs e)
        {
            selectedPeriodPreset = e.Value?.ToString() ?? "All";
            ApplyPeriodPreset(selectedPeriodPreset);
            _ = LoadMeetings();
        }

        private void ApplyPeriodPreset(string preset)
        {
            var today = DateTime.Today;
            switch (preset)
            {
                case "ThisMonth":
                    filterDateFrom = new DateTime(today.Year, today.Month, 1);
                    filterDateTo = filterDateFrom.Value.AddMonths(1).AddDays(-1);
                    break;
                case "Last30Days":
                    filterDateFrom = today.AddDays(-30);
                    filterDateTo = today;
                    break;
                case "ThisYear":
                    filterDateFrom = new DateTime(today.Year, 1, 1);
                    filterDateTo = new DateTime(today.Year, 12, 31);
                    break;
                case "All":
                default:
                    filterDateFrom = null;
                    filterDateTo = null;
                    break;
            }
        }

        private async Task ResetFilters()
        {
            searchText = string.Empty;
            selectedType = string.Empty;
            selectedPeriodPreset = "All";
            filterDateFrom = null;
            filterDateTo = null;
            await LoadMeetings();
        }

        private async Task RefreshData()
        {
            await LoadMeetings();
        }

        private void NavigateToHub()
        {
            NavigationManager.NavigateTo("/qms");
        }

        private void NavigateToCreate()
        {
            NavigationManager.NavigateTo("/qms/meetings/new?returnUrl=/qms/meetings");
        }

        private void NavigateToEdit(int entryId)
        {
            NavigationManager.NavigateTo($"/qms/meetings/{entryId}?returnUrl=/qms/meetings");
        }

        private void OpenAttachmentModal(int entryId, string title)
        {
            attachmentModalRef?.Open(
                entityType: "QMSMeeting",
                entityId: entryId.ToString(),
                secondaryId: null,
                displayName: $"Meeting: {title}"
            );
        }

        private void PromptDeleteMeeting(QMSMeetingListItemModel meeting)
        {
            meetingToDelete = meeting;
        }

        private void CancelDelete()
        {
            meetingToDelete = null;
        }

        private async Task ConfirmDelete()
        {
            if (meetingToDelete == null) return;

            try
            {
                bool success = await MeetingDataAccess.DeleteMeetingAsync(meetingToDelete.EntryID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Meeting deleted successfully.");
                    meetings.Remove(meetingToDelete);
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Warning", "Could not delete meeting.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete meeting: {ex.Message}");
            }
            finally
            {
                meetingToDelete = null;
            }
        }

        private async Task PrintMeeting(int entryId)
        {
            try
            {
                string selectionFormula = $"{{QMS_Meetings.EntryID}} = {entryId}";
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "QMS_Meeting.rpt",
                    SelectionFormula = selectionFormula
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print meeting report: {ex.Message}");
            }
        }

        private void OnContextMenuEdit(QMSMeetingListItemModel? meeting)
        {
            if (meeting != null)
            {
                NavigateToEdit(meeting.EntryID);
            }
        }

        private async Task OnContextMenuPrint(QMSMeetingListItemModel? meeting)
        {
            if (meeting != null)
            {
                await PrintMeeting(meeting.EntryID);
            }
        }

        private void OnContextMenuAttachments(QMSMeetingListItemModel? meeting)
        {
            if (meeting != null)
            {
                OpenAttachmentModal(meeting.EntryID, meeting.Code ?? meeting.Type ?? "");
            }
        }

        private void OnContextMenuDelete(QMSMeetingListItemModel? meeting)
        {
            if (meeting != null)
            {
                PromptDeleteMeeting(meeting);
            }
        }
    }
}

