using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.Meetings
{
    public partial class QMSMeetingEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery] public string? returnUrl { get; set; }

        [Inject] private IQMSMeetingDataAccess MeetingDataAccess { get; set; } = default!;
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        private QMSMeetingModel model = new();
        private List<string> meetingTypes = new();
        private List<EmployeeDropdownDto> employees = new();

        private EmployeeDropdownDto? selectedEmployee;
        private string customAttendeeName = string.Empty;

        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isLoading = true;
                await Task.WhenAll(LoadMeetingTypes(), LoadEmployees());

                if (Id > 0)
                {
                    var loaded = await MeetingDataAccess.GetMeetingByIdAsync(Id);
                    if (loaded != null)
                    {
                        model = loaded;
                    }
                    else
                    {
                        NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Meeting #{Id} was not found.");
                        NavigateBack();
                        return;
                    }

                    await LoadAttachmentCount();
                }
                else
                {
                    string nextCode = await MeetingDataAccess.GetNextMeetingCodeAsync();
                    model = new QMSMeetingModel
                    {
                        Code = nextCode,
                        DT = DateTime.Now,
                        Type = meetingTypes.FirstOrDefault() ?? string.Empty
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to initialize form: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
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

        private async Task LoadEmployees()
        {
            try
            {
                employees = await MeetingDataAccess.GetEmployeesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load employees: {ex.Message}");
            }
        }

        private async Task LoadAttachmentCount()
        {
            if (Id > 0)
            {
                try
                {
                    attachmentCount = await AttachmentService.GetCountAsync("QMSMeeting", Id.ToString());
                }
                catch
                {
                    attachmentCount = 0;
                }
            }
        }

        private Task<IEnumerable<EmployeeDropdownDto>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(employees.Take(100));
            }

            var query = searchText.Trim();
            var matches = employees.Where(e =>
                (!string.IsNullOrEmpty(e.EmpID) && e.EmpID.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Name) && e.Name.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.DeptName) && e.DeptName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(query, StringComparison.OrdinalIgnoreCase))
            );

            return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(matches);
        }

        private void OnEmployeeSelected(EmployeeDropdownDto? emp)
        {
            selectedEmployee = emp;
        }

        private void OnDateTimeChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
            }
            else
            {
                model.DT = null;
            }
        }

        private void AddSelectedEmployee()
        {
            if (selectedEmployee == null) return;

            string formattedName = $"{{{selectedEmployee.EmpID}}} {selectedEmployee.Name}";
            if (!string.IsNullOrWhiteSpace(selectedEmployee.Designation))
            {
                formattedName += $" ({selectedEmployee.Designation})";
            }

            if (model.Attendees.Any(a => string.Equals(a, formattedName, StringComparison.OrdinalIgnoreCase) ||
                                         a.StartsWith($"{{{selectedEmployee.EmpID}}}", StringComparison.OrdinalIgnoreCase)))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Already Added", $"Employee [{selectedEmployee.EmpID}] is already in the attendees list.");
                return;
            }

            model.Attendees.Add(formattedName);
            selectedEmployee = null;
        }

        private void AddCustomAttendee()
        {
            if (string.IsNullOrWhiteSpace(customAttendeeName)) return;

            string name = customAttendeeName.Trim();
            if (model.Attendees.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase)))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Already Added", $"Attendee '{name}' is already in the list.");
                return;
            }

            model.Attendees.Add(name);
            customAttendeeName = string.Empty;
        }

        private void HandleCustomAttendeeKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                AddCustomAttendee();
            }
        }

        private void RemoveAttendee(int index)
        {
            if (index >= 0 && index < model.Attendees.Count)
            {
                model.Attendees.RemoveAt(index);
            }
        }

        private void SetActiveTab(string tab)
        {
            activeTab = tab;
        }

        private async Task SaveAndClose()
        {
            await Save(closeAfterSave: true);
        }

        private async Task SaveAndNew()
        {
            await Save(closeAfterSave: false);
        }

        private async Task Save(bool closeAfterSave)
        {
            if (string.IsNullOrWhiteSpace(model.Code))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Meeting Code is required.");
                return;
            }

            if (string.IsNullOrWhiteSpace(model.Type))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Meeting Type is required.");
                return;
            }

            if (!model.DT.HasValue)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Meeting Date & Time is required.");
                return;
            }

            try
            {
                isSaving = true;

                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                string currentUserName = authState.User?.Identity?.Name ?? "Admin";

                int savedId = await MeetingDataAccess.SaveMeetingAsync(model, currentUserName, Environment.MachineName);
                model.EntryID = savedId;
                Id = savedId;

                NotificationService.Notify(NotificationSeverity.Success, "Saved", "Meeting saved successfully.");

                if (closeAfterSave)
                {
                    NavigateBack();
                }
                else
                {
                    NavigationManager.NavigateTo("/qms/meetings/new?returnUrl=" + Uri.EscapeDataString(returnUrl ?? "/qms/meetings"), forceLoad: true);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Save Failed", $"Could not save meeting: {ex.Message}");
            }
            finally
            {
                isSaving = false;
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(returnUrl))
            {
                NavigationManager.NavigateTo(returnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/qms/meetings");
            }
        }
    }
}

