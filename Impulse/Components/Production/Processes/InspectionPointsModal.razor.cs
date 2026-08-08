using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Components.Production.Processes
{
    public partial class InspectionPointsModal : ComponentBase
    {
        [Inject] public IProcessService ProcessService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public ProcessListItem? Process { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }

        public List<ProcessInspectionPointModel> Points { get; set; } = new();
        public string NewPointDescription { get; set; } = string.Empty;
        public int? EditingEntryId { get; set; }
        public string EditingDescription { get; set; } = string.Empty;

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen && Process != null)
            {
                await LoadPointsAsync();
            }
        }

        public async Task LoadPointsAsync()
        {
            if (Process == null) return;
            IsLoading = true;
            try
            {
                Points = await ProcessService.GetInspectionPointsAsync(Process.ProcessID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Points",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task AddPointAsync()
        {
            if (Process == null || string.IsNullOrWhiteSpace(NewPointDescription))
                return;

            IsSaving = true;
            try
            {
                var point = new ProcessInspectionPointModel
                {
                    ProcessID = Process.ProcessID,
                    Point_Description = NewPointDescription.Trim()
                };

                bool ok = await ProcessService.AddInspectionPointAsync(point);
                if (ok)
                {
                    NewPointDescription = string.Empty;
                    await LoadPointsAsync();
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Checkpoint Added",
                        Detail = "Quality checkpoint added successfully.",
                        Duration = 3000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Adding Point",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task HandleKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter" && !string.IsNullOrWhiteSpace(NewPointDescription))
            {
                await AddPointAsync();
            }
        }

        public void StartEdit(ProcessInspectionPointModel point)
        {
            EditingEntryId = point.EntryID;
            EditingDescription = point.Point_Description;
        }

        public void CancelEdit()
        {
            EditingEntryId = null;
            EditingDescription = string.Empty;
        }

        public async Task SaveEditAsync(ProcessInspectionPointModel point)
        {
            if (string.IsNullOrWhiteSpace(EditingDescription))
                return;

            try
            {
                point.Point_Description = EditingDescription.Trim();
                bool ok = await ProcessService.UpdateInspectionPointAsync(point);
                if (ok)
                {
                    CancelEdit();
                    await LoadPointsAsync();
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Updated",
                        Detail = "Checkpoint updated successfully.",
                        Duration = 3000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task DeletePointAsync(int entryId)
        {
            try
            {
                bool ok = await ProcessService.DeleteInspectionPointAsync(entryId);
                if (ok)
                {
                    await LoadPointsAsync();
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Deleted",
                        Detail = "Checkpoint removed.",
                        Duration = 3000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Delete Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task CloseModal()
        {
            IsOpen = false;
            await OnClosed.InvokeAsync();
        }
    }
}
