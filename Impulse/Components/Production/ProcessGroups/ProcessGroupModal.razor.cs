using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Components.Production.ProcessGroups
{
    public partial class ProcessGroupModal : ComponentBase
    {
        [Inject] public IProcessGroupService ProcessGroupService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public ProcessGroupDto? Group { get; set; }
        [Parameter] public bool IsCopy { get; set; } = false;
        [Parameter] public EventCallback OnSaved { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }

        public List<LookupItemInt> AvailableProcesses { get; set; } = new();
        public LookupItemInt? SelectedProcess { get; set; }
        public bool NewStepScanning { get; set; } = false;
        public string NewStepHubName { get; set; } = string.Empty;

        public bool IsSaving { get; set; } = false;

        protected override async Task OnParametersSetAsync()
        {
            if (IsOpen)
            {
                await LoadProcessesLookupAsync();
            }
        }

        public async Task LoadProcessesLookupAsync()
        {
            try
            {
                AvailableProcesses = await ProcessGroupService.GetAvailableProcessesLookupAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Processes",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task<IEnumerable<LookupItemInt>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AvailableProcesses;

            return AvailableProcesses.Where(p => p.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public void OnProcessSelected(LookupItemInt proc)
        {
            SelectedProcess = proc;
        }

        public void AddStep()
        {
            if (Group == null || SelectedProcess == null)
                return;

            string code = string.Empty;
            string name = SelectedProcess.Name;

            int dashIndex = SelectedProcess.Name.IndexOf('-');
            if (dashIndex > 0)
            {
                code = SelectedProcess.Name.Substring(0, dashIndex).Trim();
                name = SelectedProcess.Name.Substring(dashIndex + 1).Trim();
            }

            var step = new ProcessGroupChildDto
            {
                Group_RefID = Group.EntryID,
                Process_RefID = SelectedProcess.Id,
                ProcessCode = code,
                ProcessName = name,
                Scanning = NewStepScanning,
                Hub_Name = NewStepHubName?.Trim() ?? string.Empty,
                SeqNo = Group.Processes.Count + 1
            };

            Group.Processes.Add(step);
            SelectedProcess = null;
            NewStepScanning = false;
            NewStepHubName = string.Empty;
        }

        public void MoveStepUp(int index)
        {
            if (Group == null || index <= 0 || index >= Group.Processes.Count)
                return;

            var temp = Group.Processes[index];
            Group.Processes[index] = Group.Processes[index - 1];
            Group.Processes[index - 1] = temp;

            Resequence();
        }

        public void MoveStepDown(int index)
        {
            if (Group == null || index < 0 || index >= Group.Processes.Count - 1)
                return;

            var temp = Group.Processes[index];
            Group.Processes[index] = Group.Processes[index + 1];
            Group.Processes[index + 1] = temp;

            Resequence();
        }

        public void RemoveStep(int index)
        {
            if (Group == null || index < 0 || index >= Group.Processes.Count)
                return;

            Group.Processes.RemoveAt(index);
            Resequence();
        }

        private void Resequence()
        {
            if (Group == null) return;
            for (int i = 0; i < Group.Processes.Count; i++)
            {
                Group.Processes[i].SeqNo = i + 1;
            }
        }

        public async Task SaveGroupAsync()
        {
            if (Group == null) return;

            if (string.IsNullOrWhiteSpace(Group.GroupName))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Group Name Required",
                    Detail = "Please provide a valid process group name.",
                    Duration = 3000
                });
                return;
            }

            if (!Group.Processes.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Processes",
                    Detail = "Please add at least one process step into the sequence.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                Resequence();

                if (IsCopy || Group.EntryID == 0)
                {
                    Group.EntryID = 0;
                    int newId = await ProcessGroupService.SaveProcessGroupAsync(Group);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Group Created",
                            Detail = $"Process Group '{Group.GroupName}' created successfully with {Group.Processes.Count} steps.",
                            Duration = 4000
                        });

                        IsOpen = false;
                        await OnSaved.InvokeAsync();
                    }
                }
                else
                {
                    bool ok = await ProcessGroupService.UpdateProcessGroupAsync(Group);
                    if (ok)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Group Updated",
                            Detail = $"Process Group '{Group.GroupName}' updated successfully.",
                            Duration = 4000
                        });

                        IsOpen = false;
                        await OnSaved.InvokeAsync();
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
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task CloseModal()
        {
            IsOpen = false;
            await OnClosed.InvokeAsync();
        }
    }
}
