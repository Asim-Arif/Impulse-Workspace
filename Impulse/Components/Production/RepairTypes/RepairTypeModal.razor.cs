using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Components.Production.RepairTypes
{
    public partial class RepairTypeModal : ComponentBase
    {
        [Inject] public IRepairTypeService RepairTypeService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public RepairTypeDto? Model { get; set; }
        [Parameter] public EventCallback OnSaved { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }

        public List<RepairTypeProcessItemDto> AvailableProcesses { get; set; } = new();
        public RepairTypeProcessItemDto? SelectedProcess { get; set; }
        public bool NewStepScanning { get; set; } = false;
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
                AvailableProcesses = await RepairTypeService.GetAvailableProcessesForRepairAsync(null);
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

        public async Task<IEnumerable<RepairTypeProcessItemDto>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AvailableProcesses;

            string q = searchText.Trim();
            return AvailableProcesses.Where(p =>
                (!string.IsNullOrEmpty(p.Code) && p.Code.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(p.Description) && p.Description.Contains(q, StringComparison.OrdinalIgnoreCase))
            );
        }

        public void OnProcessSelected(RepairTypeProcessItemDto proc)
        {
            SelectedProcess = proc;
        }

        public void AddStep()
        {
            if (Model == null || SelectedProcess == null)
                return;

            if (Model.Processes.Any(p => p.ProcessID == SelectedProcess.ProcessID))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Duplicate Process",
                    Detail = $"Process '{SelectedProcess.Description}' is already in this repair sequence.",
                    Duration = 3500
                });
                return;
            }

            var step = new RepairTypeProcessItemDto
            {
                ProcessID = SelectedProcess.ProcessID,
                Code = SelectedProcess.Code,
                Description = SelectedProcess.Description,
                IsSelected = true,
                Scanning = NewStepScanning,
                SeqNo = Model.Processes.Count + 1
            };

            Model.Processes.Add(step);
            SelectedProcess = null;
            NewStepScanning = false;
        }

        public void MoveStepUp(int index)
        {
            if (Model == null || index <= 0 || index >= Model.Processes.Count)
                return;

            var temp = Model.Processes[index];
            Model.Processes[index] = Model.Processes[index - 1];
            Model.Processes[index - 1] = temp;

            Resequence();
        }

        public void MoveStepDown(int index)
        {
            if (Model == null || index < 0 || index >= Model.Processes.Count - 1)
                return;

            var temp = Model.Processes[index];
            Model.Processes[index] = Model.Processes[index + 1];
            Model.Processes[index + 1] = temp;

            Resequence();
        }

        public void RemoveStep(int index)
        {
            if (Model == null || index < 0 || index >= Model.Processes.Count)
                return;

            Model.Processes.RemoveAt(index);
            Resequence();
        }

        private void Resequence()
        {
            if (Model == null) return;
            for (int i = 0; i < Model.Processes.Count; i++)
            {
                Model.Processes[i].SeqNo = i + 1;
                Model.Processes[i].IsSelected = true;
            }
        }

        public async Task SaveRepairTypeAsync()
        {
            if (Model == null) return;

            if (string.IsNullOrWhiteSpace(Model.RepairCode))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Repair Code Required",
                    Detail = "Please provide a valid repair code.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.RepairType))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Repair Type Description Required",
                    Detail = "Please enter the repair type description.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.DefectType))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Defect Classification Required",
                    Detail = "Please enter the defect classification category.",
                    Duration = 3000
                });
                return;
            }

            if (!Model.Processes.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Processes Added",
                    Detail = "Please add at least one manufacturing process into the sequence.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                Resequence();

                if (Model.EntryID == 0)
                {
                    int newId = await RepairTypeService.SaveRepairTypeAsync(Model);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Repair Type Created",
                            Detail = $"Repair Type '{Model.RepairType}' created successfully with {Model.Processes.Count} steps.",
                            Duration = 4000
                        });

                        IsOpen = false;
                        await OnSaved.InvokeAsync();
                    }
                }
                else
                {
                    bool ok = await RepairTypeService.UpdateRepairTypeAsync(Model);
                    if (ok)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Repair Type Updated",
                            Detail = $"Repair Type '{Model.RepairType}' updated successfully.",
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
