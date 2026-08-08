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

        public string ProcessSearchFilter { get; set; } = string.Empty;
        public bool IsSaving { get; set; } = false;

        public List<RepairTypeProcessItemDto> FilteredProcesses
        {
            get
            {
                if (Model == null || Model.Processes == null)
                    return new List<RepairTypeProcessItemDto>();

                if (string.IsNullOrWhiteSpace(ProcessSearchFilter))
                    return Model.Processes;

                string q = ProcessSearchFilter.Trim();
                return Model.Processes.Where(p =>
                    (!string.IsNullOrEmpty(p.Code) && p.Code.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.Description) && p.Description.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        public void ToggleProcess(RepairTypeProcessItemDto proc, bool isChecked)
        {
            proc.IsSelected = isChecked;
            if (isChecked)
            {
                if (proc.SeqNo <= 0)
                {
                    int maxSeq = Model?.Processes.Where(p => p.IsSelected && p.SeqNo > 0).Select(p => p.SeqNo).DefaultIfEmpty(0).Max() ?? 0;
                    proc.SeqNo = maxSeq + 1;
                }
            }
            else
            {
                proc.SeqNo = 0;
                proc.Scanning = false;
            }
        }

        public void SelectAllProcesses()
        {
            if (Model == null) return;
            int seq = 1;
            foreach (var proc in Model.Processes)
            {
                proc.IsSelected = true;
                if (proc.SeqNo <= 0)
                    proc.SeqNo = seq++;
            }
        }

        public void ClearAllProcesses()
        {
            if (Model == null) return;
            foreach (var proc in Model.Processes)
            {
                proc.IsSelected = false;
                proc.SeqNo = 0;
                proc.Scanning = false;
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

            if (!Model.Processes.Any(p => p.IsSelected))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Processes Selected",
                    Detail = "Please select at least one manufacturing process for this repair type.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                if (Model.EntryID == 0)
                {
                    int newId = await RepairTypeService.SaveRepairTypeAsync(Model);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Repair Type Created",
                            Detail = $"Repair Type '{Model.RepairType}' created successfully.",
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
