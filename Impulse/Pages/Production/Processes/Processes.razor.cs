using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.Processes
{
    public partial class Processes : ComponentBase
    {
        [Inject] public IProcessService ProcessService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        public List<ProcessListItem> ProcessesList { get; set; } = new();
        public ProcessLookups Lookups { get; set; } = new();

        public ProcessModel CurrentProcess { get; set; } = new();
        public bool IsEditing { get; set; } = false;
        public bool IsSaving { get; set; } = false;
        public bool IsLoading { get; set; } = false;

        public string SearchText { get; set; } = string.Empty;
        public bool ShowUrduDescription { get; set; } = false;

        // Lookup selections
        public LookupItemInt? SelectedMaker { get; set; }
        public LookupItemInt? SelectedInspectionProcess { get; set; }
        public LookupItemInt? SelectedBillingProcess { get; set; }

        // Inspection Points Modal
        public bool ShowInspectionModal { get; set; } = false;
        public ProcessListItem? SelectedProcessForInspection { get; set; }

        // Delete Modal
        public bool ShowDeleteModal { get; set; } = false;
        public ProcessListItem? ProcessToDelete { get; set; }
        public bool IsDeleting { get; set; } = false;

        public ProcessListItem? ContextRowItem { get; set; }

        public List<ProcessListItem> FilteredProcesses
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return ProcessesList;

                string q = SearchText.Trim();
                return ProcessesList.Where(p =>
                    (!string.IsNullOrEmpty(p.Code) && p.Code.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.Description) && p.Description.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProcessNameUrdu) && p.ProcessNameUrdu.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.Supervisor) && p.Supervisor.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.VenderName) && p.VenderName.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadLookupsAsync();
            await LoadProcessesAsync();
        }

        public async Task LoadLookupsAsync()
        {
            try
            {
                Lookups = await ProcessService.GetLookupsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Lookups",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task LoadProcessesAsync()
        {
            IsLoading = true;
            try
            {
                ProcessesList = await ProcessService.GetProcessesAsync(isInspectionProcess: false);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Processes",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Lookups.Makers;

            return Lookups.Makers.Where(m => m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public async Task<IEnumerable<LookupItemInt>> SearchInspectionProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Lookups.InspectionProcesses;

            return Lookups.InspectionProcesses.Where(i => i.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public async Task<IEnumerable<LookupItemInt>> SearchBillingProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Lookups.BillingProcesses;

            return Lookups.BillingProcesses.Where(b => b.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public void OnOperationChanged()
        {
            if (CurrentProcess.Operation == 0) // Employee
            {
                CurrentProcess.Fix_Maker_RefID = null;
                SelectedMaker = null;
            }
        }

        public void OnAuthRequiredChanged()
        {
            if (CurrentProcess.AuthRequired == 0)
            {
                CurrentProcess.Insp_RefID = null;
                SelectedInspectionProcess = null;
            }
        }

        public void OnMakerSelected(LookupItemInt maker)
        {
            SelectedMaker = maker;
            CurrentProcess.Fix_Maker_RefID = maker?.Id;
        }

        public void OnInspectionProcessSelected(LookupItemInt insp)
        {
            SelectedInspectionProcess = insp;
            CurrentProcess.Insp_RefID = insp?.Id;
        }

        public void OnBillingProcessSelected(LookupItemInt bp)
        {
            SelectedBillingProcess = bp;
            CurrentProcess.BillingProcessID = bp?.Id;
        }

        public void EditProcess(ProcessListItem proc)
        {
            if (proc == null) return;

            IsEditing = true;
            CurrentProcess = new ProcessModel
            {
                ProcessID = proc.ProcessID,
                SNo = proc.SNO,
                Code = proc.Code,
                Description = proc.Description,
                Supervisor = proc.Supervisor,
                Operation = proc.Operation,
                AuthRequired = proc.AuthRequired ?? 0,
                ProcessNameUrdu = proc.ProcessNameUrdu,
                ProcessNameUrduOther = proc.ProcessNameUrduOther,
                Insp_RefID = proc.Insp_RefID,
                Fix_Maker_RefID = proc.Fix_Maker_RefID,
                InspectionProcess = proc.InspectionProcess ?? false,
                BillingProcessID = proc.BillingProcessID
            };

            SelectedMaker = proc.Fix_Maker_RefID.HasValue
                ? Lookups.Makers.FirstOrDefault(m => m.Id == proc.Fix_Maker_RefID.Value)
                : null;

            SelectedInspectionProcess = proc.Insp_RefID.HasValue
                ? Lookups.InspectionProcesses.FirstOrDefault(i => i.Id == proc.Insp_RefID.Value)
                : null;

            SelectedBillingProcess = proc.BillingProcessID.HasValue
                ? Lookups.BillingProcesses.FirstOrDefault(b => b.Id == proc.BillingProcessID.Value)
                : null;
        }

        public void ResetForm()
        {
            IsEditing = false;
            CurrentProcess = new ProcessModel();
            SelectedMaker = null;
            SelectedInspectionProcess = null;
            SelectedBillingProcess = null;
        }

        public async Task SaveProcessAsync()
        {
            if (string.IsNullOrWhiteSpace(CurrentProcess.Code))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Process Code Required",
                    Detail = "Please enter a valid process code (e.g., CUT, STITCH).",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(CurrentProcess.Description))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Process Name Required",
                    Detail = "Please enter a primary description/name for the process.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                if (IsEditing)
                {
                    bool ok = await ProcessService.UpdateProcessAsync(CurrentProcess);
                    if (ok)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Process Updated",
                            Detail = $"Process '{CurrentProcess.Description}' updated successfully.",
                            Duration = 3000
                        });
                        ResetForm();
                        await LoadProcessesAsync();
                    }
                }
                else
                {
                    int newId = await ProcessService.SaveProcessAsync(CurrentProcess);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Process Created",
                            Detail = $"Process '{CurrentProcess.Description}' created successfully.",
                            Duration = 3000
                        });
                        ResetForm();
                        await LoadProcessesAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Operation Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public void OpenInspectionPoints(ProcessListItem proc)
        {
            SelectedProcessForInspection = proc;
            ShowInspectionModal = true;
        }

        public void OnInspectionModalClosed()
        {
            ShowInspectionModal = false;
            SelectedProcessForInspection = null;
        }

        public async Task PromptDeleteProcess(ProcessListItem proc)
        {
            if (proc == null) return;

            var (canDelete, reason) = await ProcessService.CanDeleteProcessAsync(proc.ProcessID);
            if (!canDelete)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete Process",
                    Detail = reason,
                    Duration = 5000
                });
                return;
            }

            ProcessToDelete = proc;
            ShowDeleteModal = true;
        }

        public async Task ConfirmDeleteProcess()
        {
            if (ProcessToDelete == null) return;

            IsDeleting = true;
            try
            {
                bool ok = await ProcessService.DeleteProcessAsync(ProcessToDelete.ProcessID);
                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Process Deleted",
                        Detail = $"Process '{ProcessToDelete.Description}' deleted successfully.",
                        Duration = 3000
                    });

                    if (CurrentProcess.ProcessID == ProcessToDelete.ProcessID)
                    {
                        ResetForm();
                    }

                    ShowDeleteModal = false;
                    ProcessToDelete = null;
                    await LoadProcessesAsync();
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
            finally
            {
                IsDeleting = false;
            }
        }

        public async Task PrintProcessReport()
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "Processes.rpt",
                SelectionFormula = "{VProcesses.InspectionProcess} = false",
                Parameters = new Dictionary<string, object>()
            });
        }

        public void ContextEdit(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                EditProcess(ContextRowItem);
            }
        }

        public void ContextInspectionPoints(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                OpenInspectionPoints(ContextRowItem);
            }
        }

        public async Task ContextDelete(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await PromptDeleteProcess(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is ProcessListItem item)
            {
                ContextRowItem = item;
            }
        }
    }
}
