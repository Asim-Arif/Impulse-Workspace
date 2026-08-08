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

namespace Impulse.Pages.Production.RepairTypes
{
    public partial class RepairTypes : ComponentBase
    {
        [Inject] public IRepairTypeService RepairTypeService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        public List<RepairTypeListItem> RepairTypesList { get; set; } = new();
        public RepairTypeListItem? SelectedRepairType { get; set; }
        public RepairTypeListItem? ContextRowItem { get; set; }

        public string SearchText { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;

        // Modal States
        public bool ShowModal { get; set; } = false;
        public RepairTypeDto? ModelForEdit { get; set; }

        // Delete Modal
        public bool ShowDeleteModal { get; set; } = false;
        public RepairTypeListItem? ItemToDelete { get; set; }
        public bool IsDeleting { get; set; } = false;

        public List<RepairTypeListItem> FilteredRepairTypes
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return RepairTypesList;

                string q = SearchText.Trim();
                return RepairTypesList.Where(r =>
                    (!string.IsNullOrEmpty(r.RepairCode) && r.RepairCode.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(r.RepairType) && r.RepairType.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(r.DefectType) && r.DefectType.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(r.ProcessSummary) && r.ProcessSummary.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadRepairTypesAsync();
        }

        public async Task LoadRepairTypesAsync()
        {
            IsLoading = true;
            try
            {
                RepairTypesList = await RepairTypeService.GetRepairTypesAsync();
                if (SelectedRepairType != null)
                {
                    SelectedRepairType = RepairTypesList.FirstOrDefault(r => r.EntryID == SelectedRepairType.EntryID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Repair Types",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task OpenCreateModal()
        {
            try
            {
                var processes = await RepairTypeService.GetAvailableProcessesForRepairAsync(null);
                ModelForEdit = new RepairTypeDto
                {
                    Processes = processes
                };
                ShowModal = true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Initializing Modal",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OpenEditModal(RepairTypeListItem item)
        {
            if (item == null) return;
            try
            {
                var dto = await RepairTypeService.GetRepairTypeByIdAsync(item.EntryID);
                if (dto != null)
                {
                    ModelForEdit = dto;
                    ShowModal = true;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Repair Type",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnModalSaved()
        {
            ShowModal = false;
            ModelForEdit = null;
            await LoadRepairTypesAsync();
        }

        public void OnModalClosed()
        {
            ShowModal = false;
            ModelForEdit = null;
        }

        public async Task PrintReportAsync()
        {
            string formula = string.Empty;
            if (SelectedRepairType != null)
            {
                formula = $"{{RepairTypes.EntryID}} = {SelectedRepairType.EntryID}";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "RepairTypes.rpt",
                SelectionFormula = formula,
                Parameters = new Dictionary<string, object>()
            });
        }

        public async Task PromptDeleteRepairType(RepairTypeListItem item)
        {
            if (item == null) return;

            var (canDelete, reason) = await RepairTypeService.CanDeleteRepairTypeAsync(item.EntryID);
            if (!canDelete)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete Repair Type",
                    Detail = reason,
                    Duration = 5000
                });
                return;
            }

            ItemToDelete = item;
            ShowDeleteModal = true;
        }

        public async Task ConfirmDeleteRepairType()
        {
            if (ItemToDelete == null) return;

            IsDeleting = true;
            try
            {
                bool ok = await RepairTypeService.DeleteRepairTypeAsync(ItemToDelete.EntryID);
                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Repair Type Deleted",
                        Detail = $"Repair Type '{ItemToDelete.RepairType}' deleted successfully.",
                        Duration = 3000
                    });

                    ShowDeleteModal = false;
                    ItemToDelete = null;
                    await LoadRepairTypesAsync();
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

        // Context Menu Handlers
        public async Task ContextNew(ItemClickEventArgs args) => await OpenCreateModal();

        public async Task ContextEdit(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OpenEditModal(ContextRowItem);
            }
        }

        public async Task ContextPrint(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                SelectedRepairType = ContextRowItem;
                await PrintReportAsync();
            }
        }

        public async Task ContextDelete(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await PromptDeleteRepairType(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is RepairTypeListItem item)
            {
                ContextRowItem = item;
                SelectedRepairType = item;
            }
        }
    }
}
