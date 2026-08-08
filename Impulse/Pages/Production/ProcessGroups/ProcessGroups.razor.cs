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

namespace Impulse.Pages.Production.ProcessGroups
{
    public partial class ProcessGroups : ComponentBase
    {
        [Inject] public IProcessGroupService ProcessGroupService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        public List<ProcessGroupListItem> GroupsList { get; set; } = new();
        public ProcessGroupListItem? SelectedGroup { get; set; }
        public ProcessGroupListItem? ContextRowItem { get; set; }

        public string SearchText { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;

        // Modal States
        public bool ShowGroupModal { get; set; } = false;
        public ProcessGroupDto? GroupForEdit { get; set; }
        public bool IsCopyMode { get; set; } = false;

        // Cascade Modal
        public bool ShowCascadeModal { get; set; } = false;
        public ProcessGroupListItem? GroupForCascade { get; set; }
        public bool IsCascading { get; set; } = false;

        // Delete Modal
        public bool ShowDeleteModal { get; set; } = false;
        public ProcessGroupListItem? GroupToDelete { get; set; }
        public bool IsDeleting { get; set; } = false;

        public List<ProcessGroupListItem> FilteredGroups
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return GroupsList;

                string q = SearchText.Trim();
                return GroupsList.Where(g =>
                    (!string.IsNullOrEmpty(g.GroupName) && g.GroupName.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(g.FlowSummary) && g.FlowSummary.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    g.Children.Any(c =>
                        (!string.IsNullOrEmpty(c.ProcessCode) && c.ProcessCode.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(c.ProcessName) && c.ProcessName.Contains(q, StringComparison.OrdinalIgnoreCase))
                    )
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadGroupsAsync();
        }

        public async Task LoadGroupsAsync()
        {
            IsLoading = true;
            try
            {
                GroupsList = await ProcessGroupService.GetProcessGroupsAsync();
                if (SelectedGroup != null)
                {
                    SelectedGroup = GroupsList.FirstOrDefault(g => g.EntryID == SelectedGroup.EntryID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Groups",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void OpenCreateModal()
        {
            IsCopyMode = false;
            GroupForEdit = new ProcessGroupDto();
            ShowGroupModal = true;
        }

        public async Task OpenEditModal(ProcessGroupListItem g)
        {
            if (g == null) return;
            try
            {
                var dto = await ProcessGroupService.GetProcessGroupByIdAsync(g.EntryID);
                if (dto != null)
                {
                    IsCopyMode = false;
                    GroupForEdit = dto;
                    ShowGroupModal = true;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Group",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OpenCopyModal(ProcessGroupListItem g)
        {
            if (g == null) return;
            try
            {
                var dto = await ProcessGroupService.GetProcessGroupByIdAsync(g.EntryID);
                if (dto != null)
                {
                    IsCopyMode = true;
                    dto.GroupName = $"{dto.GroupName} (Copy)";
                    dto.EntryID = 0;
                    foreach (var c in dto.Processes)
                    {
                        c.EntryID = 0;
                        c.Group_RefID = 0;
                    }

                    GroupForEdit = dto;
                    ShowGroupModal = true;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Duplicating Group",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnGroupModalSaved()
        {
            ShowGroupModal = false;
            GroupForEdit = null;
            await LoadGroupsAsync();
        }

        public void OnGroupModalClosed()
        {
            ShowGroupModal = false;
            GroupForEdit = null;
        }

        public async Task PrintSelectedGroupReport()
        {
            if (SelectedGroup == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "ItemProcessGroups.rpt",
                SelectionFormula = $"{{ProcessGroups.EntryID}} = {SelectedGroup.EntryID}",
                Parameters = new Dictionary<string, object>()
            });
        }

        public void PromptCascadeUpdate(ProcessGroupListItem g)
        {
            GroupForCascade = g;
            ShowCascadeModal = true;
        }

        public async Task ConfirmCascadeUpdate()
        {
            if (GroupForCascade == null) return;

            IsCascading = true;
            try
            {
                bool ok = await ProcessGroupService.UpdateItemProcessesFromProcessGroupAsync(GroupForCascade.EntryID);
                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Linked Items Updated",
                        Detail = $"Successfully cascaded process updates from '{GroupForCascade.GroupName}' to all mapped items.",
                        Duration = 4000
                    });
                    ShowCascadeModal = false;
                    GroupForCascade = null;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Cascade Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsCascading = false;
            }
        }

        public async Task PromptDeleteGroup(ProcessGroupListItem g)
        {
            if (g == null) return;

            var (canDelete, reason) = await ProcessGroupService.CanDeleteProcessGroupAsync(g.EntryID);
            if (!canDelete)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete Group",
                    Detail = reason,
                    Duration = 5000
                });
                return;
            }

            GroupToDelete = g;
            ShowDeleteModal = true;
        }

        public async Task ConfirmDeleteGroup()
        {
            if (GroupToDelete == null) return;

            IsDeleting = true;
            try
            {
                bool ok = await ProcessGroupService.DeleteProcessGroupAsync(GroupToDelete.EntryID);
                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Group Deleted",
                        Detail = $"Process Group '{GroupToDelete.GroupName}' deleted successfully.",
                        Duration = 3000
                    });

                    ShowDeleteModal = false;
                    GroupToDelete = null;
                    await LoadGroupsAsync();
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
        public void ContextNew(ItemClickEventArgs args) => OpenCreateModal();

        public async Task ContextEdit(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OpenEditModal(ContextRowItem);
            }
        }

        public async Task ContextCopy(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OpenCopyModal(ContextRowItem);
            }
        }

        public async Task ContextPrint(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                SelectedGroup = ContextRowItem;
                await PrintSelectedGroupReport();
            }
        }

        public void ContextCascadeUpdate(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                PromptCascadeUpdate(ContextRowItem);
            }
        }

        public async Task ContextDelete(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await PromptDeleteProcessGroup(ContextRowItem);
            }
        }

        private async Task PromptDeleteProcessGroup(ProcessGroupListItem item)
        {
            await PromptDeleteGroup(item);
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is ProcessGroupListItem item)
            {
                ContextRowItem = item;
                SelectedGroup = item;
            }
        }
    }
}
