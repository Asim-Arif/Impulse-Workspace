using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;
using Impulse.Data;

namespace Impulse.Pages.Production.MakerItemAssignment
{
    public partial class MakerItemAssignment
    {
        [Inject]
        public IMakerItemAssignmentService MakerItemAssignmentService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Inject]
        public IReportNavigationService ReportNavigationService { get; set; } = default!;

        [Inject]
        public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

        public List<MakerLookupModel> Makers { get; set; } = new List<MakerLookupModel>();
        
        private MakerLookupModel? _selectedMaker;
        public MakerLookupModel? SelectedMaker
        {
            get => _selectedMaker;
            set
            {
                if (_selectedMaker != value)
                {
                    _selectedMaker = value;
                    _ = OnMakerSelectedAsync(value);
                }
            }
        }

        public List<ProcessLookupModel> Processes { get; set; } = new List<ProcessLookupModel>();
        
        private ProcessLookupModel? _selectedProcess;
        public ProcessLookupModel? SelectedProcess
        {
            get => _selectedProcess;
            set
            {
                if (_selectedProcess != value)
                {
                    _selectedProcess = value;
                    _ = OnProcessSelectedAsync(value);
                }
            }
        }

        public List<UnassignedItemModel> UnassignedItems { get; set; } = new List<UnassignedItemModel>();
        public UnassignedItemModel? SelectedUnassignedItem { get; set; }

        public List<AssignedMakerItemModel> AssignedItems { get; set; } = new List<AssignedMakerItemModel>();
        public string SearchText { get; set; } = string.Empty;

        public bool IsLoadingItems { get; set; } = false;
        public bool IsAssigning { get; set; } = false;

        // Modal Edit State
        public AssignedMakerItemModel? EditingItem { get; set; }
        public decimal EditRateValue { get; set; }
        public string EditRemarksValue { get; set; } = string.Empty;
        public bool IsSavingEdit { get; set; } = false;

        public IEnumerable<AssignedMakerItemModel> FilteredAssignedItems =>
            string.IsNullOrWhiteSpace(SearchText)
                ? AssignedItems
                : AssignedItems.Where(i => (i.ItemID != null && i.ItemID.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                        || (i.ItemName != null && i.ItemName.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                        || (i.Description != null && i.Description.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                        || (i.Remarks != null && i.Remarks.Contains(SearchText, StringComparison.OrdinalIgnoreCase)));

        protected override async Task OnInitializedAsync()
        {
            await LoadActiveMakersAsync();
        }

        private async Task LoadActiveMakersAsync()
        {
            try
            {
                Makers = await MakerItemAssignmentService.GetActiveMakersAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task OnMakerSelectedAsync(MakerLookupModel? maker)
        {
            _selectedProcess = null;
            SelectedUnassignedItem = null;
            Processes.Clear();
            UnassignedItems.Clear();
            AssignedItems.Clear();

            if (maker != null)
            {
                Processes = await MakerItemAssignmentService.GetMakerAssignedProcessesAsync(maker.VendID);
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task OnProcessSelectedAsync(ProcessLookupModel? process)
        {
            SelectedUnassignedItem = null;
            UnassignedItems.Clear();
            AssignedItems.Clear();

            if (SelectedMaker != null && process != null)
            {
                await LoadAssignedItemsAsync();
                await LoadUnassignedItemsAsync();
            }
            await InvokeAsync(StateHasChanged);
        }

        private async Task LoadAssignedItemsAsync()
        {
            if (SelectedMaker == null || SelectedProcess == null) return;
            IsLoadingItems = true;
            try
            {
                AssignedItems = await MakerItemAssignmentService.GetAssignedItemsAsync(SelectedMaker.VendID, SelectedProcess.ProcessID);
            }
            finally
            {
                IsLoadingItems = false;
            }
        }

        private async Task LoadUnassignedItemsAsync()
        {
            if (SelectedMaker == null || SelectedProcess == null) return;
            UnassignedItems = await MakerItemAssignmentService.GetUnassignedItemsAsync(SelectedMaker.VendID, SelectedProcess.ProcessID);
        }

        public Task<IEnumerable<MakerLookupModel>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<MakerLookupModel>>(Makers);

            return Task.FromResult<IEnumerable<MakerLookupModel>>(
                Makers.Where(m => m.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                               || (m.VendID1 != null && m.VendID1.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public Task<IEnumerable<ProcessLookupModel>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<ProcessLookupModel>>(Processes);

            return Task.FromResult<IEnumerable<ProcessLookupModel>>(
                Processes.Where(p => p.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public Task<IEnumerable<UnassignedItemModel>> SearchUnassignedItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<UnassignedItemModel>>(UnassignedItems);

            return Task.FromResult<IEnumerable<UnassignedItemModel>>(
                UnassignedItems.Where(u => u.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                        || u.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public void OnUnassignedItemSelected(UnassignedItemModel? item)
        {
            SelectedUnassignedItem = item;
        }

        public async Task AssignSelectedItem()
        {
            if (SelectedMaker == null || SelectedProcess == null || SelectedUnassignedItem == null) return;

            IsAssigning = true;
            try
            {
                await MakerItemAssignmentService.AssignItemAsync(SelectedMaker.VendID, SelectedProcess.ProcessID, SelectedUnassignedItem.ItemID, 0m);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Item Assigned",
                    Detail = $"Item [{SelectedUnassignedItem.ItemID}] assigned successfully.",
                    Duration = 4000
                });

                SelectedUnassignedItem = null;
                await LoadAssignedItemsAsync();
                await LoadUnassignedItemsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Assignment Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsAssigning = false;
            }
        }

        public void OpenEditModal(AssignedMakerItemModel item)
        {
            EditingItem = item;
            EditRateValue = item.Rate;
            EditRemarksValue = item.Remarks;
        }

        public void CloseEditModal()
        {
            EditingItem = null;
        }

        public async Task SaveEditModal()
        {
            if (EditingItem == null) return;

            IsSavingEdit = true;
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState?.User?.Identity?.Name ?? "System";

                await MakerItemAssignmentService.UpdateAssignedItemRateAndRemarksAsync(
                    EditingItem.EntryID,
                    EditRateValue,
                    EditingItem.Rate,
                    EditRemarksValue,
                    userName);

                EditingItem.Rate = EditRateValue;
                EditingItem.Remarks = EditRemarksValue;

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Success,
                    Summary = "Updated",
                    Detail = $"Item [{EditingItem.ItemID}] rate updated successfully.",
                    Duration = 3500
                });

                CloseEditModal();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSavingEdit = false;
            }
        }

        public async Task ConfirmUnassign(AssignedMakerItemModel item)
        {
            if (SelectedMaker == null || SelectedProcess == null) return;

            // Check if receivable before deleting
            bool isReceivable = await MakerItemAssignmentService.IsItemReceivableAsync(SelectedMaker.VendID, SelectedProcess.ProcessID, item.ItemID);
            if (isReceivable)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Cannot Unassign",
                    Detail = $"Item [{item.ItemID}] cannot be unassigned. It is currently receivable for this maker.",
                    Duration = 5000
                });
                return;
            }

            try
            {
                await MakerItemAssignmentService.UnassignItemAsync(item.EntryID);
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Unassigned",
                    Detail = $"Item [{item.ItemID}] unassigned successfully.",
                    Duration = 4000
                });

                await LoadAssignedItemsAsync();
                await LoadUnassignedItemsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Unassign Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        public async Task PrintReport()
        {
            if (SelectedMaker == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerAssignedItems.rpt",
                SelectionFormula = $"{{VendAssItems.VendID}}={SelectedMaker.VendID}"
            });
        }

        // Revision History Modal State & Handlers
        public AssignedMakerItemModel? HistoryItem { get; set; }
        public List<ItemRevisionHistoryModel> RevisionHistory { get; set; } = new List<ItemRevisionHistoryModel>();
        public bool IsLoadingHistory { get; set; } = false;

        public async Task OpenHistoryModal(AssignedMakerItemModel item)
        {
            HistoryItem = item;
            RevisionHistory.Clear();
            IsLoadingHistory = true;

            try
            {
                RevisionHistory = await MakerItemAssignmentService.GetItemRevisionHistoryAsync(item.EntryID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoadingHistory = false;
            }
        }

        public void CloseHistoryModal()
        {
            HistoryItem = null;
            RevisionHistory.Clear();
        }

        public async Task OnShowHistoryContextClick(ItemClickEventArgs e)
        {
            if (e.Data is AssignedMakerItemModel item)
            {
                await OpenHistoryModal(item);
            }
        }

        public void OnEditContextClick(ItemClickEventArgs e)
        {
            if (e.Data is AssignedMakerItemModel item)
            {
                OpenEditModal(item);
            }
        }

        public async Task OnUnassignContextClick(ItemClickEventArgs e)
        {
            if (e.Data is AssignedMakerItemModel item)
            {
                await ConfirmUnassign(item);
            }
        }
    }
}
