using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.WebUtilities;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using DataAccessLibrary.Interface.Setup;
using Impulse.Services.Production;
using Impulse.Services;
using Impulse.Services.WorkflowTasks;
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

        [Inject]
        public IWorkflowTaskEngine WorkflowTaskEngine { get; set; } = default!;

        [Inject]
        public IUserDataAccess UserDataAccess { get; set; } = default!;

        [Inject]
        public IUserRoleDataAccess UserRoleDataAccess { get; set; } = default!;

        public bool IsDirectorOrAdmin { get; set; } = false;
        public string CurrentUserName { get; set; } = "System";

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
        public string RateChangeReason { get; set; } = string.Empty;
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
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                if (authState?.User?.Identity?.IsAuthenticated == true)
                {
                    CurrentUserName = authState.User.Identity.Name ?? "System";
                    var user = authState.User;

                    bool requiresApproval = await WorkflowTaskEngine.IsApprovalRequiredAsync("MakerRateChange", CurrentUserName, user);
                    IsDirectorOrAdmin = !requiresApproval;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking user roles in MakerItemAssignment: {ex.Message}");
            }

            await LoadActiveMakersAsync();

            // Query parameters support (e.g. navigation from Director task alert: ?makerId=...&processId=...)
            try
            {
                var uri = NavigationManager.ToAbsoluteUri(NavigationManager.Uri);
                var query = QueryHelpers.ParseQuery(uri.Query);
                if (query.TryGetValue("makerId", out var qMakerId) && long.TryParse(qMakerId, out long parsedMakerId))
                {
                    var maker = Makers.FirstOrDefault(m => m.VendID == parsedMakerId);
                    if (maker != null)
                    {
                        _selectedMaker = maker;
                        await OnMakerSelectedAsync(maker);
                        if (query.TryGetValue("processId", out var qProcessId) && int.TryParse(qProcessId, out int pid))
                        {
                            var proc = Processes.FirstOrDefault(p => p.ProcessID == pid);
                            if (proc != null)
                            {
                                _selectedProcess = proc;
                                await OnProcessSelectedAsync(proc);
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore query parsing errors
            }
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
            try
            {
                _selectedProcess = null;
                SelectedUnassignedItem = null;
                Processes.Clear();
                UnassignedItems.Clear();
                AssignedItems.Clear();
                await InvokeAsync(StateHasChanged);

                if (maker != null)
                {
                    Processes = await MakerItemAssignmentService.GetMakerAssignedProcessesAsync(maker.VendID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Maker Processes",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task OnProcessSelectedAsync(ProcessLookupModel? process)
        {
            try
            {
                SelectedUnassignedItem = null;
                UnassignedItems.Clear();
                AssignedItems.Clear();
                await InvokeAsync(StateHasChanged);

                if (SelectedMaker != null && process != null)
                {
                    await LoadAssignedItemsAsync();
                    await LoadUnassignedItemsAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Process Data",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoadingItems = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadAssignedItemsAsync()
        {
            if (SelectedMaker == null || SelectedProcess == null) return;
            IsLoadingItems = true;
            await InvokeAsync(StateHasChanged);
            try
            {
                AssignedItems = await MakerItemAssignmentService.GetAssignedItemsAsync(SelectedMaker.VendID, SelectedProcess.ProcessID);
            }
            catch (Exception ex)
            {
                AssignedItems = new List<AssignedMakerItemModel>();
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Assigned Items",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsLoadingItems = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task LoadUnassignedItemsAsync()
        {
            if (SelectedMaker == null || SelectedProcess == null) return;
            try
            {
                UnassignedItems = await MakerItemAssignmentService.GetUnassignedItemsAsync(SelectedMaker.VendID, SelectedProcess.ProcessID);
            }
            catch (Exception ex)
            {
                UnassignedItems = new List<UnassignedItemModel>();
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Unassigned Items",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                await InvokeAsync(StateHasChanged);
            }
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
            RateChangeReason = string.Empty;
        }

        public void CloseEditModal()
        {
            EditingItem = null;
            RateChangeReason = string.Empty;
        }

        public async Task SaveEditModal()
        {
            if (EditingItem == null) return;

            IsSavingEdit = true;
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState?.User?.Identity?.Name ?? CurrentUserName;

                bool isRateChanged = EditRateValue != EditingItem.Rate;

                // Non-director modifying rate -> generate approval task & notification for Director role
                if (!IsDirectorOrAdmin && isRateChanged)
                {
                    if (string.IsNullOrWhiteSpace(RateChangeReason))
                    {
                        NotificationService.Notify(new Radzen.NotificationMessage
                        {
                            Severity = Radzen.NotificationSeverity.Warning,
                            Summary = "Reason Required",
                            Detail = "Please enter a reason for the rate change request.",
                            Duration = 4000
                        });
                        return;
                    }

                    var rateRequest = new MakerItemRateWorkflowRequestDto
                    {
                        EntryID = EditingItem.EntryID,
                        VendID = SelectedMaker?.VendID ?? EditingItem.VendID,
                        MakerName = SelectedMaker?.DisplayText ?? SelectedMaker?.VenderName ?? "Maker",
                        ProcessID = SelectedProcess?.ProcessID ?? EditingItem.ProcessID,
                        ProcessName = SelectedProcess?.Description ?? "Process",
                        ItemID = EditingItem.ItemID,
                        ItemName = EditingItem.ItemName,
                        OldRate = EditingItem.Rate,
                        NewRate = EditRateValue,
                        AssignedUnit = string.IsNullOrEmpty(EditingItem.AssignedUnit) ? "Pcs" : EditingItem.AssignedUnit,
                        Remarks = EditRemarksValue,
                        Reason = RateChangeReason.Trim(),
                        OriginatorUserName = userName,
                        CreatedAt = DateTime.UtcNow
                    };

                    int taskId = await WorkflowTaskEngine.RequestMakerItemRateChangeAsync(rateRequest);

                    // If remarks were also updated, update remarks while keeping original rate intact
                    if (EditRemarksValue != EditingItem.Remarks)
                    {
                        await MakerItemAssignmentService.UpdateAssignedItemRateAndRemarksAsync(
                            EditingItem.EntryID,
                            EditingItem.Rate,
                            EditingItem.Rate,
                            EditRemarksValue,
                            userName);

                        EditingItem.Remarks = EditRemarksValue;
                    }

                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Request Submitted",
                        Detail = $"Rate change request has been submitted to the Director for approval (Task #{taskId}).",
                        Duration = 5000
                    });

                    CloseEditModal();
                    return;
                }

                // Director/Admin direct update OR only remarks changed
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
                    Detail = $"Item [{EditingItem.ItemID}] updated successfully.",
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
