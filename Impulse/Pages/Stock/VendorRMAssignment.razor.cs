using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock
{
    public partial class VendorRMAssignment
    {
        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<NewVendorViewModel> Vendors = new();
        private List<UnassignedRMLookupModel> UnassignedRMs = new();
        private List<VendorRMAssignmentModel> AssignedRMs = new();

        private NewVendorViewModel _selectedVendor;
        private NewVendorViewModel SelectedVendor
        {
            get => _selectedVendor;
            set
            {
                if (_selectedVendor != value)
                {
                    _selectedVendor = value;
                    _ = OnVendorSelected();
                }
            }
        }

        private UnassignedRMLookupModel SelectedUnassignedRM;
        
        private bool isLoading = false;
        private bool isAssigning = false;
        private int? editingRowId = null;

        private bool showHistoryModal = false;
        private VendorRMAssignmentModel currentHistoryItem = null;
        private List<VendorRMAssignmentHistoryModel> historyItems = new();

        // Keep a backup of the item being edited in case of cancel
        private decimal backupRate;
        private string backupRemarks;

        protected override async Task OnInitializedAsync()
        {
            Vendors = await VendorRMAssignmentService.GetVendorsAsync();
        }

        private async Task<IEnumerable<NewVendorViewModel>> SearchVendors(string searchText)
        {
            return await Task.FromResult(Vendors.Where(x => x.VendorName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task<IEnumerable<UnassignedRMLookupModel>> SearchUnassignedRMs(string searchText)
        {
            return await Task.FromResult(UnassignedRMs.Where(x => x.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnVendorSelected()
        {
            SelectedUnassignedRM = null;
            CancelEdit();

            if (SelectedVendor != null)
            {
                isLoading = true;
                StateHasChanged();
                await LoadVendorData();
                isLoading = false;
                StateHasChanged();
            }
            else
            {
                AssignedRMs.Clear();
                UnassignedRMs.Clear();
                StateHasChanged();
            }
        }

        private async Task LoadVendorData()
        {
            if (SelectedVendor != null)
            {
                AssignedRMs = await VendorRMAssignmentService.GetAssignedRMsAsync(SelectedVendor.AccNo);
                UnassignedRMs = await VendorRMAssignmentService.GetUnassignedRMsAsync(SelectedVendor.AccNo);
            }
        }

        private async Task AssignRM()
        {
            if (SelectedVendor == null || SelectedUnassignedRM == null) return;

            try
            {
                isAssigning = true;
                await VendorRMAssignmentService.AssignRMAsync(SelectedVendor.AccNo, SelectedUnassignedRM.RMID);
                NotificationService.ShowSuccess("Success", "Raw Material assigned successfully.");
                
                SelectedUnassignedRM = null;
                await LoadVendorData();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Could not assign RM: {ex.Message}");
            }
            finally
            {
                isAssigning = false;
            }
        }

        private void EditRow(VendorRMAssignmentModel item)
        {
            // Reset any previous edit
            if (editingRowId.HasValue)
            {
                CancelEdit();
            }

            editingRowId = item.EntryID;
            backupRate = item.Rate;
            backupRemarks = item.Remarks;
        }

        private void CancelEdit()
        {
            if (editingRowId.HasValue)
            {
                var item = AssignedRMs.FirstOrDefault(x => x.EntryID == editingRowId);
                if (item != null)
                {
                    item.Rate = backupRate;
                    item.Remarks = backupRemarks;
                }
                editingRowId = null;
            }
        }

        private async Task SaveRow(VendorRMAssignmentModel item)
        {
            try
            {
                await VendorRMAssignmentService.UpdateRMAssignmentAsync(item.EntryID, item.Rate, item.Remarks, "Admin");
                NotificationService.ShowSuccess("Success", "Assignment updated.");
                editingRowId = null;
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Could not update assignment: {ex.Message}");
            }
        }

        private void OnEditRowClick(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is VendorRMAssignmentModel item)
            {
                EditRow(item);
            }
        }

        private async Task OnUnassignClick(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is VendorRMAssignmentModel item)
            {
                try
                {
                    await VendorRMAssignmentService.UnassignRMAsync(item.EntryID);
                    NotificationService.ShowSuccess("Success", "Unassigned successfully.");
                    await LoadVendorData();
                }
                catch (Exception ex)
                {
                    NotificationService.ShowError("Error", ex.Message);
                }
            }
        }

        private async Task OnViewHistoryClick(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is VendorRMAssignmentModel item)
            {
                currentHistoryItem = item;
                historyItems = await VendorRMAssignmentService.GetHistoryAsync(item.EntryID);
                showHistoryModal = true;
            }
        }

        private void CloseHistoryModal()
        {
            showHistoryModal = false;
            currentHistoryItem = null;
            historyItems.Clear();
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo(string.IsNullOrEmpty(ReturnUrl) ? "/stock" : ReturnUrl);
        }
    }
}
