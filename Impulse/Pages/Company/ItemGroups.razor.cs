using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Company;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Company
{
    public partial class ItemGroups : ComponentBase
    {
        [Inject] private IItemGroupsService ItemGroupsService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;

        private List<ItemGroupsViewModel> allGroups = new();
        private List<PurchaserLookupModel> purchasers = new();
        
        private string searchText = "";
        private bool isLoading = true;
        
        private bool showModal = false;
        private ItemGroupsViewModel editingModel = new();
        private string hexColor = "#000000";
        private PurchaserLookupModel? selectedPurchaser;

        private IEnumerable<ItemGroupsViewModel> FilteredGroups => 
            string.IsNullOrWhiteSpace(searchText) 
                ? allGroups 
                : allGroups.Where(g => g.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                                       g.PurchaserName.Contains(searchText, StringComparison.OrdinalIgnoreCase));

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                var groupsTask = ItemGroupsService.GetItemGroupsAsync();
                var purchasersTask = ItemGroupsService.GetPurchasersAsync();
                
                await Task.WhenAll(groupsTask, purchasersTask);
                
                allGroups = await groupsTask;
                purchasers = await purchasersTask;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshList()
        {
            await LoadDataAsync();
            NotificationService.Notify(NotificationSeverity.Info, "Refreshed", "Item Groups data refreshed.");
        }

        private void AddNewItemGroup()
        {
            editingModel = new ItemGroupsViewModel();
            hexColor = "#FFFFFF";
            selectedPurchaser = null;
            showModal = true;
        }

        private void EditItemGroup(ItemGroupsViewModel group)
        {
            editingModel = new ItemGroupsViewModel
            {
                ID = group.ID,
                Description = group.Description,
                GrpColor = group.GrpColor,
                EmpID_Purchaser = group.EmpID_Purchaser,
                PurchaserName = group.PurchaserName
            };
            hexColor = ConvertWin32ColorToHex(group.GrpColor);
            selectedPurchaser = purchasers.FirstOrDefault(p => p.EmpID == group.EmpID_Purchaser);
            showModal = true;
        }

        private async Task SaveItemGroup()
        {
            if (string.IsNullOrWhiteSpace(editingModel.Description))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Description is required.");
                return;
            }

            try
            {
                editingModel.GrpColor = ConvertHexToWin32Color(hexColor);
                editingModel.EmpID_Purchaser = selectedPurchaser?.EmpID ?? string.Empty;

                bool success;
                if (editingModel.ID == 0)
                {
                    success = await ItemGroupsService.InsertItemGroupAsync(editingModel);
                    if (success) NotificationService.Notify(NotificationSeverity.Success, "Saved", "Item Group added successfully.");
                }
                else
                {
                    success = await ItemGroupsService.UpdateItemGroupAsync(editingModel);
                    if (success) NotificationService.Notify(NotificationSeverity.Success, "Updated", "Item Group updated successfully.");
                }

                if (success)
                {
                    showModal = false;
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save: {ex.Message}");
            }
        }

        private async Task ConfirmDelete(ItemGroupsViewModel group)
        {
            // Ideally use a confirmation dialog here, simple implementation for now
            try
            {
                var success = await ItemGroupsService.DeleteItemGroupAsync(group.ID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Item Group deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete: {ex.Message}");
            }
        }

        private void CloseModal()
        {
            showModal = false;
        }

        private async Task OpenManageMenu(MouseEventArgs e, ItemGroupsViewModel group)
        {
            // Pass the model into the context menu data so we know which row was clicked
            await ContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY, group);
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }

        private Task<IEnumerable<PurchaserLookupModel>> SearchPurchasers(string searchText)
        {
            var result = string.IsNullOrWhiteSpace(searchText) 
                ? purchasers 
                : purchasers.Where(p => p.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                                        p.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult(result);
        }

        private void OnPurchaserSelected(PurchaserLookupModel purchaser)
        {
            selectedPurchaser = purchaser;
        }

        private void OnColorChanged(ChangeEventArgs e)
        {
            if (e.Value is string newHex)
            {
                hexColor = newHex;
            }
        }

        private string ConvertWin32ColorToHex(long colorValue)
        {
            var b = (colorValue >> 16) & 0xFF;
            var g = (colorValue >> 8) & 0xFF;
            var r = colorValue & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        private long ConvertHexToWin32Color(string hex)
        {
            if (string.IsNullOrEmpty(hex) || !hex.StartsWith("#") || hex.Length != 7) return 0;
            
            try 
            {
                var r = Convert.ToInt64(hex.Substring(1, 2), 16);
                var g = Convert.ToInt64(hex.Substring(3, 2), 16);
                var b = Convert.ToInt64(hex.Substring(5, 2), 16);
                return (b << 16) | (g << 8) | r;
            }
            catch 
            {
                return 0; 
            }
        }
    }
}
