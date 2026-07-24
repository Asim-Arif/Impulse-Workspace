using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock
{
    public partial class RMGroups : ComponentBase
    {
        [Inject] private IRMGroupsService RMGroupsService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IJSRuntime JSRuntime { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;

        private List<RMGroupViewModel> AllGroups = new();
        private List<RMGroupViewModel> FilteredGroups = new();
        private bool isLoading = true;
        private string searchText = string.Empty;

        private bool showModal = false;
        private RMGroupViewModel editingModel = new();
        private string hexColor = "#FFFFFF";

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                AllGroups = await RMGroupsService.GetRMGroupsAsync();
                ApplyFilter();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading RM groups: {ex.Message}");
                await JSRuntime.InvokeVoidAsync("alert", "Error loading data.");
            }
            finally
            {
                isLoading = false;
            }
        }

        private void ApplyFilter()
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                FilteredGroups = AllGroups.ToList();
            }
            else
            {
                FilteredGroups = AllGroups
                    .Where(g => (g.Description ?? "").Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
        }

        private async Task RefreshList()
        {
            await LoadDataAsync();
        }

        private void AddNewGroup()
        {
            editingModel = new RMGroupViewModel { GrpColor = 16777215 }; // Default White
            hexColor = ConvertWin32ColorToHex(editingModel.GrpColor);
            showModal = true;
        }

        private void EditGroup(RMGroupViewModel group)
        {
            editingModel = new RMGroupViewModel
            {
                ID = group.ID,
                Description = group.Description,
                GrpColor = group.GrpColor
            };
            hexColor = ConvertWin32ColorToHex(editingModel.GrpColor);
            showModal = true;
        }

        private void CloseModal()
        {
            showModal = false;
            editingModel = new RMGroupViewModel();
        }

        private async Task SaveGroup()
        {
            if (string.IsNullOrWhiteSpace(editingModel.Description))
            {
                await JSRuntime.InvokeVoidAsync("alert", "Description is required.");
                return;
            }

            try
            {
                editingModel.GrpColor = ConvertHexToWin32Color(hexColor);

                if (editingModel.ID == 0)
                {
                    await RMGroupsService.InsertRMGroupAsync(editingModel);
                }
                else
                {
                    await RMGroupsService.UpdateRMGroupAsync(editingModel);
                }
                
                await LoadDataAsync();
                CloseModal();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving RM group: {ex.Message}");
                await JSRuntime.InvokeVoidAsync("alert", "Error saving data.");
            }
        }

        private async Task ConfirmDelete(RMGroupViewModel group)
        {
            bool confirmed = await JSRuntime.InvokeAsync<bool>("confirm", $"Are you sure you want to delete '{group.Description}'?");
            if (confirmed)
            {
                try
                {
                    var linkedCount = await RMGroupsService.CountLinkedMaterialsAsync(group.ID);
                    if (linkedCount > 0)
                    {
                        await JSRuntime.InvokeVoidAsync("alert", "Can't Delete, Materials are linked.");
                        return;
                    }

                    await RMGroupsService.DeleteRMGroupAsync(group.ID);
                    await LoadDataAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting RM group: {ex.Message}");
                    await JSRuntime.InvokeVoidAsync("alert", "Error deleting group.");
                }
            }
        }

        private async Task OpenManageMenu(MouseEventArgs e, RMGroupViewModel group)
        {
            await ContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY, group);
        }

        private void NavigateToDashboard()
        {
            NavigationManager.NavigateTo("/stock");
        }

        private void OnColorChanged(ChangeEventArgs e)
        {
            hexColor = e.Value?.ToString() ?? "#FFFFFF";
        }

        // Color Conversion Helpers
        private string ConvertWin32ColorToHex(long win32Color)
        {
            long r = win32Color & 0xFF;
            long g = (win32Color >> 8) & 0xFF;
            long b = (win32Color >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        private long ConvertHexToWin32Color(string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor) || !hexColor.StartsWith("#") || hexColor.Length != 7)
                return 16777215;

            try
            {
                long r = Convert.ToInt64(hexColor.Substring(1, 2), 16);
                long g = Convert.ToInt64(hexColor.Substring(3, 2), 16);
                long b = Convert.ToInt64(hexColor.Substring(5, 2), 16);
                return r | (g << 8) | (b << 16);
            }
            catch
            {
                return 16777215;
            }
        }
    }
}
