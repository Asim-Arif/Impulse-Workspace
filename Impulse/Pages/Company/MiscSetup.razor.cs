using BlazorContextMenu;
using DataAccessLibrary.Models.Company;
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
    public partial class MiscSetup : ComponentBase
    {
        [Parameter]
        public string ConfigKey { get; set; } = string.Empty;

        [Inject] private IMiscSetupService MiscService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;

        private MiscSetupConfig? currentConfig;
        private List<GenericLookupModel> allData = new();
        
        private string searchText = "";
        private bool isLoading = true;
        
        private bool showModal = false;
        private GenericLookupModel editingModel = new();

        private bool IsValidConfig => currentConfig != null;
        private string PageTitle => currentConfig?.PageTitle ?? "Invalid Setup";

        private IEnumerable<GenericLookupModel> FilteredData => 
            string.IsNullOrWhiteSpace(searchText) 
                ? allData 
                : allData.Where(d => d.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase));

        protected override async Task OnParametersSetAsync()
        {
            if (MiscSetupConfig.Configurations.TryGetValue(ConfigKey, out var config))
            {
                currentConfig = config;
                await LoadDataAsync();
            }
            else
            {
                currentConfig = null;
                isLoading = false;
            }
        }

        private async Task LoadDataAsync()
        {
            if (!IsValidConfig) return;
            
            isLoading = true;
            try
            {
                allData = await MiscService.GetLookupDataAsync(currentConfig!);
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
            NotificationService.Notify(NotificationSeverity.Info, "Refreshed", "Data refreshed.");
        }

        private void AddNewRecord()
        {
            editingModel = new GenericLookupModel();
            showModal = true;
        }

        private void EditRecord(GenericLookupModel item)
        {
            editingModel = new GenericLookupModel
            {
                Id = item.Id,
                Description = item.Description
            };
            showModal = true;
        }

        private async Task SaveRecord()
        {
            if (!IsValidConfig) return;

            if (string.IsNullOrWhiteSpace(editingModel.Description))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Description is required.");
                return;
            }

            try
            {
                bool success;
                if (editingModel.Id == 0)
                {
                    success = await MiscService.InsertLookupDataAsync(currentConfig!, editingModel);
                    if (success) NotificationService.Notify(NotificationSeverity.Success, "Saved", "Record added successfully.");
                }
                else
                {
                    success = await MiscService.UpdateLookupDataAsync(currentConfig!, editingModel);
                    if (success) NotificationService.Notify(NotificationSeverity.Success, "Updated", "Record updated successfully.");
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

        private async Task ConfirmDelete(GenericLookupModel item)
        {
            if (!IsValidConfig) return;

            try
            {
                var success = await MiscService.DeleteLookupDataAsync(currentConfig!, item.Id);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Record deleted successfully.");
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

        private async Task OpenManageMenu(MouseEventArgs e, GenericLookupModel item)
        {
            await ContextMenuService.ShowMenu("miscRowContextMenu", (int)e.ClientX, (int)e.ClientY, item);
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }
    }
}
