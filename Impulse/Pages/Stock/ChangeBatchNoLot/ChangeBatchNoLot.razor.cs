using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.ChangeBatchNoLot
{
    public partial class ChangeBatchNoLot : ComponentBase
    {
        private bool IsLoadingInitialData = true;
        private bool IsLoadingGrid = false;
        private bool IsSaving = false;

        private List<ItemViewModel> Materials = new();
        private List<LotListViewModel> Lots = new();

        private ItemViewModel _selectedMaterial;
        public ItemViewModel SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                if (_selectedMaterial != value)
                {
                    _selectedMaterial = value;
                    _ = LoadLotsAsync();
                }
            }
        }

        private string _userName;
        private string _machineName;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                _userName = authState.User.Identity.Name;
                _machineName = "WebClient";

                Materials = await ChangeService.GetItemsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading data", ex.Message);
            }
            finally
            {
                IsLoadingInitialData = false;
            }
        }

        private async Task LoadLotsAsync()
        {
            if (SelectedMaterial == null)
            {
                Lots.Clear();
                return;
            }

            IsLoadingGrid = true;
            StateHasChanged();

            try
            {
                Lots = await ChangeService.GetLotsByMaterialAsync(SelectedMaterial.ItemID);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading lots", ex.Message);
            }
            finally
            {
                IsLoadingGrid = false;
                StateHasChanged();
            }
        }

        private async Task<IEnumerable<ItemViewModel>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Materials;
                
            return await Task.FromResult(Materials.Where(x => x.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private void OnFieldChanged(LotListViewModel item, string fieldName, string value)
        {
            bool changed = false;
            switch (fieldName)
            {
                case "Batch_No":
                    if (item.Batch_No != value)
                    {
                        item.Batch_No = value;
                        changed = true;
                    }
                    break;
                case "Mill_Certificate_No":
                    if (item.Mill_Certificate_No != value)
                    {
                        item.Mill_Certificate_No = value;
                        changed = true;
                    }
                    break;
                case "Lot_Remarks":
                    if (item.Lot_Remarks != value)
                    {
                        item.Lot_Remarks = value;
                        changed = true;
                    }
                    break;
            }

            if (changed)
            {
                item.IsModified = true;
            }
        }

        private async Task SaveLotsAsync()
        {
            if (!Lots.Any(l => l.IsModified))
            {
                NotificationService.ShowWarning("No Changes", "There are no modified lots to save.");
                return;
            }

            bool confirm = await NotificationService.ShowQuestionNotification("Confirm Save", "Are you sure you want to save these changes?");
            if (!confirm) return;

            IsSaving = true;
            try
            {
                await ChangeService.UpdateLotsAsync(Lots, _userName, _machineName);
                NotificationService.ShowSuccess("Success", "Lots updated successfully.");
                
                // Reset modified flags
                foreach (var lot in Lots.Where(l => l.IsModified))
                {
                    lot.IsModified = false;
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
