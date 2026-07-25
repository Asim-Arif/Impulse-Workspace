using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Stock.ChangeBatchNo
{
    public partial class ChangeBatchNo : ComponentBase
    {
        private bool IsLoading = false;
        private IEnumerable<object> AllMaterials = new List<object>();
        private object SelectedMaterial;

        private List<ChangeBatchNoViewModel> GridData = new List<ChangeBatchNoViewModel>();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                AllMaterials = await BatchService.GetMaterialsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Error, "Error", $"Failed to load materials: {ex.Message}");
            }
        }

        private async Task<IEnumerable<object>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return await Task.FromResult(AllMaterials);

            return await Task.FromResult(AllMaterials.Where(m =>
                ((dynamic)m).DisplayName.ToString().Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task LoadDataAsync()
        {
            if (SelectedMaterial == null) return;

            IsLoading = true;
            StateHasChanged();

            try
            {
                string materialId = ((dynamic)SelectedMaterial).RMID1;
                var data = await BatchService.GetMaterialBatchesAsync(materialId);
                GridData = data.ToList();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Error, "Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private async Task OnSaveAsync()
        {
            var modifiedItems = GridData.Where(x => x.IsModified).ToList();
            if (!modifiedItems.Any())
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Info, "Info", "No changes to save.");
                return;
            }

            try
            {
                IsLoading = true;
                StateHasChanged();

                await BatchService.SaveBatchModificationsAsync(modifiedItems);
                
                NotificationService.Notify(Radzen.NotificationSeverity.Success, "Success", "Batch numbers updated successfully!");
                await LoadDataAsync(); // Reload to reset original states
            }
            catch (Exception ex)
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Error, "Error", $"Failed to save changes: {ex.Message}");
                IsLoading = false;
                StateHasChanged();
            }
        }

        private void OnPrint()
        {
            if (SelectedMaterial == null) return;
            
            string materialId = ((dynamic)SelectedMaterial).RMID1;
            
            var request = new ReportRequest
            {
                ReportName = "RM_Lots.rpt",
                SelectionFormula = $"{{Lots_List.MaterialID}}='{materialId}'"
            };

            _ = ReportNavigation.PrintReportAsync(request);
        }
    }
}
