using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using BlazorContextMenu;

namespace Impulse.Pages.Stock.MaterialPlacementList
{
    public partial class MaterialPlacementList : ComponentBase
    {
        private List<PlacedMaterialItemViewModel> PlacedMaterials;
        private bool IsLoading = false;
        
        public bool ShowPicture { get; set; }
        public string SelectedMaterialId { get; set; }

        public class ListFilter
        {
            public int DateRangeIndex { get; set; } = 1; // Default to Last 15 Days
            public DateTime DateFrom { get; set; } = DateTime.Today.AddDays(-15);
            public DateTime DateTo { get; set; } = DateTime.Today;
        }

        private ListFilter Filter = new ListFilter();
        private object SelectedVendor;
        private object SelectedMaterial;
        
        private IEnumerable<object> AllVendors = new List<object>();
        private IEnumerable<object> AllMaterials = new List<object>();
        
        private PlacedMaterialItemViewModel contextMenuRowItem;

        protected override async Task OnInitializedAsync()
        {
            await LoadDropdownsAsync();
            await LoadDataAsync();
        }
        
        private async Task LoadDropdownsAsync()
        {
            AllVendors = await PlacementListService.GetVendorsAsync();
            AllMaterials = await PlacementListService.GetMaterialsAsync();
        }

        private async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();

            string vendId = "0";
            if (SelectedVendor != null)
            {
                vendId = ((dynamic)SelectedVendor).AccNo;
            }

            string matId = "0";
            if (SelectedMaterial != null)
            {
                matId = ((dynamic)SelectedMaterial).RMID1;
            }

            try
            {
                var result = await PlacementListService.GetPlacedMaterialsListAsync(
                    Filter.DateFrom,
                    Filter.DateTo,
                    vendId,
                    matId
                );
                PlacedMaterials = result.ToList();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage { Severity = Radzen.NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private Task<IEnumerable<object>> SearchVendors(string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
            {
                return Task.FromResult(AllVendors);
            }
            var filtered = AllVendors.Where(v => ((dynamic)v).AccTitle.ToLower().Contains(searchText.ToLower())).ToList();
            return Task.FromResult((IEnumerable<object>)filtered);
        }

        private Task<IEnumerable<object>> SearchMaterials(string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
            {
                return Task.FromResult(AllMaterials);
            }
            var filtered = AllMaterials.Where(m => ((dynamic)m).MaterialName.ToLower().Contains(searchText.ToLower())).ToList();
            return Task.FromResult((IEnumerable<object>)filtered);
        }

        private async Task OnDateRangeChanged()
        {
            DateTime serverDate = DateTime.Today;
            switch (Filter.DateRangeIndex)
            {
                case 0: // Today
                    Filter.DateFrom = serverDate;
                    Filter.DateTo = serverDate;
                    break;
                case 1: // Last 15 Days
                    Filter.DateFrom = serverDate.AddDays(-15);
                    Filter.DateTo = serverDate;
                    break;
                case 2: // Last 30 Days
                    Filter.DateFrom = serverDate.AddDays(-30);
                    Filter.DateTo = serverDate;
                    break;
                case 3: // Last 60 Days
                    Filter.DateFrom = serverDate.AddDays(-60);
                    Filter.DateTo = serverDate;
                    break;
                case 4: // Last 90 Days
                    Filter.DateFrom = serverDate.AddDays(-90);
                    Filter.DateTo = serverDate;
                    break;
                case 5: // Custom
                    break;
            }

            if (Filter.DateRangeIndex != 5)
            {
                await LoadDataAsync();
            }
        }
        
        private void SelectRowMaterial(string materialId)
        {
            SelectedMaterialId = materialId;
        }

        private async Task OpenRowContextMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e, PlacedMaterialItemViewModel item)
        {
            contextMenuRowItem = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task OnPrintSlip(ItemClickEventArgs e)
        {
            if (contextMenuRowItem != null)
            {
                var req = new ReportRequest
                {
                    ReportName = "MaterialPlacement_Slip.rpt",
                    SelectionFormula = $"{{VMaterialLocationWiseStatus.EntryID}}={contextMenuRowItem.EntryID}"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
        }
    }
}
