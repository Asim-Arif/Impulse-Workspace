using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Microsoft.AspNetCore.Components;

namespace Impulse.Shared.Components
{
    public record StoreSelectionResult(bool Confirmed, bool AllStores, List<int> SelectedStoreIds);

    public partial class StoreSelectionModal : ComponentBase
    {
        [Inject] private IDBHelperService DbHelper { get; set; } = default!;
        [Inject] private INotificationService NotificationService { get; set; } = default!;

        [Parameter] public string Title { get; set; } = "Select Stores";

        private bool IsVisible { get; set; }
        private bool IsLoadingStores { get; set; }
        private string SearchText { get; set; } = string.Empty;
        private bool SelectAllStores { get; set; } = true;
        private List<StoreLookupModel> AllStores { get; set; } = new();
        private HashSet<int> SelectedIds { get; set; } = new();
        private TaskCompletionSource<StoreSelectionResult>? _tcs;

        private IEnumerable<StoreLookupModel> FilteredStores =>
            string.IsNullOrWhiteSpace(SearchText)
                ? AllStores
                : AllStores.Where(s => s.StoreName.Contains(SearchText.Trim(), StringComparison.OrdinalIgnoreCase));

        /// <summary>
        /// Opens the store selection modal, loads all stores from the Stores table, and awaits user confirmation.
        /// </summary>
        public async Task<StoreSelectionResult> ShowAsync(string? title = null)
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                Title = title;
            }

            SearchText = string.Empty;
            SelectAllStores = true;
            SelectedIds.Clear();
            IsVisible = true;
            _tcs = new TaskCompletionSource<StoreSelectionResult>();
            StateHasChanged();

            IsLoadingStores = true;
            StateHasChanged();
            try
            {
                AllStores = await DbHelper.GetListAsync<StoreLookupModel>("EntryID, StoreName", "Stores", "ORDER BY StoreName");
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Store Load Error", ex.Message);
                AllStores = new List<StoreLookupModel>();
            }
            finally
            {
                IsLoadingStores = false;
                StateHasChanged();
            }

            return await _tcs.Task;
        }

        private void OnAllStoresChanged(ChangeEventArgs e)
        {
            SelectAllStores = e.Value is bool b && b;
            if (SelectAllStores)
            {
                SelectedIds.Clear();
            }
        }

        private void OnStoreCheckedChanged(int storeId, ChangeEventArgs e)
        {
            bool isChecked = e.Value is bool b && b;
            if (isChecked)
            {
                SelectAllStores = false;
                SelectedIds.Add(storeId);
            }
            else
            {
                SelectedIds.Remove(storeId);
            }
        }

        private void CheckAll()
        {
            SelectAllStores = false;
            SelectedIds = AllStores.Select(s => s.EntryID).ToHashSet();
        }

        private void UncheckAll()
        {
            SelectAllStores = false;
            SelectedIds.Clear();
        }

        private void Submit()
        {
            if (!SelectAllStores && !SelectedIds.Any())
            {
                NotificationService.ShowWarning("Selection Required", "Please select at least one store.");
                return;
            }

            IsVisible = false;
            _tcs?.TrySetResult(new StoreSelectionResult(true, SelectAllStores, SelectedIds.OrderBy(id => id).ToList()));
        }

        private void Cancel()
        {
            IsVisible = false;
            _tcs?.TrySetResult(new StoreSelectionResult(false, false, new List<int>()));
        }
    }
}
