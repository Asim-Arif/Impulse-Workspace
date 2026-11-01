using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.FixedAssets
{
    public partial class FixedAssetDisposal : ComponentBase
    {
        [Inject] private IFixedAssetProfileDataAccess DataAccess { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        [Parameter] public string? AssetId { get; set; }

        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private FixedAssetDisposalModel model = new();
        private IReadOnlyList<FixedAssetListRow> assetList = Array.Empty<FixedAssetListRow>();
        private bool isBusy;
        private FixedAssetListRow? selectedAsset => assetList.FirstOrDefault(x => x.AssetId == model.AssetId);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isBusy = true;
                assetList = await DataAccess.GetListAsync(null, null, false);
                model = await DataAccess.NewDisposalAsync(AssetId);
                if (string.IsNullOrWhiteSpace(model.DisposalType))
                {
                    model.DisposalType = "0";
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private Task<IEnumerable<FixedAssetListRow>> SearchAssets(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
                return Task.FromResult<IEnumerable<FixedAssetListRow>>(assetList);

            var q = query.Trim();
            var matches = assetList.Where(a =>
                (a.AssetId != null && a.AssetId.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (a.Description != null && a.Description.Contains(q, StringComparison.OrdinalIgnoreCase)));

            return Task.FromResult(matches);
        }

        private async Task OnAssetSelected(FixedAssetListRow? asset)
        {
            var selectedAssetId = asset?.AssetId ?? "";
            model.AssetId = selectedAssetId;
            if (!string.IsNullOrWhiteSpace(selectedAssetId))
            {
                try
                {
                    isBusy = true;
                    var updated = await DataAccess.NewDisposalAsync(selectedAssetId);
                    model.Wdv = updated.Wdv;
                    model.DisposalType = string.IsNullOrWhiteSpace(model.DisposalType) ? "0" : model.DisposalType;
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Error", ex.Message, 5000);
                }
                finally
                {
                    isBusy = false;
                }
            }
        }

        private void OnDisposalDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DisposalDate = dt;
            }
        }

        private async Task<bool> SaveAsync()
        {
            if (string.IsNullOrWhiteSpace(model.AssetId))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select an asset.", 4000);
                return false;
            }

            if (string.IsNullOrWhiteSpace(model.Remarks))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Disposal remarks are required.", 4000);
                return false;
            }

            try
            {
                isBusy = true;
                await DataAccess.SaveDisposalAsync(model);
                NotificationService.Notify(NotificationSeverity.Success, "Disposed", $"Asset '{model.AssetId}' marked as {(model.DisposalType == "1" ? "Disposed" : "Written Off")}.", 4000);
                return true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Disposal Failed", ex.Message, 5000);
                return false;
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task SaveAndNewAsync()
        {
            var ok = await SaveAsync();
            if (ok)
            {
                model = await DataAccess.NewDisposalAsync();
                model.DisposalType = "0";
            }
        }

        private async Task SaveAndCloseAsync()
        {
            var ok = await SaveAsync();
            if (ok)
            {
                Cancel();
            }
        }

        private void Cancel()
        {
            var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/fixedassets/list";
            NavigationManager.NavigateTo(target);
        }
    }
}

