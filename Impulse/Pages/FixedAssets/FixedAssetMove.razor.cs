using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.FixedAssets
{
    public partial class FixedAssetMove : ComponentBase
    {
        [Inject] private IFixedAssetProfileDataAccess DataAccess { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = null!;

        [Parameter] public string? AssetId { get; set; }

        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private FixedAssetMoveModel model = new();
        private FixedAssetLookupOptions lookup = new();
        private IReadOnlyList<FixedAssetListRow> allAssets = Array.Empty<FixedAssetListRow>();
        private string? selectedHeadId;
        private bool isBusy;
        private FixedAssetListRow? selectedAsset => allAssets.FirstOrDefault(x => x.AssetId == model.AssetId);

        private IEnumerable<FixedAssetListRow> filteredAssets =>
            string.IsNullOrWhiteSpace(selectedHeadId)
                ? allAssets
                : allAssets.Where(x => x.Category == lookup.Masters.FirstOrDefault(m => m.Id == selectedHeadId)?.Name);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isBusy = true;
                lookup = await DataAccess.GetLookupOptionsAsync();
                allAssets = await DataAccess.GetListAsync(null, null, false);
                model = await DataAccess.NewMoveAsync(AssetId);
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
            var source = filteredAssets;
            if (string.IsNullOrWhiteSpace(query))
                return Task.FromResult(source);

            var q = query.Trim();
            var matches = source.Where(a =>
                (a.AssetId != null && a.AssetId.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (a.Description != null && a.Description.Contains(q, StringComparison.OrdinalIgnoreCase)));

            return Task.FromResult(matches);
        }

        private void OnHeadFilterChanged(ChangeEventArgs e)
        {
            selectedHeadId = e.Value?.ToString();
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
                    var updated = await DataAccess.NewMoveAsync(selectedAssetId);
                    model.OriginalHeadName = updated.OriginalHeadName;
                    model.CurrentHeadId = updated.CurrentHeadId;
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

        private async Task SaveAsync()
        {
            if (string.IsNullOrWhiteSpace(model.AssetId))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select an asset to move.", 4000);
                return;
            }

            if (!model.NewHeadId.HasValue || model.NewHeadId.Value <= 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a destination head.", 4000);
                return;
            }

            try
            {
                isBusy = true;
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var userName = authState.User.Identity?.Name ?? authState.User.FindFirstValue(ClaimTypes.Name) ?? "Administrator";

                await DataAccess.SaveMoveAsync(model, userName, Environment.MachineName);
                NotificationService.Notify(NotificationSeverity.Success, "Moved", $"Asset '{model.AssetId}' successfully moved.", 4000);
                Cancel();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Move Failed", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private void Cancel()
        {
            var target = !string.IsNullOrWhiteSpace(ReturnUrl) ? ReturnUrl : "/fixedassets/list";
            NavigationManager.NavigateTo(target);
        }
    }
}

