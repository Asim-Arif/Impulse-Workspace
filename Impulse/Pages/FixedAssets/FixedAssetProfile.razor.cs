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
    public partial class FixedAssetProfile : ComponentBase
    {
        [Inject] private IFixedAssetProfileDataAccess DataAccess { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = null!;

        [Parameter] public int? EntryId { get; set; }

        [SupplyParameterFromQuery(Name = "copyFrom")]
        public int? CopyFromId { get; set; }

        [SupplyParameterFromQuery(Name = "categoryId")]
        public string? CategoryId { get; set; }

        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private FixedAssetProfileModel model = new();
        private FixedAssetLookupOptions lookup = new();
        private IReadOnlyList<FixedAssetRepairLogRow> repairLogs = Array.Empty<FixedAssetRepairLogRow>();
        private string activeTab = "general";
        private bool isBusy;
        private bool isNew => model.EntryId <= 0;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isBusy = true;
                lookup = await DataAccess.GetLookupOptionsAsync();

                if (EntryId.HasValue && EntryId.Value > 0)
                {
                    var existing = await DataAccess.GetAsync(EntryId.Value);
                    if (existing != null)
                    {
                        model = existing;
                        if (!string.IsNullOrWhiteSpace(model.AssetId))
                        {
                            repairLogs = await DataAccess.GetRepairLogsAsync(model.AssetId);
                        }
                    }
                    else
                    {
                        NotificationService.Notify(NotificationSeverity.Error, "Not Found", $"Asset with entry ID {EntryId} was not found.", 5000);
                        model = await DataAccess.NewAsync();
                    }
                }
                else if (CopyFromId.HasValue && CopyFromId.Value > 0)
                {
                    var source = await DataAccess.GetAsync(CopyFromId.Value);
                    if (source != null)
                    {
                        model = source;
                        model.EntryId = 0;
                        model.AssetId = "";
                        model.AcquiredDate = DateTime.Today;
                        model.PlacedInServiceDate = DateTime.Today;
                    }
                    else
                    {
                        model = await DataAccess.NewAsync();
                    }
                }
                else
                {
                    model = await DataAccess.NewAsync();
                    if (!string.IsNullOrWhiteSpace(CategoryId) && int.TryParse(CategoryId, out var catId))
                    {
                        model.MasterAssetRefId = catId;
                        ApplyCategoryDefaults(catId);
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading Profile", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private void OnMasterCategoryChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out var catId))
            {
                model.MasterAssetRefId = catId;
                ApplyCategoryDefaults(catId);
            }
            else
            {
                model.MasterAssetRefId = null;
            }
        }

        private void ApplyCategoryDefaults(int catId)
        {
            var match = lookup.Masters.FirstOrDefault(x => x.Id == catId.ToString());
            if (match != null && match.PercentRate > 0 && model.DepreciationRate == 0)
            {
                model.DepreciationRate = match.PercentRate;
            }
        }

        private void OnAcquiredDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.AcquiredDate = dt;
            }
            else
            {
                model.AcquiredDate = null;
            }
        }

        private async Task<bool> SaveAssetAsync()
        {
            if (string.IsNullOrWhiteSpace(model.AssetId))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Asset ID is required.", 4000);
                return false;
            }

            if (string.IsNullOrWhiteSpace(model.Description))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Description is required.", 4000);
                return false;
            }

            try
            {
                isBusy = true;
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var userName = authState.User.Identity?.Name ?? authState.User.FindFirstValue(ClaimTypes.Name) ?? "Administrator";

                var newId = await DataAccess.SaveAsync(model, userName, Environment.MachineName);
                model.EntryId = newId;

                NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Asset '{model.AssetId}' saved successfully.", 4000);
                return true;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Save Failed", ex.Message, 5000);
                return false;
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task SaveAndNewAsync()
        {
            var ok = await SaveAssetAsync();
            if (ok)
            {
                model = await DataAccess.NewAsync();
                repairLogs = Array.Empty<FixedAssetRepairLogRow>();
                activeTab = "general";
            }
        }

        private async Task SaveAndCloseAsync()
        {
            var ok = await SaveAssetAsync();
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

