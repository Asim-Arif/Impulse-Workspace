using System;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.FixedAssets
{
    public partial class FixedAssetsDashboard : ComponentBase
    {
        [Inject] public IFixedAssetDashboardDataAccess DataAccess { get; set; } = default!;
        [Inject] public NavigationManager Navigation { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;

        protected FixedAssetDashboardFilter Filter { get; set; } = new();
        protected FixedAssetDashboardResult Data { get; set; } = new();
        protected bool IsLoading { get; set; } = true;
        protected string? ErrorMessage { get; set; }
        protected string ActiveTab { get; set; } = "top-assets";

        protected static readonly (int Index, string Name, string ShortName)[] QuarterList = new[]
        {
            (0, "All Quarters", "All"),
            (1, "Q1 (Jul - Sep)", "Q1"),
            (2, "Q2 (Oct - Dec)", "Q2"),
            (3, "Q3 (Jan - Mar)", "Q3"),
            (4, "Q4 (Apr - Jun)", "Q4")
        };

        protected override async Task OnInitializedAsync()
        {
            await LoadDashboardDataAsync();
        }

        protected async Task LoadDashboardDataAsync()
        {
            try
            {
                IsLoading = true;
                ErrorMessage = null;
                Data = await DataAccess.GetDashboardDataAsync(Filter);
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Failed to load dashboard data: {ex.Message}";
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task SetFiscalYearAsync(string fy)
        {
            Filter.FiscalYear = fy;
            await LoadDashboardDataAsync();
        }

        protected async Task SetQuarterAsync(int quarterIndex)
        {
            Filter.QuarterIndex = quarterIndex;
            await LoadDashboardDataAsync();
        }

        protected async Task OnFilterChangedAsync()
        {
            await LoadDashboardDataAsync();
        }

        protected async Task ResetFilterAsync()
        {
            Filter = new FixedAssetDashboardFilter();
            await LoadDashboardDataAsync();
        }

        protected void SetActiveTab(string tab)
        {
            ActiveTab = tab;
        }

        protected void NavigateToProfile(string? assetId = null)
        {
            if (string.IsNullOrWhiteSpace(assetId))
                Navigation.NavigateTo("/fixedassets/profile");
            else
                Navigation.NavigateTo($"/fixedassets/profile?assetId={assetId}");
        }

        protected void NavigateTo(string url)
        {
            Navigation.NavigateTo(url);
        }

        protected static string FormatCurrency(decimal value)
        {
            return value.ToString("N0");
        }

        protected static string FormatShortNumber(decimal value)
        {
            if (value >= 1_000_000_000)
                return (value / 1_000_000_000m).ToString("0.##") + " B";
            if (value >= 1_000_000)
                return (value / 1_000_000m).ToString("0.##") + " M";
            if (value >= 1_000)
                return (value / 1_000m).ToString("0.##") + " K";
            return value.ToString("N0");
        }
    }
}

