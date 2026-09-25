using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using BlazorContextMenu;
using Impulse.Models.Setup;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages
{
    public partial class FavouritesDashboard : ComponentBase, IDisposable
    {
        [Inject]
        public IFavouriteService FavouriteService { get; set; } = null!;

        [Inject]
        public IUserPermissionService PermissionService { get; set; } = null!;

        [Inject]
        public NotificationService NotificationService { get; set; } = null!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = null!;

        protected List<FavouriteLinkDto> Favourites { get; set; } = new();
        protected bool isLoading = true;

        protected override async Task OnInitializedAsync()
        {
            await PermissionService.InitializeAsync();
            FavouriteService.OnFavouritesChanged += OnFavouritesChangedHandler;
            await LoadFavouritesAsync();
        }

        protected async Task LoadFavouritesAsync()
        {
            isLoading = true;
            try
            {
                Favourites = await FavouriteService.GetCurrentUserFavouritesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = "Unable to load favourite links: " + ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private void OnFavouritesChangedHandler()
        {
            _ = InvokeAsync(async () =>
            {
                await LoadFavouritesAsync();
                StateHasChanged();
            });
        }

        protected async Task RemoveFavouriteAsync(string optionId, string title)
        {
            var result = await FavouriteService.RemoveCurrentUserFavouriteAsync(optionId);
            if (result.Success)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Favourites",
                    Detail = $"'{title}' was removed from your favourites.",
                    Duration = 3000
                });
                await LoadFavouritesAsync();
            }
        }

        protected async Task HandleRemoveFromMenu(ItemClickEventArgs e)
        {
            if (e.Data is FavouriteLinkDto item)
            {
                await RemoveFavouriteAsync(item.OptionId, item.Title);
            }
        }

        protected string GetModuleBadgeClass(string module)
        {
            return (module?.ToLowerInvariant()) switch
            {
                "stock" => "bg-purple-subtle text-purple",
                "production" => "bg-danger-subtle text-danger",
                "accounts" or "financial" => "bg-primary-subtle text-primary",
                "payroll" => "bg-success-subtle text-success",
                "export" => "bg-warning-subtle text-warning-emphasis",
                "company" => "bg-info-subtle text-info-emphasis",
                "dashboard" or "dashboards" => "bg-secondary-subtle text-secondary",
                _ => "bg-light text-secondary"
            };
        }

        protected string GetModuleIconColor(string module)
        {
            return (module?.ToLowerInvariant()) switch
            {
                "stock" => "text-purple",
                "production" => "text-danger",
                "accounts" or "financial" => "text-primary",
                "payroll" => "text-success",
                "export" => "text-warning",
                "company" => "text-info",
                _ => "text-secondary"
            };
        }

        public void Dispose()
        {
            FavouriteService.OnFavouritesChanged -= OnFavouritesChangedHandler;
        }
    }
}
