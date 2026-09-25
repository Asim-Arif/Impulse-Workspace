using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Impulse.Models.Setup;

namespace Impulse.Services.Setup
{
    public interface IFavouriteService
    {
        Task<List<FavouriteLinkDto>> GetCurrentUserFavouritesAsync();
        Task<(bool Success, string Message)> AddCurrentUserFavouriteAsync(string optionId, string? title = null, string? href = null, string? module = null, string? iconClass = null);
        Task<(bool Success, string Message)> RemoveCurrentUserFavouriteAsync(string optionId);
        Task<bool> IsCurrentUserFavouriteAsync(string optionId);
        Task<int> GetCurrentUserFavouritesCountAsync();
        event Action? OnFavouritesChanged;
    }
}
