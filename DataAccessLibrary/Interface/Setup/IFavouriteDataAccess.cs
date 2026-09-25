using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IFavouriteDataAccess
    {
        Task<List<FavouriteItemModel>> GetUserFavouritesAsync(int userId);
        Task<bool> AddFavouriteAsync(int userId, string optionId);
        Task<bool> RemoveFavouriteAsync(int userId, string optionId);
        Task<bool> IsFavouriteAsync(int userId, string optionId);
    }
}
