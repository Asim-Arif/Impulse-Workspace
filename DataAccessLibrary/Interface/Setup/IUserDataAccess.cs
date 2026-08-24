using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IUserDataAccess
    {
        Task<List<UserModel>> GetUsersAsync(string? search = null, bool? activeOnly = null);
        Task<UserModel?> GetUserByIdAsync(int userId);
        Task<UserModel?> GetUserByUserNameAsync(string userName);
        Task<int> CreateUserAsync(UserModel user);
        Task<bool> UpdateUserAsync(UserModel user);
        Task<bool> ToggleUserStatusAsync(int userId, bool inActive);
        Task<bool> ResetPasswordAsync(int userId, string newPassword);
        Task<bool> DeleteUserAsync(int userId);
        Task<int> GetTotalUsersCountAsync();
        Task<int> GetActiveUsersCountAsync();
    }
}
