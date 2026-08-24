using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace Impulse.Services.Setup
{
    public interface IUserService
    {
        Task<List<UserModel>> GetUsersAsync(string? search = null, bool? activeOnly = null);
        Task<UserModel?> GetUserByIdAsync(int userId);
        Task<UserModel?> GetUserByUserNameAsync(string userName);
        Task<(bool Success, string Message, int UserId)> CreateUserAsync(UserModel user);
        Task<(bool Success, string Message)> UpdateUserAsync(UserModel user);
        Task<bool> ToggleUserStatusAsync(int userId, bool inActive);
        Task<(bool Success, string Message)> ResetPasswordAsync(int userId, string newPassword);
        Task<bool> DeleteUserAsync(int userId);
        Task<(int TotalUsers, int ActiveUsers, int InactiveUsers)> GetUserStatisticsAsync();
    }
}
