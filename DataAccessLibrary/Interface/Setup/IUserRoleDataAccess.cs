using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IUserRoleDataAccess
    {
        Task<List<UserRoleModel>> GetAllRolesAsync();
        Task<bool> AddRoleAsync(string roleName);
        Task<bool> DeleteRoleAsync(string roleName);
        Task<bool> RoleExistsAsync(string roleName);
        Task<List<string>> GetRolesByUserIdAsync(int userId);
        Task<Dictionary<int, List<string>>> GetAllUserRolesMapAsync();
        Task<List<UserModel>> GetUsersByRoleAsync(string roleName);
        Task<bool> SaveUserRolesAsync(int userId, IEnumerable<string> roles);
        Task<bool> CopyUserRolesAsync(int sourceUserId, int targetUserId);
    }
}
