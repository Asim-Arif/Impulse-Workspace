using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface IUserPermissionDataAccess
    {
        Task<List<MenuOptionModel>> GetAllMenuOptionsAsync();
        Task<List<MenuOptionModel>> GetMenuOptionsByModuleAsync(string moduleName);
        Task<List<string>> GetDistinctModulesAsync();
        Task<HashSet<string>> GetUserMenuOptionIdsAsync(int userId);
        Task<HashSet<string>> GetUserMenuOptionIdsByUserNameAsync(string userName);
        Task<UserModel?> GetUserPermissionsByUserNameAsync(string userName);
        Task<bool> SaveUserMenuOptionsAsync(int userId, string moduleName, IEnumerable<string> selectedOptionIds);
        Task<bool> SaveAllUserMenuOptionsAsync(int userId, IEnumerable<string> allOptionIds);
        Task<bool> CopyUserMenuOptionsAsync(int fromUserId, int toUserId);
        Task<bool> SyncMenuOptionsAsync(IEnumerable<MenuOptionModel> options);
    }
}
