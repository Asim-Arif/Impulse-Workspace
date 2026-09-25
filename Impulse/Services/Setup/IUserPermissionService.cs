using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace Impulse.Services.Setup
{
    public interface IUserPermissionService
    {
        Task InitializeAsync(string? userName = null);
        bool IsAdministrator { get; }
        int CurrentUserId { get; }
        string CurrentUserName { get; }
        bool HasModuleAccess(string moduleName);
        bool HasOptionAccess(string optionId);
        bool IsRouteAuthorized(string relativePath);
        Task<List<MenuOptionModel>> GetMenuOptionsByModuleAsync(string moduleName);
        Task<List<string>> GetDistinctModulesAsync();
        Task<HashSet<string>> GetUserMenuOptionIdsAsync(int userId);
        Task<bool> SaveUserMenuOptionsAsync(int userId, string moduleName, IEnumerable<string> selectedOptionIds);
        Task<bool> SaveAllUserMenuOptionsAsync(int userId, IEnumerable<string> allOptionIds);
        Task<bool> SyncMenuOptionsAsync(IEnumerable<MenuOptionModel> options);
        Task<bool> ResetAllOptionsAsync();
    }
}
