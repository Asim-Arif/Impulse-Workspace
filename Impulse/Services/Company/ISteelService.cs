using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public interface ISteelService
    {
        Task<List<SteelTypeNodeModel>> GetFullSteelTreeAsync();
        Task<List<SteelCompModel>> GetAllCompositionsAsync();
        Task<List<string>> GetUnitsAsync();
        
        Task<bool> SaveSteelTypeAsync(SteelTypeNodeModel model);
        Task<bool> DeleteSteelTypeAsync(int steelId);

        Task<bool> SaveSteelGageAsync(SteelGageNodeModel model, bool isInsert);
        Task<bool> DeleteSteelGageAsync(string gageId);
    }
}
