using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IWastageTypeService
    {
        Task<List<WastageTypeModel>> GetWastageTypesAsync();
        Task<WastageTypeDto?> GetWastageTypeByIdAsync(int entryId);
        Task<int> SaveWastageTypeAsync(WastageTypeDto model);
        Task<bool> UpdateWastageTypeAsync(WastageTypeDto model);
        Task<(bool CanDelete, string Reason)> CanDeleteWastageTypeAsync(int entryId);
        Task<bool> DeleteWastageTypeAsync(int entryId);
    }
}
