using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IWastageTypeDataAccess
    {
        Task<List<WastageTypeModel>> GetWastageTypesAsync();
        Task<WastageTypeDto?> GetWastageTypeByIdAsync(int entryId);
        Task<int> SaveWastageTypeAsync(WastageTypeDto model);
        Task<bool> UpdateWastageTypeAsync(WastageTypeDto model);
        Task<(bool CanDelete, string Reason)> CanDeleteWastageTypeAsync(int entryId);
        Task<bool> DeleteWastageTypeAsync(int entryId);
    }
}
