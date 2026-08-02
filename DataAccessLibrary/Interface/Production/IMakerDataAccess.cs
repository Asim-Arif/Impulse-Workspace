using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerDataAccess
    {
        Task<MakerDto?> GetMakerByIdAsync(long vendId);
        Task<List<ProcessModel>> GetAllProcessesAsync();
        Task<List<int>> GetMakerProcessIdsAsync(long vendId);
        Task<List<MakerCategoryModel>> GetMakerCategoriesAsync();
        Task<List<string>> GetPaymentTermsAsync();
        Task<string> GenerateNextMakerCodeAsync();
        Task<string> GenerateNextAccNoAsync(string parentAccNo);
        Task<bool> IsNicDuplicateAsync(string nicNo, long? excludeVendId);
        Task<long> SaveMakerAsync(MakerDto dto);
        Task UpdateMakerAsync(MakerDto dto);
        Task SaveMakerProcessesAsync(long vendId, List<int> processIds);
    }
}
