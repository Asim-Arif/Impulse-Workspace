using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IMakerService
    {
        Task<MakerDto?> GetMakerByIdAsync(long vendId);
        Task<List<ProcessModel>> GetAllProcessesAsync();
        Task<List<int>> GetMakerProcessIdsAsync(long vendId);
        Task<List<MakerCategoryModel>> GetMakerCategoriesAsync();
        Task<List<string>> GetPaymentTermsAsync();
        Task<string> GenerateNextMakerCodeAsync();
        Task<string> GenerateNextAccNoAsync(string parentAccNo);
        Task<bool> IsNicDuplicateAsync(string nicNo, long? excludeVendId);
        Task<long> SaveMakerAsync(MakerDto dto, List<int> processIds);
        Task UpdateMakerAsync(MakerDto dto, List<int> processIds);
    }
}
