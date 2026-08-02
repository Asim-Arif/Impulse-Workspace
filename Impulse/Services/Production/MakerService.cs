using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerService : IMakerService
    {
        private readonly IMakerDataAccess _makerDataAccess;

        public MakerService(IMakerDataAccess makerDataAccess)
        {
            _makerDataAccess = makerDataAccess;
        }

        public Task<MakerDto?> GetMakerByIdAsync(long vendId)
        {
            return _makerDataAccess.GetMakerByIdAsync(vendId);
        }

        public Task<List<ProcessModel>> GetAllProcessesAsync()
        {
            return _makerDataAccess.GetAllProcessesAsync();
        }

        public Task<List<int>> GetMakerProcessIdsAsync(long vendId)
        {
            return _makerDataAccess.GetMakerProcessIdsAsync(vendId);
        }

        public Task<List<MakerCategoryModel>> GetMakerCategoriesAsync()
        {
            return _makerDataAccess.GetMakerCategoriesAsync();
        }

        public Task<List<string>> GetPaymentTermsAsync()
        {
            return _makerDataAccess.GetPaymentTermsAsync();
        }

        public Task<string> GenerateNextMakerCodeAsync()
        {
            return _makerDataAccess.GenerateNextMakerCodeAsync();
        }

        public Task<string> GenerateNextAccNoAsync(string parentAccNo)
        {
            return _makerDataAccess.GenerateNextAccNoAsync(parentAccNo);
        }

        public Task<bool> IsNicDuplicateAsync(string nicNo, long? excludeVendId)
        {
            return _makerDataAccess.IsNicDuplicateAsync(nicNo, excludeVendId);
        }

        public async Task<long> SaveMakerAsync(MakerDto dto, List<int> processIds)
        {
            long vendId = await _makerDataAccess.SaveMakerAsync(dto);
            await _makerDataAccess.SaveMakerProcessesAsync(vendId, processIds);
            return vendId;
        }

        public async Task UpdateMakerAsync(MakerDto dto, List<int> processIds)
        {
            await _makerDataAccess.UpdateMakerAsync(dto);
            await _makerDataAccess.SaveMakerProcessesAsync(dto.VendID, processIds);
        }
    }
}
