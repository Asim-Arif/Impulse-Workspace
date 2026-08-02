using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class AuthorizeReceivedService : IAuthorizeReceivedService
    {
        private readonly IAuthorizeReceivedDataAccess _dataAccess;

        public AuthorizeReceivedService(IAuthorizeReceivedDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<PendingAuthorizeReceivedItemModel>> GetPendingAuthorizeReceivedItemsAsync(AuthorizeReceivedFilterModel filter)
        {
            return _dataAccess.GetPendingAuthorizeReceivedItemsAsync(filter);
        }

        public Task<List<WastageTypeLookupModel>> GetWastageTypesAsync()
        {
            return _dataAccess.GetWastageTypesAsync();
        }

        public Task<List<RepairTypeLookupModel>> GetRepairTypesAsync()
        {
            return _dataAccess.GetRepairTypesAsync();
        }

        public Task<List<EmployeeLookupModel>> GetInspectorsAsync()
        {
            return _dataAccess.GetInspectorsAsync();
        }

        public Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            return _dataAccess.GetProcessesAsync();
        }

        public Task<List<MakerPOLookupModel>> GetMakersAsync()
        {
            return _dataAccess.GetMakersAsync();
        }

        public Task<bool> SaveAuthorizationAsync(List<PendingAuthorizeReceivedItemModel> items, string inspectorEmpId, string userName, string machineName, int userId)
        {
            return _dataAccess.SaveAuthorizationAsync(items, inspectorEmpId, userName, machineName, userId);
        }
    }
}
