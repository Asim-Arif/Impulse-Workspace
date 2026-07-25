using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class VendGateRcvdInspectionService : IVendGateRcvdInspectionService
    {
        private readonly IVendGateRcvdInspectionDataAccess _dataAccess;

        public VendGateRcvdInspectionService(IVendGateRcvdInspectionDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<VendGateRcvdInspectionMaterialItem>> GetMaterialsByRcvIDAsync(string rcvId)
        {
            return _dataAccess.GetMaterialsByRcvIDAsync(rcvId);
        }

        public Task<VendGateRcvdInspectionViewModel> GetInspectionByMaterialAsync(int entryId, int groupId)
        {
            return _dataAccess.GetInspectionByMaterialAsync(entryId, groupId);
        }

        public Task SaveInspectionAsync(VendGateRcvdInspectionViewModel model, string userName, string machineName)
        {
            return _dataAccess.SaveInspectionAsync(model, userName, machineName);
        }
    }
}
