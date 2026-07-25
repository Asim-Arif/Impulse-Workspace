using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IVendGateRcvdInspectionDataAccess
    {
        Task<List<VendGateRcvdInspectionMaterialItem>> GetMaterialsByRcvIDAsync(string rcvId);
        Task<VendGateRcvdInspectionViewModel> GetInspectionByMaterialAsync(int entryId, int groupId);
        Task SaveInspectionAsync(VendGateRcvdInspectionViewModel model, string userName, string machineName);
    }
}
