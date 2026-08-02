using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IReceiveAgainstPODataAccess
    {
        Task<MasterPOSearchResultModel?> SearchMasterPOAsync(string receiptId);
        Task<List<MasterPOLineItemModel>> GetMasterPOLinesAsync(long vendIssuedEntryId);
        Task<long> SavePOReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName);
    }
}
