using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class ReceiveAgainstPOService : IReceiveAgainstPOService
    {
        private readonly IReceiveAgainstPODataAccess _dataAccess;

        public ReceiveAgainstPOService(IReceiveAgainstPODataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<MasterPOSearchResultModel?> SearchMasterPOAsync(string receiptId)
        {
            return _dataAccess.SearchMasterPOAsync(receiptId);
        }

        public Task<List<MasterPOLineItemModel>> GetMasterPOLinesAsync(long vendIssuedEntryId)
        {
            return _dataAccess.GetMasterPOLinesAsync(vendIssuedEntryId);
        }

        public Task<long> SavePOReceivingAsync(CreatePOReceivingHeaderModel header, List<CreatePOReceivingLineModel> lines, string userName, int userId, string machineName)
        {
            return _dataAccess.SavePOReceivingAsync(header, lines, userName, userId, machineName);
        }
    }
}
