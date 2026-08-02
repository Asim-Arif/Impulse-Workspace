using System;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IProductionSequenceService
    {
        Task<string> GetNextBatchNoAsync();
        Task<string> GetNextMasterPONoAsync(DateTime dt);
        Task<string> GetNextHeaderReceiptIDAsync(DateTime dt);
        Task<string> GetNextDetailReceiptIDAsync(DateTime dt);
        Task<string> GetNextSubLotNoAsync(string refLotNo);
        Task<string> GetNextReceivingReceiptIDAsync(DateTime dt);
        Task<string> GetNextMainLotNoAsync(DateTime dt);
    }
}
