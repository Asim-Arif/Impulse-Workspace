using DataAccessLibrary.Interface.Production;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class ProductionSequenceService : IProductionSequenceService
    {
        private readonly IProductionSequenceDataAccess _dataAccess;

        public ProductionSequenceService(IProductionSequenceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<string> GetNextBatchNoAsync()
        {
            return _dataAccess.GetNextBatchNoAsync();
        }

        public Task<string> GetNextMasterPONoAsync(DateTime dt)
        {
            return _dataAccess.GetNextMasterPONoAsync(dt);
        }

        public Task<string> GetNextHeaderReceiptIDAsync(DateTime dt)
        {
            return _dataAccess.GetNextHeaderReceiptIDAsync(dt);
        }

        public Task<string> GetNextDetailReceiptIDAsync(DateTime dt)
        {
            return _dataAccess.GetNextDetailReceiptIDAsync(dt);
        }

        public Task<string> GetNextSubLotNoAsync(string refLotNo)
        {
            return _dataAccess.GetNextSubLotNoAsync(refLotNo);
        }

        public Task<string> GetNextReceivingReceiptIDAsync(DateTime dt)
        {
            return _dataAccess.GetNextReceivingReceiptIDAsync(dt);
        }

        public Task<string> GetNextMainLotNoAsync(DateTime dt)
        {
            return _dataAccess.GetNextMainLotNoAsync(dt);
        }
    }
}
