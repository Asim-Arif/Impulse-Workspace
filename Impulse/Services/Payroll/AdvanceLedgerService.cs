using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class AdvanceLedgerService : IAdvanceLedgerService
    {
        private readonly IAdvanceLedgerDataAccess _dataAccess;

        public AdvanceLedgerService(IAdvanceLedgerDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<AdvanceLedgerItemModel>> GetAdvanceLedgerAsync(AdvanceLedgerFilterDto filter)
        {
            return await _dataAccess.GetAdvanceLedgerAsync(filter);
        }
    }
}
