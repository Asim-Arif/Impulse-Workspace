using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IAdvanceLedgerDataAccess
    {
        /// <summary>
        /// Queries VAdvances view for loan & advance ledger transactions matching the filter criteria.
        /// </summary>
        Task<List<AdvanceLedgerItemModel>> GetAdvanceLedgerAsync(AdvanceLedgerFilterDto filter);
    }
}
