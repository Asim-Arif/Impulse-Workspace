using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
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

        /// <summary>
        /// Calculates the opening long-term loan balance for an employee strictly prior to fromDate.
        /// </summary>
        Task<decimal> GetLongTermOpeningBalanceAsync(string empId, DateTime fromDate);
    }
}
