using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IAdvanceLedgerService
    {
        /// <summary>
        /// Gets the list of loan & advance transactions matching the filter.
        /// </summary>
        Task<List<AdvanceLedgerItemModel>> GetAdvanceLedgerAsync(AdvanceLedgerFilterDto filter);

        /// <summary>
        /// Calculates the opening long-term loan balance for an employee strictly prior to fromDate.
        /// </summary>
        Task<decimal> GetLongTermOpeningBalanceAsync(string empId, DateTime fromDate);
    }
}
