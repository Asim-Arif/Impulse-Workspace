using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IEmpFineLedgerDataAccess
    {
        /// <summary>
        /// Gets the opening fine balance prior to fromDate.
        /// </summary>
        Task<decimal> GetOpeningBalanceAsync(string empId, DateTime fromDate);

        /// <summary>
        /// Queries fine ledger transactions for the date range.
        /// </summary>
        Task<List<EmpFineLedgerItemModel>> GetFineLedgerAsync(EmpFineLedgerFilterDto filter);

        /// <summary>
        /// Deletes a fine record if monthly salary for that month has not been finalized.
        /// </summary>
        Task DeleteFineAsync(long entryId);
    }
}
