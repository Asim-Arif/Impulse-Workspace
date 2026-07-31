using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IEmpFineLedgerService
    {
        /// <summary>
        /// Gets fine ledger records along with opening and running balance calculations.
        /// </summary>
        Task<(List<EmpFineLedgerItemModel> Items, EmpFineLedgerKpiSummaryModel Kpi)> GetFineLedgerDataAsync(EmpFineLedgerFilterDto filter);

        /// <summary>
        /// Deletes an un-finalized fine record.
        /// </summary>
        Task DeleteFineAsync(long entryId);
    }
}
