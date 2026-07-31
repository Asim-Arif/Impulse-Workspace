using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IAdvanceShortDataAccess
    {
        /// <summary>
        /// Returns the employee's outstanding short-term advance balance.
        /// </summary>
        Task<AdvanceShortSummaryModel> GetEmployeeAdvanceSummaryAsync(string empId);

        /// <summary>
        /// Inserts a record into the Advances table (Type=0, DAmount=0) and returns.
        /// The voucher itself is saved by the service layer via IVoucherService before this is called.
        /// </summary>
        Task SaveAdvanceShortRecordAsync(AdvanceShortDto dto, string postedVchrNo);
    }
}
