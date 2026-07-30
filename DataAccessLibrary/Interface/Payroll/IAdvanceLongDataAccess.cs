using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IAdvanceLongDataAccess
    {
        /// <summary>
        /// Returns the employee's outstanding long-term loan balance
        /// and their most recent monthly deduction amount.
        /// </summary>
        Task<AdvanceLongSummaryModel> GetEmployeeAdvanceSummaryAsync(string empId);

        /// <summary>
        /// Inserts a record into the Advances table (Type=1) and returns
        /// the voucher number that was posted. The voucher itself is saved
        /// by the service layer via IVoucherService before this is called.
        /// </summary>
        Task SaveAdvanceLongRecordAsync(AdvanceLongDto dto, string postedVchrNo);
    }
}
