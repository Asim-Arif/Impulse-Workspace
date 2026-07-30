using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IClearAdvanceDataAccess
    {
        /// <summary>
        /// Returns the employee's outstanding long-term loan balance,
        /// the current monthly deduction, and the EntryID of their latest
        /// Type=1 advance (null when no long-term advance exists).
        /// </summary>
        Task<ClearAdvanceSummaryModel> GetSummaryAsync(string empId);

        /// <summary>
        /// Inserts a row into the AmtCleared table and, when a NewDeduction &gt; 0
        /// is supplied and a latest advance exists, updates Advances.DAmount.
        /// The voucher (if any) has already been posted by the service layer
        /// and its VchrNo is passed in here.
        /// </summary>
        Task SaveClearanceAsync(ClearAdvanceDto dto, string vchrNo);
    }
}
