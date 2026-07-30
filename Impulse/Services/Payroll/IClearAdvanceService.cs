using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IClearAdvanceService
    {
        /// <summary>
        /// Returns the employee's outstanding long-term loan balance,
        /// current monthly deduction, and latest advance EntryID for display.
        /// </summary>
        Task<ClearAdvanceSummaryModel> GetSummaryAsync(string empId);

        /// <summary>
        /// Returns the AccNo configured in GeneralData for EmpLongTermAccNo.
        /// Used at page init to surface a clear error if not configured.
        /// </summary>
        Task<string> GetEmpLongTermAccNoAsync();

        /// <summary>
        /// Posts the receipt voucher (CRV/JV) via IVoucherService unless
        /// "Don't Post Voucher" was selected, then records the clearance in
        /// the AmtCleared table (and updates the latest advance's DAmount
        /// when a new deduction was supplied).
        /// Returns the posted voucher number, or an empty string when no
        /// voucher was posted.
        /// </summary>
        Task<string> SaveAsync(ClearAdvanceDto dto);
    }
}
