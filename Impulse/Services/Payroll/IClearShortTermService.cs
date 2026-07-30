using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IClearShortTermService
    {
        /// <summary>
        /// Returns the employee's open short-term advances
        /// (Advances WHERE Type=0 AND Cleared=0) for the checkbox table.
        /// </summary>
        Task<List<ShortTermAdvanceItem>> GetOpenAdvancesAsync(string empId);

        /// <summary>
        /// Returns the AccNo configured in GeneralData for EmpShortTermAccNo.
        /// Used at page init to surface a clear error if not configured.
        /// </summary>
        Task<string> GetEmpShortTermAccNoAsync();

        /// <summary>
        /// Posts the receipt voucher (CRV/JV) via IVoucherService unless
        /// "Don't Post Voucher" was selected, then records the clearance
        /// (header + details + Advances.Cleared=2) via the data access layer.
        /// Returns the posted voucher number, or an empty string when no
        /// voucher was posted.
        /// </summary>
        Task<string> SaveAsync(ClearShortTermDto dto);
    }
}
