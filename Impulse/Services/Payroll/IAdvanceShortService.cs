using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IAdvanceShortService
    {
        /// <summary>
        /// Returns the employee's outstanding short-term advance balance for display.
        /// </summary>
        Task<AdvanceShortSummaryModel> GetEmployeeSummaryAsync(string empId);

        /// <summary>
        /// Builds and posts the appropriate voucher (CPV/JV/BPV) via IVoucherService,
        /// then records the advance in the Advances table.
        /// Returns the posted voucher number.
        /// </summary>
        Task<string> SaveAsync(AdvanceShortDto dto);

        /// <summary>
        /// Returns the AccNo configured in GeneralData for EmpShortTermAccNo.
        /// Used at page init to surface a clear error if not configured.
        /// </summary>
        Task<string> GetEmpShortTermAccNoAsync();

        /// <summary>
        /// Previews the next voucher number for the given date and payment mode.
        /// </summary>
        Task<string> GetPreviewVchrNoAsync(DateTime date, AdvanceShortPaymentMode mode);
    }
}
