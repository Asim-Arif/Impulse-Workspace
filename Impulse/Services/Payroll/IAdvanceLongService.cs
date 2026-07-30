using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IAdvanceLongService
    {
        /// <summary>
        /// Returns the employee's outstanding long-term loan balance
        /// and their most recent monthly deduction amount for display.
        /// </summary>
        Task<AdvanceLongSummaryModel> GetEmployeeSummaryAsync(string empId);

        /// <summary>
        /// Builds and posts the appropriate voucher (CPV/JV/BPV) via IVoucherService,
        /// then records the advance in the Advances table.
        /// Returns the posted voucher number.
        /// </summary>
        Task<string> SaveAsync(AdvanceLongDto dto);

        /// <summary>
        /// Returns the AccNo configured in GeneralData for EmpLongTermAccNo.
        /// Used at page init to surface a clear error if not configured.
        /// </summary>
        Task<string> GetEmpLongTermAccNoAsync();

        /// <summary>
        /// Previews the next voucher number for the given date and payment mode.
        /// </summary>
        Task<string> GetPreviewVchrNoAsync(System.DateTime date, AdvanceLongPaymentMode mode);
    }
}
