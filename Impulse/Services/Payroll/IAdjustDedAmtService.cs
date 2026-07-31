using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IAdjustDedAmtService
    {
        /// <summary>
        /// Gets the latest active Long Term loan deduction details for the given employee.
        /// </summary>
        Task<EmpDeductionInfoModel> GetEmployeeLatestLoanDeductionAsync(string empId);

        /// <summary>
        /// Updates the monthly deduction installment amount (DAmount) for the specified Advances record.
        /// </summary>
        Task UpdateDeductionAmountAsync(AdjustDedAmtDto dto);
    }
}
