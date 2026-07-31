using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IAdjustDedAmtDataAccess
    {
        /// <summary>
        /// Gets the latest active Long Term loan deduction details for the given employee.
        /// </summary>
        Task<EmpDeductionInfoModel> GetEmployeeLatestLoanDeductionAsync(string empId);

        /// <summary>
        /// Updates the monthly deduction installment amount (DAmount) for the specified Advances record.
        /// </summary>
        Task UpdateDeductionAmountAsync(long entryId, decimal newDAmount);
    }
}
