using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IEmpToEmpTransferService
    {
        /// <summary>
        /// Gets current outstanding LT loan and ST advance balances for the source employee.
        /// </summary>
        Task<EmpTransferBalanceSummaryModel> GetEmployeeBalancesAsync(string empId);

        /// <summary>
        /// Executes the LT loan and/or ST advance transfer from Source Employee to Destination Employee.
        /// </summary>
        Task ExecuteTransferAsync(EmpToEmpTransferDto dto);
    }
}
