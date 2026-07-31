using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IEmpToEmpTransferDataAccess
    {
        /// <summary>
        /// Gets current outstanding LT loan and ST advance balances for the given employee.
        /// </summary>
        Task<EmpTransferBalanceSummaryModel> GetEmployeeBalancesAsync(string empId);

        /// <summary>
        /// Executes the LT loan and/or ST advance transfer atomically inside a database transaction.
        /// </summary>
        Task ExecuteEmpToEmpTransferAsync(EmpToEmpTransferDto dto, string userName, string machineName);
    }
}
