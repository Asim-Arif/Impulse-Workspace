using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class EmpToEmpTransferService : IEmpToEmpTransferService
    {
        private readonly IEmpToEmpTransferDataAccess _dataAccess;
        private readonly IAuditService                _auditService;

        public EmpToEmpTransferService(IEmpToEmpTransferDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess   = dataAccess;
            _auditService = auditService;
        }

        public async Task<EmpTransferBalanceSummaryModel> GetEmployeeBalancesAsync(string empId)
        {
            return await _dataAccess.GetEmployeeBalancesAsync(empId);
        }

        public async Task ExecuteTransferAsync(EmpToEmpTransferDto dto)
        {
            string userName    = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.ExecuteEmpToEmpTransferAsync(dto, userName, machineName);
        }
    }
}
