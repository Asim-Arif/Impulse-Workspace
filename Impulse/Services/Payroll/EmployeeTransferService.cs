using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class EmployeeTransferService : IEmployeeTransferService
    {
        private readonly IEmployeeTransferDataAccess _dataAccess;

        public EmployeeTransferService(IEmployeeTransferDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<DepartmentListItemModel>> GetSourceDepartmentsAsync() => _dataAccess.GetSourceDepartmentsAsync();

        public Task<List<DepartmentListItemModel>> GetTargetDepartmentsAsync(string sourceDeptId) => _dataAccess.GetTargetDepartmentsAsync(sourceDeptId);

        public Task<List<EmployeeListItemModel>> GetEmployeesByDepartmentAsync(string deptId) => _dataAccess.GetEmployeesByDepartmentAsync(deptId);

        public Task<bool> SaveTransferAsync(EmployeeTransferDto dto) => _dataAccess.SaveTransferAsync(dto);

        public Task<List<EmployeeTransferHistoryRow>> GetRecentTransfersAsync() => _dataAccess.GetRecentTransfersAsync();
    }
}
