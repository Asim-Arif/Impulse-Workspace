using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IEmployeeTransferService
    {
        Task<List<DepartmentListItemModel>> GetSourceDepartmentsAsync();
        Task<List<DepartmentListItemModel>> GetTargetDepartmentsAsync(string sourceDeptId);
        Task<List<EmployeeListItemModel>> GetEmployeesByDepartmentAsync(string deptId);
        Task<bool> SaveTransferAsync(EmployeeTransferDto dto);
        Task<List<EmployeeTransferHistoryRow>> GetRecentTransfersAsync();
    }
}
