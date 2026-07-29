using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IEmployeeTransferDataAccess
    {
        Task<List<DepartmentListItemModel>> GetSourceDepartmentsAsync();
        Task<List<DepartmentListItemModel>> GetTargetDepartmentsAsync(string sourceDeptId);
        Task<List<EmployeeListItemModel>> GetEmployeesByDepartmentAsync(string deptId);
        Task<bool> SaveTransferAsync(EmployeeTransferDto dto);
        Task<List<EmployeeTransferHistoryRow>> GetRecentTransfersAsync();
    }
}
