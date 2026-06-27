using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IDepartmentDataAccess
    {
        Task<List<DepartmentModel>> GetDepartmentsListAsync(bool showInactive);
        Task<DepartmentModel?> GetDepartmentByIdAsync(string deptId);
        Task SaveDepartmentAsync(DepartmentModel department, bool isAdd);
        Task<string> GetNextDeptIdAsync();
        Task SetDepartmentActiveStateAsync(string deptId, bool active);
    }
}
