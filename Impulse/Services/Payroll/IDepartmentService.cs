using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IDepartmentService
    {
        Task<List<DepartmentModel>> GetDepartmentsListAsync(bool showInactive);
        Task<DepartmentModel?> GetDepartmentByIdAsync(string deptId);
        Task SaveDepartmentAsync(DepartmentModel department, bool isAdd);
        Task<string> GetNextDeptIdAsync();
        Task SetDepartmentActiveStateAsync(string deptId, bool active);
    }
}
