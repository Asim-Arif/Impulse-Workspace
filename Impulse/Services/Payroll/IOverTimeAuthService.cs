using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IOverTimeAuthService
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync();
        Task<List<OverTimeAuthRowDto>> GetOverTimeAuthListAsync(string deptId, System.DateTime date);
        Task<bool> SaveOverTimeAuthAsync(string deptId, System.DateTime date, List<OverTimeAuthRowDto> rows);
    }
}
