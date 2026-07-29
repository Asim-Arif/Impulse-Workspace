using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class OverTimeAuthService : IOverTimeAuthService
    {
        private readonly IOverTimeAuthDataAccess _dataAccess;

        public OverTimeAuthService(IOverTimeAuthDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<DepartmentListItemModel>> GetDepartmentsAsync() => _dataAccess.GetDepartmentsAsync();

        public Task<List<OverTimeAuthRowDto>> GetOverTimeAuthListAsync(string deptId, System.DateTime date) => _dataAccess.GetOverTimeAuthListAsync(deptId, date);

        public Task<bool> SaveOverTimeAuthAsync(string deptId, System.DateTime date, List<OverTimeAuthRowDto> rows) => _dataAccess.SaveOverTimeAuthAsync(deptId, date, rows);
    }
}
