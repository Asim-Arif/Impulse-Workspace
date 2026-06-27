using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class DepartmentService : IDepartmentService
    {
        private readonly IDepartmentDataAccess _dataAccess;

        public DepartmentService(IDepartmentDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<DepartmentModel>> GetDepartmentsListAsync(bool showInactive)
        {
            return await _dataAccess.GetDepartmentsListAsync(showInactive);
        }

        public async Task<DepartmentModel?> GetDepartmentByIdAsync(string deptId)
        {
            return await _dataAccess.GetDepartmentByIdAsync(deptId);
        }

        public async Task SaveDepartmentAsync(DepartmentModel department, bool isAdd)
        {
            await _dataAccess.SaveDepartmentAsync(department, isAdd);
        }

        public async Task<string> GetNextDeptIdAsync()
        {
            return await _dataAccess.GetNextDeptIdAsync();
        }

        public async Task SetDepartmentActiveStateAsync(string deptId, bool active)
        {
            await _dataAccess.SetDepartmentActiveStateAsync(deptId, active);
        }
    }
}
