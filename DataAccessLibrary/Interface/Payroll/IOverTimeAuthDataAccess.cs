using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IOverTimeAuthDataAccess
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync();
        Task<List<OverTimeAuthRowDto>> GetOverTimeAuthListAsync(string deptId, DateTime date);
        Task<bool> SaveOverTimeAuthAsync(string deptId, DateTime date, List<OverTimeAuthRowDto> rows);
    }
}
