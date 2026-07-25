using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public interface IRMIssuanceListService
    {
        Task<IEnumerable<RMGroupLookupModel>> GetRMGroupsAsync();
        Task<IEnumerable<DepartmentLookupModel>> GetDepartmentsAsync();
        Task<IEnumerable<RMIssuanceListRowModel>> GetRMIssuanceListAsync(
            DateTime dtFrom, 
            DateTime dtTo, 
            string? materialId = null, 
            int? groupId = null, 
            string? employeeId = null, 
            int? vendorId = null, 
            string? deptId = null);
        Task<bool> DeleteIssuanceAsync(string issNo);
    }
}
