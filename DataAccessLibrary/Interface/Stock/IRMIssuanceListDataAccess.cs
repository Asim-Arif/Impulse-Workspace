using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IRMIssuanceListDataAccess
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
