using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public class RMIssuanceListService : IRMIssuanceListService
    {
        private readonly IRMIssuanceListDataAccess _dataAccess;

        public RMIssuanceListService(IRMIssuanceListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<IEnumerable<RMGroupLookupModel>> GetRMGroupsAsync()
        {
            return await _dataAccess.GetRMGroupsAsync();
        }

        public async Task<IEnumerable<DepartmentLookupModel>> GetDepartmentsAsync()
        {
            return await _dataAccess.GetDepartmentsAsync();
        }

        public async Task<IEnumerable<RMIssuanceListRowModel>> GetRMIssuanceListAsync(
            DateTime dtFrom, 
            DateTime dtTo, 
            string? materialId = null, 
            int? groupId = null, 
            string? employeeId = null, 
            int? vendorId = null, 
            string? deptId = null)
        {
            return await _dataAccess.GetRMIssuanceListAsync(dtFrom, dtTo, materialId, groupId, employeeId, vendorId, deptId);
        }

        public async Task<bool> DeleteIssuanceAsync(string issNo)
        {
            return await _dataAccess.DeleteIssuanceAsync(issNo);
        }
    }
}
