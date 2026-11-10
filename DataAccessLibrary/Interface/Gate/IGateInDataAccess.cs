using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IGateInDataAccess
{
    Task<IReadOnlyList<GateInListRowModel>> GetGateInEntriesAsync(GateInFilterModel filter, CancellationToken cancellationToken = default);
    Task<GateInEntryModel?> GetGateInAsync(int entryId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeesAsync(string departmentId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateInEmployeeLookupModel>> SearchEmployeesAsync(string searchText, string? deptId = null, CancellationToken cancellationToken = default);
    Task<GateInEmployeeLookupModel?> GetEmployeeByIdAsync(string empId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateInVendorLookupModel>> SearchVendorsAsync(string searchText, CancellationToken cancellationToken = default);
    Task<GateInVendorLookupModel?> GetVendorByAccNoAsync(string accNo, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetVendorOptionsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetVendorFilterOptionsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetGatePassTypeOptionsAsync(CancellationToken cancellationToken = default);
    Task<string> GetNextGatePassNoAsync(DateTime date, CancellationToken cancellationToken = default);
    Task<int> SaveGateInAsync(GateInEntryModel entry, string userName, string machineName, CancellationToken cancellationToken = default);
    Task<int> SaveGateInAsync(GateInEntryModel entry, string userName, CancellationToken cancellationToken = default);
    Task ApproveFinanceAsync(int entryId, string userName, CancellationToken cancellationToken = default);
    Task<bool> SetApprovalAsync(int entryId, string approvalColumn, bool approved, CancellationToken cancellationToken = default);
    Task<bool> DeleteGateInAsync(int entryId, CancellationToken cancellationToken = default);
    Task<GateReceiveModel?> GetReceiveAsync(int entryId, CancellationToken cancellationToken = default);
    Task SaveReceiveAsync(GateReceiveModel model, string userName, CancellationToken cancellationToken = default);
}
