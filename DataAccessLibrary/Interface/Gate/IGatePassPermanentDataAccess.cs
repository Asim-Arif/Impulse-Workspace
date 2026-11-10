using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IGatePassPermanentDataAccess
{
    Task<IReadOnlyList<GatePassPermanentListRowModel>> GetListAsync(
        GatePassPermanentFilterModel filter, 
        CancellationToken cancellationToken = default);

    Task<GatePassPermanentEntryModel?> GetByIdAsync(
        int entryId, 
        CancellationToken cancellationToken = default);

    Task<int> SaveAsync(
        GatePassPermanentEntryModel model, 
        string userName, 
        string machineName, 
        CancellationToken cancellationToken = default);

    Task DeleteAsync(
        int entryId, 
        CancellationToken cancellationToken = default);

    Task<string> GenerateGatePassNoAsync(
        DateTime dt, 
        CancellationToken cancellationToken = default);

    Task SetApprovalAsync(
        int entryId, 
        string approvalColumn, 
        bool value, 
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<CourierEmployeeLookupModel>> GetEmployeeOptionsAsync(
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<CourierVendorLookupModel>> GetVendorOptionsAsync(
        CancellationToken cancellationToken = default);
}

