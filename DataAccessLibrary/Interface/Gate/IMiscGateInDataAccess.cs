using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IMiscGateInDataAccess
{
    Task<IReadOnlyList<MiscGateInListRowModel>> GetVisitorsAsync(MiscGateInFilterModel filter, CancellationToken cancellationToken = default);
    Task<MiscGateInEntryModel?> GetVisitorAsync(int entryId, CancellationToken cancellationToken = default);
    Task<MiscGateOutModel?> GetVisitorGateOutAsync(int entryId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeeOptionsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetVendorOptionsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeeFilterOptionsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetVendorFilterOptionsAsync(CancellationToken cancellationToken = default);
    Task<int> SaveVisitorAsync(MiscGateInEntryModel visitor, string userName, CancellationToken cancellationToken = default);
    Task SaveVisitorGateOutAsync(int entryId, DateTime timeOut, string verificationInput, string userName, CancellationToken cancellationToken = default);
    Task<ConcernedPersonContactModel?> GetConcernedPersonContactAsync(string employeeId, CancellationToken cancellationToken = default);
    Task<VisitorPassCardModel> EnsurePassAsync(int entryId, CancellationToken cancellationToken = default);
}

