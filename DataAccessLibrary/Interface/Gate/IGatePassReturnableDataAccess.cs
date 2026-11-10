using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface IGatePassReturnableDataAccess
{
    Task<IReadOnlyList<GatePassReturnableListRowModel>> GetListAsync(GatePassReturnableFilterModel filter, CancellationToken cancellationToken = default);
    Task<GatePassReturnableEntryModel?> GetByIdAsync(int entryId, CancellationToken cancellationToken = default);
    Task<GatePassReceiveModel?> GetReceiveModelByIdAsync(int entryId, CancellationToken cancellationToken = default);
    Task<int> SaveAsync(GatePassReturnableEntryModel model, string userName, string machineName, CancellationToken cancellationToken = default);
    Task<int> SaveReceivingAsync(GatePassReceiveModel model, string userName, string machineName, CancellationToken cancellationToken = default);
    Task<string> GenerateGatePassNoAsync(DateTime date, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<GateLookupOptionModel>> GetGatePassTypesAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<CourierEmployeeLookupModel>> GetEmployeesAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<CourierVendorLookupModel>> GetVendorsAsync(CancellationToken cancellationToken = default);
}

