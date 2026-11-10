using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Gate;

namespace DataAccessLibrary.Interface.Gate;

public interface ICourierIssuanceDataAccess
{
    Task<IReadOnlyList<CourierIssuanceListRowModel>> GetCourierIssuanceListAsync(
        CourierIssuanceFilterModel filter, 
        CancellationToken cancellationToken = default);

    Task<CourierIssuanceEntryModel?> GetCourierIssuanceAsync(
        int entryId, 
        CancellationToken cancellationToken = default);

    Task<int> SaveCourierIssuanceAsync(
        CourierIssuanceEntryModel model, 
        string userName, 
        string machineName, 
        CancellationToken cancellationToken = default);

    Task DeleteCourierIssuanceAsync(
        int entryId, 
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<CourierVendorLookupModel>> GetVendorOptionsAsync(
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<CourierEmployeeLookupModel>> GetEmployeeOptionsAsync(
        CancellationToken cancellationToken = default);

    Task<(string? Name, string? Address)> GetVendorDetailsAsync(
        string accNo, 
        CancellationToken cancellationToken = default);
}

