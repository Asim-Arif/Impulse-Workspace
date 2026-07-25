using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IRMIssuanceDataAccess
    {
        Task<IEnumerable<MakerLookupModel>> GetMakersAsync();
        Task<IEnumerable<EmployeeLookupModel>> GetEmployeesAsync();
        Task<IEnumerable<RMLookupModel>> GetMaterialsAsync();
        Task<IEnumerable<RMLocationModel>> GetMaterialLocationsAsync(int rmid);
        Task<RMDetailsModel> GetRMDetailsAsync(int rmid);
        Task<IEnumerable<RMHistoryRowModel>> GetIssuanceHistoryAsync(bool isMaker, string recipientId, DateTime fromDate, DateTime toDate);
        Task<IEnumerable<IssuanceReceiptLookupModel>> GetProductionIssuancesAsync(int vendorId);
        Task<IEnumerable<RMLookupModel>> GetMaterialsByIssuanceAsync(int entryId);
        Task<string> SaveIssuanceTransactionAsync(RMIssuanceSaveRequest request);
    }
}
