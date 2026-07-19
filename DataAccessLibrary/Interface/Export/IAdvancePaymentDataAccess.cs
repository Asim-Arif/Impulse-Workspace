using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IAdvancePaymentDataAccess
    {
        Task<List<AdvancePaymentListViewModel>> GetAdvancePaymentListAsync(DateTime dtFrom, DateTime dtTo, string custCode, string country);
        Task<AdvancePaymentViewModel?> GetAdvancePaymentAsync(int entryId);
        Task<bool> SaveAdvancePaymentAsync(AdvancePaymentViewModel payment);
        Task<bool> UpdateExchangeRateAsync(int entryId, decimal exchRate);
        Task<string> PostToFinancialAsync(int entryId, AdvancePaymentViewModel payment, List<PrcDeductionModel> deductions, DateTime postingDate);
        Task<bool> DeleteAdvancePaymentAsync(int entryId);
        Task<List<AdvancePaymentUsageViewModel>> GetUsageDetailsAsync(int entryId);
        
        Task<List<AdvancePaymentBankModel>> GetBanksAsync();
        Task<List<string>> GetOrdersAsync(string custCode, string country);
        Task<(string Currency, string Address)> GetCustomerDetailsAsync(string custCode, string country);
    }
}
