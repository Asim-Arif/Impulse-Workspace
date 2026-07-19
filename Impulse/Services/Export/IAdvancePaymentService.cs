using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IAdvancePaymentService
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

    public class AdvancePaymentService : IAdvancePaymentService
    {
        private readonly IAdvancePaymentDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public AdvancePaymentService(IAdvancePaymentDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public async Task<List<AdvancePaymentListViewModel>> GetAdvancePaymentListAsync(DateTime dtFrom, DateTime dtTo, string custCode, string country)
        {
            return await _dataAccess.GetAdvancePaymentListAsync(dtFrom, dtTo, custCode, country);
        }

        public async Task<AdvancePaymentViewModel?> GetAdvancePaymentAsync(int entryId)
        {
            return await _dataAccess.GetAdvancePaymentAsync(entryId);
        }

        public async Task<bool> SaveAdvancePaymentAsync(AdvancePaymentViewModel payment)
        {
            return await _dataAccess.SaveAdvancePaymentAsync(payment);
        }

        public async Task<bool> UpdateExchangeRateAsync(int entryId, decimal exchRate)
        {
            return await _dataAccess.UpdateExchangeRateAsync(entryId, exchRate);
        }

        /// <summary>
        /// Posts to the financial ledger. Populates UserName/MachineName from IAuditService
        /// before delegating to the data access layer.
        /// </summary>
        public async Task<string> PostToFinancialAsync(int entryId, AdvancePaymentViewModel payment, List<PrcDeductionModel> deductions, DateTime postingDate)
        {
            payment.UserName = _auditService.GetCurrentUserName();
            payment.MachineName = _auditService.GetClientIpAddress();
            return await _dataAccess.PostToFinancialAsync(entryId, payment, deductions, postingDate);
        }

        public async Task<bool> DeleteAdvancePaymentAsync(int entryId)
        {
            return await _dataAccess.DeleteAdvancePaymentAsync(entryId);
        }

        public async Task<List<AdvancePaymentBankModel>> GetBanksAsync()
        {
            return await _dataAccess.GetBanksAsync();
        }

        public async Task<List<string>> GetOrdersAsync(string custCode, string country)
        {
            return await _dataAccess.GetOrdersAsync(custCode, country);
        }

        public async Task<(string Currency, string Address)> GetCustomerDetailsAsync(string custCode, string country)
        {
            return await _dataAccess.GetCustomerDetailsAsync(custCode, country);
        }

        public async Task<List<AdvancePaymentUsageViewModel>> GetUsageDetailsAsync(int entryId)
        {
            return await _dataAccess.GetUsageDetailsAsync(entryId);
        }
    }
}
