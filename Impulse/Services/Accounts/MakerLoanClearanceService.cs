using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public class MakerLoanClearanceService : IMakerLoanClearanceService
    {
        private readonly IMakerLoanClearanceDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public MakerLoanClearanceService(
            IMakerLoanClearanceDataAccess dataAccess,
            IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public Task<List<MakerAccountModel>> GetMakerAccountsAsync()
        {
            return _dataAccess.GetMakerAccountsAsync();
        }

        public Task<MakerLoanBalanceModel> GetMakerBalancesAsync(string accNo)
        {
            return _dataAccess.GetMakerBalancesAsync(accNo);
        }

        public async Task SaveClearanceAsync(MakerLoanClearanceDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.AccNo))
                throw new InvalidOperationException("Please select a valid Maker account.");

            if (dto.LTToClear <= 0 && dto.STToClear <= 0)
                throw new InvalidOperationException("Please enter a valid Long Term or Short Term clearance amount.");

            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.SaveClearanceAsync(dto, userName, machineName);
        }

        public async Task UpdateLTDeductionAmountAsync(string accNo, decimal newDAmount)
        {
            if (string.IsNullOrWhiteSpace(accNo))
                throw new InvalidOperationException("Please select a valid Maker account.");

            if (newDAmount < 0)
                throw new InvalidOperationException("Deduction amount cannot be negative.");

            await _dataAccess.UpdateLTDeductionAmountAsync(accNo, newDAmount);
        }
    }
}
