using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public class MakerLoanTransferService : IMakerLoanTransferService
    {
        private readonly IMakerLoanTransferDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public MakerLoanTransferService(
            IMakerLoanTransferDataAccess dataAccess,
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

        public async Task SaveTransferAsync(MakerLoanTransferDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.AccNo))
                throw new InvalidOperationException("Please select a valid Maker account.");

            if (dto.LTToClear <= 0 && dto.STToClear <= 0)
                throw new InvalidOperationException("Please enter a valid transfer amount for L.T to S.T or S.T to L.T.");

            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            await _dataAccess.SaveTransferAsync(dto, userName, machineName);
        }
    }
}
