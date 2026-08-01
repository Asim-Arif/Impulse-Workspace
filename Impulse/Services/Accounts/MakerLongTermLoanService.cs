using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Accounts
{
    public class MakerLongTermLoanService : IMakerLongTermLoanService
    {
        private readonly IMakerLongTermLoanDataAccess _dataAccess;
        private readonly IVoucherService _voucherService;
        private readonly IDBHelperService _dbHelperService;
        private readonly IAuditService _auditService;

        public MakerLongTermLoanService(
            IMakerLongTermLoanDataAccess dataAccess,
            IVoucherService voucherService,
            IDBHelperService dbHelperService,
            IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _voucherService = voucherService;
            _dbHelperService = dbHelperService;
            _auditService = auditService;
        }

        public Task<List<MakerAccountModel>> GetMakerAccountsAsync()
        {
            return _dataAccess.GetMakerAccountsAsync();
        }

        public Task<decimal> GetMakerBalanceAsync(string accNo)
        {
            return _dataAccess.GetMakerLongTermBalanceAsync(accNo);
        }

        public Task<List<MakerMasterPoModel>> GetMakerMasterPOsAsync(long vendId, string accNo)
        {
            return _dataAccess.GetMakerMasterPOsAsync(vendId, accNo);
        }

        public Task<MakerMasterPoModel> GetMasterPoDetailsAsync(string masterPoNo, bool isVendor)
        {
            return _dataAccess.GetMasterPoDetailsAsync(masterPoNo, isVendor);
        }

        public async Task<string> SaveLoanAsync(MakerLongTermLoanDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.AccNo))
                throw new InvalidOperationException("Please select a valid Maker account.");

            if (dto.Amount <= 0)
                throw new InvalidOperationException("Loan amount must be greater than zero.");

            string vchrNo = string.Empty;

            if (!dto.DontPostVoucher)
            {
                if (string.IsNullOrWhiteSpace(dto.PaymentAccNo))
                    throw new InvalidOperationException("Please select a valid payment account.");

                string vchrType = dto.PaymentMode == MakerLoanPaymentMode.Cash ? "CPV"
                                : dto.PaymentMode == MakerLoanPaymentMode.Bank ? "BPV" : "JV";

                vchrNo = await _voucherService.GetNextVchrNo(dto.IssueDate, vchrType);
                string makerDisplayName = !string.IsNullOrWhiteSpace(dto.MakerNo)
                    ? $"{dto.MakerNo} {dto.MakerName}"
                    : dto.MakerName;

                string description = $"Long Term Advance Paid To {makerDisplayName} Amount Paid :{dto.Amount} Ded. Amt : {dto.DAmount}";
                if (!string.IsNullOrWhiteSpace(dto.MasterPONo))
                {
                    description += $" Against PO #{dto.MasterPONo}, PO Value :{dto.POValue} , Already L.T Taken :{dto.LTTaken} , Already S.T Taken : {dto.STTaken}";
                }
                if (!string.IsNullOrWhiteSpace(dto.Description))
                {
                    description += $" — {dto.Description}";
                }

                var voucher = new VoucherViewModel
                {
                    VchrNo = vchrNo,
                    DT = dto.IssueDate,
                    UserName = _auditService.GetCurrentUserName(),
                    MachineName = _auditService.GetClientIpAddress(),
                    Notes = description,
                    LineItems = new List<VoucherLineItemViewModel>
                    {
                        new VoucherLineItemViewModel
                        {
                            AccNo = dto.AccNo,
                            Description = description,
                            Debit = dto.Amount,
                            Credit = 0
                        },
                        new VoucherLineItemViewModel
                        {
                            AccNo = dto.PaymentAccNo,
                            Description = description,
                            Debit = 0,
                            Credit = dto.Amount
                        }
                    }
                };

                if (dto.PaymentMode == MakerLoanPaymentMode.Bank)
                {
                    int.TryParse(dto.BankId, out int bankIdVal);
                    voucher.ChequeDetails = new ChequeViewModel
                    {
                        CDate = dto.IssueDate,
                        AccNo = dto.PaymentAccNo,
                        BankAccNo = dto.PaymentAccNo,
                        BankID = bankIdVal,
                        ChequeNo = dto.ChequeNo,
                        chequeType = dto.ChequeType ?? "Open",
                        ChequeDate = dto.ChequeDate,
                        Amount = dto.Amount,
                        Description = description,
                        Payment = true,
                        Posted = false,
                        Bounced = false
                    };
                }

                await _voucherService.SaveVoucherAsync(voucher);
            }

            await _dataAccess.SaveMakerLongTermLoanAsync(dto, vchrNo);
            return vchrNo;
        }
    }
}
