using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class ShortTermSheetService : IShortTermSheetService
    {
        private readonly IShortTermSheetDataAccess _dataAccess;
        private readonly IVoucherService _voucherService;
        private readonly IDBHelperService _dbHelperService;
        private readonly IAuditService _auditService;

        public ShortTermSheetService(
            IShortTermSheetDataAccess dataAccess,
            IVoucherService voucherService,
            IDBHelperService dbHelperService,
            IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _voucherService = voucherService;
            _dbHelperService = dbHelperService;
            _auditService = auditService;
        }

        public Task<List<ShortTermSheetItemViewModel>> GetShortTermSheetDataAsync(DateTime date, string deptId)
        {
            return _dataAccess.GetShortTermSheetDataAsync(date, deptId);
        }

        public Task SaveShortTermSheetEntriesAsync(DateTime date, IEnumerable<ShortTermSheetSaveDto> items)
        {
            return _dataAccess.SaveShortTermSheetEntriesAsync(date, items);
        }

        public Task<ShortTermPostingStatusModel> GetPostingStatusAsync(DateTime date, string deptId)
        {
            return _dataAccess.GetPostingStatusAsync(date, deptId);
        }

        public async Task<string> PostVouchersAndAdvancesAsync(ShortTermPostingDto dto)
        {
            var approvedItems = dto.Items.Where(x => x.Approved > 0).ToList();
            if (!approvedItems.Any())
            {
                throw new InvalidOperationException("No approved advance entries available to post.");
            }

            decimal totalAmount = approvedItems.Sum(x => x.Approved);

            string empSTAccNo = await _dbHelperService.getSingleStringValue(
                "DataValue", "GeneralData", "WHERE DataName = 'EmpShortTermAccNo'");

            if (string.IsNullOrWhiteSpace(empSTAccNo))
            {
                throw new InvalidOperationException(
                    "Employee Short-Term Loan Account (EmpShortTermAccNo) is not configured in General Data. " +
                    "Please contact your system administrator.");
            }

            string vchrNo = string.Empty;
            string bankVchrNo = string.Empty;

            if (!dto.DontPostVoucher)
            {
                if (string.IsNullOrWhiteSpace(dto.PaymentAccNo))
                {
                    throw new InvalidOperationException("Please select a valid payment account.");
                }

                string vchrType = dto.PaymentMode == ShortTermPostingMode.Cash ? "CPV"
                                : dto.PaymentMode == ShortTermPostingMode.Bank ? "BPV" : "JV";

                vchrNo = await _voucherService.GetNextVchrNo(dto.Date, vchrType);
                string description = $"Advance Salary Short Term — Dept: {dto.DeptID} Month: {dto.Date:MMM yyyy}";

                var voucher = new VoucherViewModel
                {
                    VchrNo = vchrNo,
                    DT = dto.Date,
                    UserName = _auditService.GetCurrentUserName(),
                    MachineName = _auditService.GetClientIpAddress(),
                    Notes = description,
                    LineItems = new List<VoucherLineItemViewModel>
                    {
                        new VoucherLineItemViewModel
                        {
                            AccNo = empSTAccNo,
                            Description = description,
                            Debit = totalAmount,
                            Credit = 0
                        },
                        new VoucherLineItemViewModel
                        {
                            AccNo = dto.PaymentAccNo,
                            Description = description,
                            Debit = 0,
                            Credit = totalAmount
                        }
                    }
                };

                if (dto.PaymentMode == ShortTermPostingMode.Bank)
                {
                    bankVchrNo = vchrNo;
                    int.TryParse(dto.BankId, out int bankIdVal);
                    voucher.ChequeDetails = new ChequeViewModel
                    {
                        CDate = dto.ChequeDate,
                        AccNo = dto.PaymentAccNo,
                        BankAccNo = dto.PaymentAccNo,
                        BankID = bankIdVal,
                        ChequeNo = dto.ChequeNo,
                        chequeType = dto.ChequeType ?? "Open",
                        ChequeDate = dto.ChequeDate,

                        Amount = totalAmount,
                        Description = description,
                        Payment = true,
                        Posted = false,
                        Bounced = false
                    };
                }



                await _voucherService.SaveVoucherAsync(voucher);
            }

            await _dataAccess.SavePostingRecordAsync(
                dto.Date,
                dto.DeptID,
                vchrNo,
                bankVchrNo,
                dto.ChequeNo,
                dto.ChequeDate,
                (int)dto.PaymentMode,
                _auditService.GetCurrentUserName(),
                _auditService.GetClientIpAddress(),
                approvedItems);

            return vchrNo;
        }
    }
}

