using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;

namespace Impulse.Services.Payroll
{
    public class HoldSalaryService : IHoldSalaryService
    {
        private readonly IHoldSalaryDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public HoldSalaryService(IHoldSalaryDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public async Task<HoldSalarySummaryModel> GetEmployeeHoldSalaryInfoAsync(string empId, int year, int month)
        {
            if (string.IsNullOrWhiteSpace(empId))
            {
                return new HoldSalarySummaryModel();
            }

            return await _dataAccess.GetEmployeeHoldSalaryInfoAsync(empId, year, month);
        }

        public async Task<List<HoldSalaryListItemModel>> GetHoldSalariesAsync(string? deptId, int year, int month)
        {
            return await _dataAccess.GetHoldSalariesAsync(deptId, year, month);
        }

        public async Task<bool> SaveHoldSalaryAsync(HoldSalaryDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.EmpID))
            {
                throw new InvalidOperationException("Please select a valid employee.");
            }

            if (dto.HoldAmt <= 0)
            {
                throw new InvalidOperationException("Invalid hold amount. Amount must be greater than zero.");
            }

            var summary = await _dataAccess.GetEmployeeHoldSalaryInfoAsync(dto.EmpID, dto.Year, dto.Month);

            if (dto.HoldAmt > summary.AvailableBalance)
            {
                throw new InvalidOperationException($"Hold amount ({dto.HoldAmt:N0}) cannot be more than the available salary balance ({summary.AvailableBalance:N0}).");
            }

            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();

            return await _dataAccess.SaveHoldSalaryAsync(dto, userName, machineName);
        }

        public async Task<bool> DeleteHoldSalaryAsync(int entryId)
        {
            return await _dataAccess.DeleteHoldSalaryAsync(entryId);
        }
    }
}
