using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class AdjustDedAmtService : IAdjustDedAmtService
    {
        private readonly IAdjustDedAmtDataAccess _dataAccess;

        public AdjustDedAmtService(IAdjustDedAmtDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<EmpDeductionInfoModel> GetEmployeeLatestLoanDeductionAsync(string empId)
        {
            return await _dataAccess.GetEmployeeLatestLoanDeductionAsync(empId);
        }

        public async Task UpdateDeductionAmountAsync(AdjustDedAmtDto dto)
        {
            if (dto.EntryID <= 0)
            {
                throw new InvalidOperationException("Invalid Loan Record selected for adjustment.");
            }

            await _dataAccess.UpdateDeductionAmountAsync(dto.EntryID, dto.NewDAmount);
        }
    }
}
