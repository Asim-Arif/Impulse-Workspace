using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class PostFineService : IPostFineService
    {
        private readonly IPostFineDataAccess _dataAccess;

        public PostFineService(IPostFineDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task SaveFineAsync(PostFineDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.EmpID))
            {
                throw new InvalidOperationException("Please select a valid employee.");
            }
            if (dto.Amount <= 0)
            {
                throw new InvalidOperationException("Fine Amount must be greater than zero.");
            }

            await _dataAccess.SaveFineAsync(dto);
        }
    }
}
