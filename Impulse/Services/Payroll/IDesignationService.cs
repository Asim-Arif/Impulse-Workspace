using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IDesignationService
    {
        Task<List<DesignationModel>> GetDesignationsAsync();
        Task<DesignationModel?> GetDesignationByIdAsync(int entryId);
        Task<bool> SaveDesignationAsync(DesignationModel model);
        Task<bool> DeleteDesignationAsync(int entryId);
    }
}
