using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IEOBIService
    {
        Task<EOBISheetModel> GetEOBISheetAsync(EOBIFilterDto filter);
        Task<EOBISheetModel> GenerateEOBISheetAsync(EOBIFilterDto filter);
        Task<bool> SaveEOBISheetAsync(SaveEOBIDto dto);
    }
}
