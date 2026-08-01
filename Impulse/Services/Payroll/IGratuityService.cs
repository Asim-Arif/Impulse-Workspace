using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IGratuityService
    {
        Task<GratuitySheetModel> GetGratuitySheetAsync(int year, string? deptId);
        Task<bool> SaveGratuitySheetAsync(SaveGratuityDto dto);
    }
}
