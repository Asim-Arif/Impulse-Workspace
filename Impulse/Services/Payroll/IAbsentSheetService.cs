using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IAbsentSheetService
    {
        Task<AbsentSheetResultModel> GetAbsentSheetAsync(AbsentSheetFilterDto filter);
    }
}
