using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IPostFineService
    {
        /// <summary>
        /// Saves an employee fine record into the database.
        /// </summary>
        Task SaveFineAsync(PostFineDto dto);
    }
}
