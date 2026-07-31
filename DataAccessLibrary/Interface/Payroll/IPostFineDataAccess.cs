using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IPostFineDataAccess
    {
        /// <summary>
        /// Inserts a new fine record into EmpFine table.
        /// </summary>
        Task SaveFineAsync(PostFineDto dto);
    }
}
