using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface ISupplierReEvaluationDataAccess
    {
        Task<List<SupplierReEvaluationListItemDto>> GetSupplierReEvaluationsAsync(SupplierReEvaluationFilterDto filter);
        Task<SupplierReEvaluationModel?> GetReEvaluationByIdAsync(int entryId);
        Task<SupplierReEvaluationModel?> GetLatestReEvaluationByAccNoAsync(string accNo);
        Task<int> SaveReEvaluationAsync(SupplierReEvaluationModel model);
        Task<bool> DeleteReEvaluationAsync(int entryId);
    }
}

