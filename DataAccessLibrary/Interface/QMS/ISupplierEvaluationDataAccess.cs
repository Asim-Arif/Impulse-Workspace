using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface ISupplierEvaluationDataAccess
    {
        Task<List<SupplierEvaluationListItemDto>> GetSupplierEvaluationsAsync(SupplierEvaluationFilterDto filter);
        Task<SupplierEvaluationModel?> GetEvaluationByAccNoAsync(string accNo);
        Task<SupplierEvaluationModel?> GetEvaluationByIdAsync(int entryId);
        Task<List<SupplierDropdownDto>> GetSuppliersListAsync();
        Task<bool> SaveEvaluationAsync(SupplierEvaluationModel model);
        Task<bool> DeleteEvaluationAsync(string accNo);
    }
}

