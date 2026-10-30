using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface IQMSReasonDataAccess
    {
        Task<List<QMSReasonModel>> GetAllReasonsAsync(string? searchText = null);
        Task<QMSReasonModel?> GetReasonByIdAsync(int entryId);
        Task<bool> ReasonExistsAsync(string reason, int? excludeEntryId = null);
        Task<int> AddReasonAsync(string reason);
        Task<bool> UpdateReasonAsync(int entryId, string originalReason, string newReason);
        Task<bool> DeleteReasonAsync(int entryId, string reason);
        Task<int> GetUsageCountAsync(string reason);
    }
}
