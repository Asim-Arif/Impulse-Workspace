using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IVendorRMAssignmentService
    {
        Task<List<NewVendorViewModel>> GetVendorsAsync();
        Task<List<UnassignedRMLookupModel>> GetUnassignedRMsAsync(string vendId);
        Task<List<VendorRMAssignmentModel>> GetAssignedRMsAsync(string vendId);
        Task AssignRMAsync(string vendId, int rmId);
        Task UnassignRMAsync(int entryId);
        Task UpdateRMAssignmentAsync(int entryId, decimal rate, string remarks, string userName);
        Task<List<VendorRMAssignmentHistoryModel>> GetHistoryAsync(int entryId);
    }
}
