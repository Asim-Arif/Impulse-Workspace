using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IPpcMakerOrderDataAccess
    {
        Task EnsureSchemaAsync();
        Task<PpcMakerPoOrderHeaderDto?> GetOrderHeaderAsync(string orderNo);
        Task<List<PpcMakerPoOrderHeaderDto>> GetActiveOrdersWithPendingPurchasesAsync();
        Task<List<PpcMakerPoItemRowDto>> GetPpcPurchasesForOrderAsync(string orderNo);
        Task<List<MakerPOLookupModel>> GetAllMakersAsync();
        Task<GenerateMakerPoResult> GenerateMakerPosAsync(GenerateMakerPoRequest request);
    }
}
