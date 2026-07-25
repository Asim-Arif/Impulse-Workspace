using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IRMService
    {
        Task<byte[]> GetMaterialPictureAsync(string materialId);
    }
}
