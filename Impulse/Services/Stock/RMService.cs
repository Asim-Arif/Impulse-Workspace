using DataAccessLibrary.Interface.Stock;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class RMService : IRMService
    {
        private readonly IRMDataAccess _dataAccess;

        public RMService(IRMDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<byte[]> GetMaterialPictureAsync(string materialId)
        {
            return _dataAccess.GetMaterialPictureAsync(materialId);
        }
    }
}
