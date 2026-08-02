using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerListService : IMakerListService
    {
        private readonly IMakerListDataAccess _makerListDataAccess;

        public MakerListService(IMakerListDataAccess makerListDataAccess)
        {
            _makerListDataAccess = makerListDataAccess;
        }

        public Task<List<MakerDto>> GetMakersListAsync(bool includeInactive)
        {
            return _makerListDataAccess.GetMakersListAsync(includeInactive);
        }

        public Task ToggleMakerActiveStatusAsync(long vendId, bool newActiveState)
        {
            return _makerListDataAccess.ToggleMakerActiveStatusAsync(vendId, newActiveState);
        }
    }
}
