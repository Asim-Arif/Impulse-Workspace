using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public interface IMakerListService
    {
        Task<List<MakerDto>> GetMakersListAsync(bool includeInactive);
        Task ToggleMakerActiveStatusAsync(long vendId, bool newActiveState);
    }
}
