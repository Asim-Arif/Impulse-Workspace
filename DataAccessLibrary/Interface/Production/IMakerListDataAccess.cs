using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Production
{
    public interface IMakerListDataAccess
    {
        Task<List<MakerDto>> GetMakersListAsync(bool includeInactive);
        Task ToggleMakerActiveStatusAsync(long vendId, bool newActiveState);
    }
}
