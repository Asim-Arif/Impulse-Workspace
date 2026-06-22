using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;

namespace DataAccessLibrary.Interface
    
{
    public interface ICommonDataAccess
    {
        Task<int> GetOrCreateBODMeetingID(DateTime meetingDate);
        Task SaveBODMeetingParticipant(int meetingID, int folioNo, SIALDirectorViewModel director);
        Task<SIALDirectorViewModel?> GetDirectorByFolio(int folioNo);
    }
}
