using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components.Forms;

namespace CIP.Services
{
    public interface ICommonServices
    {
        Task<int> ProcessBODMeetingAttendanceExcel(IBrowserFile file, DateTime meetingDate);
        string GetValueFromXMLString(string xml, string tagName);
    }
}
