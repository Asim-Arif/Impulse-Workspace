using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using CIP.Pages.Accounts;
using System.Security.Cryptography.X509Certificates;

namespace CIP.Services
{
    public interface ITrailBalance_BreakupService
    {
        Task<List<MainHeadsModel>> GetMainHeadsList();
        //Task<List<TrailBalance_BreakupModel>> GetAccountsList(DateTime fromDate, DateTime toDate, bool isSummary, string accountCode = "");
        //Task<TrailBalance_BreakupModel> CalculateAccountBalance(string accNo, DateTime fromDate, DateTime toDate, bool isParent);
        Task<List<TrailBalance_BreakupModel>> GenerateReport(DateTime fromDate, DateTime toDate, bool isSummary);

        Task<List<TrailBalance_BreakupModel>> GenerateAndSaveReport(List<TrailBalance_BreakupModel> myReportList, bool bWith_Opening_Closing);
    }
}
