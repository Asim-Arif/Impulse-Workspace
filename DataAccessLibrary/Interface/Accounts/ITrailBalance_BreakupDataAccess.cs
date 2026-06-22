using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;
using static DataAccessLibrary.Models.ViewModels.Accounts.MainHeadsModel;

public interface ITrailBalance_BreakupDataAccess
{
    Task<List<MainHeadsModel>> GetMainHeadsList();
    Task<List<TrailBalance_BreakupModel>> GetAccountsList(DateTime fromDate, DateTime toDate, bool isSummary, string accountCode = "");
    //Task<TrailBalance_BreakupModel> CalculateAccountBalance(string accNo, DateTime fromDate, DateTime toDate, bool isParent);
    //Task<TrailBalance_BreakupModel> CalculateAccountBalances(string accNo, DateTime fromDate, DateTime toDate);
    // CORRECT: Returns a list of items
    Task<List<TrailBalance_BreakupModel>> CalculateAccountBalances(string accNo, DateTime fromDate, DateTime toDate);
    Task SaveLoadedDataToTempTable(List<TrailBalance_BreakupModel> modeldata, bool bWith_Opening_Closing);

}