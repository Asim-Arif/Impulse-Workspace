using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;
using static DataAccessLibrary.Models.ViewModels.Accounts.BankListModel;

public interface IBankListDataAccess
{
    //Task AddBankAsync(BankListModel bank);
    Task UpdateBankAsync(BankListModel bank);
    Task DeleteBankAsync(int bankID);
    Task EditBankAsync(BankListModel bank,int bankid);
    Task<List<BankListModel>> GetBankList();
    Task SaveNewBank(BankListModel newbank);

}
