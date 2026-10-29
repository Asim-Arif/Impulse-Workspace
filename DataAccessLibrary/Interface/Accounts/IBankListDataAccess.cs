using DataAccessLibrary.Models.ViewModels.Accounts;
using System.Collections.Generic;
using System.Threading.Tasks;

public interface IBankListDataAccess
{
    // Modern Hierarchical API
    Task<List<BankBranchRecord>> GetAllBankBranchesAsync();
    Task<List<BankAccountItem>> GetAllBankAccountsAsync();
    Task<List<BankAccountItem>> GetBankAccountsByBankIdAsync(int bankId);
    Task<BankAccountItem?> GetBankAccountByAccNoAsync(string accNo);
    Task<BankBranchRecord?> GetBankBranchByIdAsync(int bankId);
    Task<BankOperationResult> SaveBankBranchAsync(SaveBankBranchDto dto, string userName, string machineName);
    Task<BankOperationResult> SaveBankAccountAsync(SaveBankAccountDto dto, string userName, string machineName);
    Task<BankOperationResult> SetDefaultBankAccountAsync(string accNo);
    Task<BankOperationResult> UpdateOverdraftAsync(OverdraftConfigDto dto);
    Task<BankOperationResult> DeleteBankBranchAsync(int bankId);
    Task<BankOperationResult> DeleteBankAccountAsync(string accNo);
    Task<List<string>> GetDistinctBankNamesAsync();
    Task<List<string>> GetDistinctCitiesAsync();

    // Legacy compatibility
    Task<List<BankListModel>> GetBankList();
    Task SaveNewBank(BankListModel newbank);
    Task EditBankAsync(BankListModel bank, int bankid);
    Task DeleteBankAsync(int bankID);
    Task UpdateBankAsync(BankListModel bank);
}
