using DataAccessLibrary.Models.ViewModels.Accounts;

namespace CIP.Services
{
    public interface IHabibBankServices
    {
        Task<int> ExportExcelAsync(int entryId);
    }
}
