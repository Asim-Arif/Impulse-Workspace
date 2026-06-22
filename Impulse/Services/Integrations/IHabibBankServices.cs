using DataAccessLibrary.Models.ViewModels.Accounts;

namespace Impulse.Services.Integrations
{
    public interface IHabibBankServices
    {
        Task<int> ExportExcelAsync(int entryId);
    }
}
