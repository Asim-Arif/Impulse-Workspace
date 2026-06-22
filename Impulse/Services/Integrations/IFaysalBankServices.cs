using DataAccessLibrary.Models.ViewModels.Accounts;

namespace Impulse.Services.Integrations
{
    public interface IFaysalBankServices
    {
        //Task<int> ExportAndUploadExcelAsync(int entryId);
        Task<int> ExportAndUploadExcelForBOD(int entryId);
    }
}
