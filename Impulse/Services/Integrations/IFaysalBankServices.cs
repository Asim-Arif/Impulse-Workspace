using DataAccessLibrary.Models.ViewModels.Accounts;

namespace CIP.Services
{
    public interface IFaysalBankServices
    {
        //Task<int> ExportAndUploadExcelAsync(int entryId);
        Task<int> ExportAndUploadExcelForBOD(int entryId);
    }
}
