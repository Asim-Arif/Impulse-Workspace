using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface IQMSDocumentDataAccess
    {
        Task<List<QMSDocumentListItemModel>> GetDocumentsAsync(QMSDocumentFilterDto filter);
        Task<QMSDocumentModel?> GetDocumentByIdAsync(int entryId);
        Task<int> SaveDocumentAsync(QMSDocumentModel model, List<int> deletedDistributionIds, string currentUserName);
        Task<bool> DeleteDocumentAsync(int entryId);
        Task<List<EmployeeDropdownDto>> GetEmployeesAsync();
        Task<List<string>> GetDocumentTypesListAsync();
    }
}

