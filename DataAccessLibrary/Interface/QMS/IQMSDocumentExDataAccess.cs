using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface IQMSDocumentExDataAccess
    {
        Task<List<QMSDocumentExListItemDto>> GetDocumentsAsync(string masterCategory, string? searchText = null, DateTime? fromDate = null, DateTime? toDate = null);
        Task<QMSDocumentExModel?> GetDocumentByIdAsync(int entryId);
        Task<List<QMSDocumentExDistributionItemModel>> GetDistributionListAsync(int docEntryId);
        Task<List<string>> GetDocumentTypesByMasterAsync(string masterCategory);
        Task<List<EmployeeDropdownDto>> GetEmployeesAsync();
        Task<int> SaveDocumentAsync(QMSDocumentExModel model, List<QMSDocumentExDistributionItemModel> distList);
        Task<bool> DeleteDocumentAsync(int entryId);
        Task<string> GetNextDocumentNumberAsync();
    }
}
