using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface IQMSDocumentTypeDataAccess
    {
        Task<List<QMSDocumentTypeModel>> GetDocumentTypesAsync(string? searchTerm = null);
        Task<QMSDocumentTypeModel?> GetDocumentTypeByIdAsync(int entryId);
        Task<bool> DocumentTypeExistsAsync(string documentType, int? excludeEntryId = null);
        Task<int> AddDocumentTypeAsync(string documentType);
        Task<bool> UpdateDocumentTypeAsync(int entryId, string newDocumentType);
        Task<bool> DeleteDocumentTypeAsync(int entryId);
        Task<int> GetUsageCountAsync(int entryId);
    }
}

