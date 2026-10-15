using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Attachments;

namespace DataAccessLibrary.Interface.Common
{
    public interface IEntityNoteDataAccess
    {
        Task<EntityNoteDto?> GetNoteAsync(string entityType, string entityId, string? secondaryId = null);
        Task<bool> SaveNoteAsync(string entityType, string entityId, string? secondaryId, string noteText, string user);
    }
}
