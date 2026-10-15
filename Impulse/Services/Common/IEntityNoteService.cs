using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Attachments;

namespace Impulse.Services.Common
{
    public interface IEntityNoteService
    {
        Task<EntityNoteDto?> GetNoteAsync(string entityType, string entityId, string entitySecondaryId = null);
        Task<bool> SaveNoteAsync(string entityType, string entityId, string entitySecondaryId, string noteText, string user = null);
    }
}
