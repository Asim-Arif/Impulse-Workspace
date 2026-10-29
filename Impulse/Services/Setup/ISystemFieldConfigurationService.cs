using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace Impulse.Services.Setup
{
    public interface ISystemFieldConfigurationService
    {
        Task<Dictionary<string, SystemFieldConfigurationModel>> GetFormConfigurationsAsync(string module, string formKey);
        Task<bool> IsFieldRequiredAsync(string module, string formKey, string fieldKey, bool defaultIfMissing = true);
        Task<bool> IsFieldVisibleAsync(string module, string formKey, string fieldKey, bool defaultIfMissing = true);
        Task<List<SystemFieldConfigurationModel>> GetAllConfigurationsAsync();
        Task<bool> UpdateConfigurationAsync(int settingId, bool isRequired, bool isVisible);
    }
}

