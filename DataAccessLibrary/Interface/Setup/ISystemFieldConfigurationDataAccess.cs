using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.Setup;

namespace DataAccessLibrary.Interface.Setup
{
    public interface ISystemFieldConfigurationDataAccess
    {
        Task<List<SystemFieldConfigurationModel>> GetConfigurationsByFormAsync(string module, string formKey);
        Task<List<SystemFieldConfigurationModel>> GetAllConfigurationsAsync();
        Task<bool> UpdateConfigurationAsync(int settingId, bool isRequired, bool isVisible);
    }
}

