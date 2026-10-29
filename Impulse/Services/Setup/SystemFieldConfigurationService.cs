using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;

namespace Impulse.Services.Setup
{
    public class SystemFieldConfigurationService : ISystemFieldConfigurationService
    {
        private readonly ISystemFieldConfigurationDataAccess _dataAccess;
        private static readonly ConcurrentDictionary<string, (DateTime CachedAt, Dictionary<string, SystemFieldConfigurationModel> Configs)> _cache = new(StringComparer.OrdinalIgnoreCase);
        private static readonly TimeSpan CacheDuration = TimeSpan.FromMinutes(10);

        public SystemFieldConfigurationService(ISystemFieldConfigurationDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<Dictionary<string, SystemFieldConfigurationModel>> GetFormConfigurationsAsync(string module, string formKey)
        {
            string cacheKey = $"{module}:{formKey}";

            if (_cache.TryGetValue(cacheKey, out var cached) && (DateTime.UtcNow - cached.CachedAt) < CacheDuration)
            {
                return cached.Configs;
            }

            var list = await _dataAccess.GetConfigurationsByFormAsync(module, formKey);
            var dict = list.ToDictionary(k => k.FieldKey, v => v, StringComparer.OrdinalIgnoreCase);

            _cache[cacheKey] = (DateTime.UtcNow, dict);
            return dict;
        }

        public async Task<bool> IsFieldRequiredAsync(string module, string formKey, string fieldKey, bool defaultIfMissing = true)
        {
            var configs = await GetFormConfigurationsAsync(module, formKey);
            if (configs.TryGetValue(fieldKey, out var model))
            {
                return model.IsRequired;
            }
            return defaultIfMissing;
        }

        public async Task<bool> IsFieldVisibleAsync(string module, string formKey, string fieldKey, bool defaultIfMissing = true)
        {
            var configs = await GetFormConfigurationsAsync(module, formKey);
            if (configs.TryGetValue(fieldKey, out var model))
            {
                return model.IsVisible;
            }
            return defaultIfMissing;
        }

        public Task<List<SystemFieldConfigurationModel>> GetAllConfigurationsAsync()
        {
            return _dataAccess.GetAllConfigurationsAsync();
        }

        public async Task<bool> UpdateConfigurationAsync(int settingId, bool isRequired, bool isVisible)
        {
            bool success = await _dataAccess.UpdateConfigurationAsync(settingId, isRequired, isVisible);
            if (success)
            {
                _cache.Clear(); // Invalidate cache so changes take effect immediately
            }
            return success;
        }
    }
}

