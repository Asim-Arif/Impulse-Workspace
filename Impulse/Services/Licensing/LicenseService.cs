using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using Impulse.Models.License;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Licensing
{
    public class LicenseService : ILicenseService
    {
        private readonly string _connectionString;
        private readonly IHardwareFingerprintService _fingerprintService;
        private readonly ILogger<LicenseService> _logger;
        private LicenseInfo _currentLicense = new();

        private DateTime _lastCheckUtc = DateTime.MinValue;
        private readonly object _lock = new();

        public LicenseInfo CurrentLicense
        {
            get
            {
                EnsureLicenseFresh();
                return _currentLicense;
            }
        }

        public bool IsLicensed
        {
            get
            {
                EnsureLicenseFresh();
                return _currentLicense.IsLicensed;
            }
        }

        public LicenseService(
            IConfiguration configuration,
            IHardwareFingerprintService fingerprintService,
            ILogger<LicenseService> logger)
        {
            _fingerprintService = fingerprintService;
            _logger = logger;
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("DefaultConnection not found in configuration.");

            // Initialize synchronously on startup
            try
            {
                ReloadLicenseInternal();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Initial license check failed on startup.");
            }
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public string GetCurrentHardwareId() => _fingerprintService.GetHardwareId();

        public bool IsModuleAllowed(AppModule module)
        {
            EnsureLicenseFresh();
            return _currentLicense.IsModuleAllowed(module);
        }

        public bool IsModuleAllowed(string moduleName)
        {
            EnsureLicenseFresh();
            return _currentLicense.IsModuleAllowed(moduleName);
        }

        public Task ReloadLicenseAsync()
        {
            lock (_lock)
            {
                ReloadLicenseInternal();
            }
            return Task.CompletedTask;
        }

        private void EnsureLicenseFresh()
        {
            var now = DateTime.UtcNow;
            if (_currentLicense.IsLicensed && (now - _lastCheckUtc).TotalMinutes < 5)
            {
                return;
            }

            if (!_currentLicense.IsLicensed && (now - _lastCheckUtc).TotalSeconds < 3)
            {
                return;
            }

            lock (_lock)
            {
                var checkNow = DateTime.UtcNow;
                if (_currentLicense.IsLicensed && (checkNow - _lastCheckUtc).TotalMinutes < 5)
                    return;
                if (!_currentLicense.IsLicensed && (checkNow - _lastCheckUtc).TotalSeconds < 3)
                    return;

                ReloadLicenseInternal();
            }
        }

        private void ReloadLicenseInternal()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT TOP 1 EntryID, LicenseKey, MachineName FROM License_Keys ORDER BY EntryID DESC";
                var record = db.QueryFirstOrDefault<LicenseDbRecord>(sql);

                if (record == null || string.IsNullOrWhiteSpace(record.LicenseKey))
                {
                    _currentLicense = new LicenseInfo
                    {
                        Status = LicenseStatus.Missing,
                        HardwareId = _fingerprintService.GetHardwareId(),
                        StatusMessage = "No license key found in database. System is unlicensed."
                    };
                    _logger.LogWarning("System is unlicensed: No key found in dbo.License_Keys.");
                    return;
                }

                var (isValid, payload, errorMsg) = LicenseCrypto.VerifyLicense(record.LicenseKey, LicenseKeys.PublicKeyXml);
                if (!isValid || payload == null)
                {
                    _currentLicense = new LicenseInfo
                    {
                        Status = LicenseStatus.InvalidSignature,
                        HardwareId = _fingerprintService.GetHardwareId(),
                        StatusMessage = $"Invalid digital signature: {errorMsg}"
                    };
                    _logger.LogError("License verification failed: {Error}", errorMsg);
                    return;
                }

                // Check Hardware Lock
                if (!_fingerprintService.MatchesHardwareId(payload.HardwareId))
                {
                    _currentLicense = new LicenseInfo
                    {
                        Status = LicenseStatus.HardwareMismatch,
                        ClientName = payload.ClientName,
                        HardwareId = _fingerprintService.GetHardwareId(),
                        StatusMessage = $"Hardware ID mismatch. License is bound to '{payload.HardwareId}', but this server is '{_fingerprintService.GetHardwareId()}'."
                    };
                    _logger.LogError("License hardware mismatch for client '{Client}': expected '{Exp}', actual '{Act}'",
                        payload.ClientName, payload.HardwareId, _fingerprintService.GetHardwareId());
                    return;
                }

                // Check Expiry Date
                if (payload.ExpiresAt.HasValue && payload.ExpiresAt.Value < DateTime.UtcNow)
                {
                    _currentLicense = new LicenseInfo
                    {
                        Status = LicenseStatus.Expired,
                        ClientName = payload.ClientName,
                        HardwareId = _fingerprintService.GetHardwareId(),
                        ExpiresAt = payload.ExpiresAt,
                        IssuedAt = payload.IssuedAt,
                        StatusMessage = $"License expired on {payload.ExpiresAt.Value:dd-MMM-yyyy}."
                    };
                    _logger.LogWarning("License for client '{Client}' expired on {Expiry}", payload.ClientName, payload.ExpiresAt);
                    return;
                }

                // Valid active license
                var allowedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (payload.AllowedModules != null)
                {
                    foreach (var m in payload.AllowedModules)
                    {
                        if (!string.IsNullOrWhiteSpace(m)) allowedSet.Add(m.Trim());
                    }
                }

                _currentLicense = new LicenseInfo
                {
                    Status = LicenseStatus.Valid,
                    ClientName = payload.ClientName,
                    HardwareId = payload.HardwareId,
                    AllowedModules = allowedSet,
                    ExpiresAt = payload.ExpiresAt,
                    IssuedAt = payload.IssuedAt,
                    StatusMessage = "Active"
                };

                _logger.LogInformation("License active for '{Client}' on hardware '{HWID}'. Modules: {Modules}",
                    payload.ClientName, payload.HardwareId, string.Join(", ", allowedSet));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to read or parse license from database.");
                _currentLicense = new LicenseInfo
                {
                    Status = LicenseStatus.Corrupted,
                    HardwareId = _fingerprintService.GetHardwareId(),
                    StatusMessage = $"Database/License error: {ex.Message}"
                };
            }
            finally
            {
                _lastCheckUtc = DateTime.UtcNow;
            }
        }

        public async Task<(bool Success, string Message)> ActivateLicenseAsync(string newLicenseKey)
        {
            if (string.IsNullOrWhiteSpace(newLicenseKey))
            {
                return (false, "License key cannot be empty.");
            }

            try
            {
                var (isValid, payload, errorMsg) = LicenseCrypto.VerifyLicense(newLicenseKey, LicenseKeys.PublicKeyXml);
                if (!isValid || payload == null)
                {
                    return (false, $"Invalid license key signature: {errorMsg}");
                }

                if (!_fingerprintService.MatchesHardwareId(payload.HardwareId))
                {
                    return (false, $"Hardware mismatch. License is registered for '{payload.HardwareId}', but this server is '{_fingerprintService.GetHardwareId()}'.");
                }

                if (payload.ExpiresAt.HasValue && payload.ExpiresAt.Value < DateTime.UtcNow)
                {
                    return (false, $"This license key has already expired on {payload.ExpiresAt.Value:dd-MMM-yyyy}.");
                }

                using var db = CreateConnection();
                const string insertSql = @"
                    INSERT INTO License_Keys (LicenseKey, MachineName)
                    VALUES (@LicenseKey, @MachineName);";

                await db.ExecuteAsync(insertSql, new
                {
                    LicenseKey = newLicenseKey.Trim(),
                    MachineName = Environment.MachineName
                });

                ReloadLicenseInternal();
                return (true, $"License successfully activated for '{payload.ClientName}'.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error activating license key.");
                return (false, $"Activation error: {ex.Message}");
            }
        }
    }
}
