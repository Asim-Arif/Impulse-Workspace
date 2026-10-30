using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.QMS
{
    public class CalibrationDataAccess : ICalibrationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<CalibrationDataAccess> _logger;

        public CalibrationDataAccess(IConfiguration configuration, ILogger<CalibrationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<CalibrationListItemModel>> GetCalibrationsAsync(CalibrationFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();
                var sql = new StringBuilder(@"
                    SELECT 
                        c.EntryID, c.AssetID, a.Description, c.DT, c.CertificateNo, c.MasterReference,
                        c.AccNo, v.AccTitle, c.Remarks, c.NextDueDT, c.UserName, c.DTEntry,
                        (SELECT COUNT(1) FROM EntityAttachments ea WHERE ea.EntityType = 'Calibration' AND ea.EntityId = CAST(c.EntryID AS VARCHAR(50))) AS AttachmentCount
                    FROM FixedAssets_Calibration c
                    LEFT JOIN FixedAssets a ON c.AssetID = a.AssetID
                    LEFT JOIN VVenders v ON c.AccNo = v.AccNo
                    WHERE 1 = 1");

                var parameters = new DynamicParameters();

                if (filter != null)
                {
                    if (filter.DateFrom.HasValue && filter.DateTo.HasValue)
                    {
                        sql.Append(" AND c.DT >= @DateFrom AND c.DT <= @DateTo");
                        parameters.Add("DateFrom", filter.DateFrom.Value.Date);
                        parameters.Add("DateTo", filter.DateTo.Value.Date.AddDays(1).AddTicks(-1));
                    }

                    if (!string.IsNullOrWhiteSpace(filter.AssetID) && filter.AssetID != "0")
                    {
                        sql.Append(" AND c.AssetID = @AssetID");
                        parameters.Add("AssetID", filter.AssetID.Trim());
                    }

                    if (!string.IsNullOrWhiteSpace(filter.SearchTerm))
                    {
                        sql.Append(@" AND (
                            c.CertificateNo LIKE @Search 
                            OR c.AssetID LIKE @Search 
                            OR a.Description LIKE @Search 
                            OR c.Remarks LIKE @Search 
                            OR c.MasterReference LIKE @Search
                            OR v.AccTitle LIKE @Search
                        )");
                        parameters.Add("Search", $"%{filter.SearchTerm.Trim()}%");
                    }
                }

                sql.Append(" ORDER BY c.DT DESC, c.EntryID DESC");

                var results = await db.QueryAsync<CalibrationListItemModel>(sql.ToString(), parameters);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching calibrations list");
                throw;
            }
        }

        public async Task<CalibrationModel?> GetCalibrationByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        c.EntryID, c.AssetID, a.Description AS AssetDescription, a.CalibrationFrequency, c.DT, c.NextDueDT,
                        c.CertificateNo, c.MasterReference, c.AccNo, v.AccTitle, c.Remarks,
                        c.UserName, c.MachineName, c.DTEntry
                    FROM FixedAssets_Calibration c
                    LEFT JOIN FixedAssets a ON c.AssetID = a.AssetID
                    LEFT JOIN VVenders v ON c.AccNo = v.AccNo
                    WHERE c.EntryID = @EntryID;";

                return await db.QueryFirstOrDefaultAsync<CalibrationModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching calibration by ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<int> SaveCalibrationAsync(CalibrationModel model, string currentUserName, string? machineName)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();

            try
            {
                int calibrationId = model.EntryID;

                if (calibrationId == 0)
                {
                    const string insertSql = @"
                        INSERT INTO FixedAssets_Calibration (
                            AssetID, DT, Remarks, NextDueDT, UserName, MachineName, DTEntry, CertificateNo, MasterReference, AccNo
                        ) VALUES (
                            @AssetID, @DT, @Remarks, @NextDueDT, @UserName, @MachineName, GETDATE(), @CertificateNo, @MasterReference, @AccNo
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    calibrationId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.AssetID,
                        model.DT,
                        model.Remarks,
                        model.NextDueDT,
                        UserName = currentUserName,
                        MachineName = machineName ?? Environment.MachineName,
                        model.CertificateNo,
                        model.MasterReference,
                        model.AccNo
                    });
                }
                else
                {
                    const string updateSql = @"
                        UPDATE FixedAssets_Calibration SET
                            AssetID = @AssetID,
                            DT = @DT,
                            Remarks = @Remarks,
                            NextDueDT = @NextDueDT,
                            CertificateNo = @CertificateNo,
                            MasterReference = @MasterReference,
                            AccNo = @AccNo
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.AssetID,
                        model.DT,
                        model.Remarks,
                        model.NextDueDT,
                        model.CertificateNo,
                        model.MasterReference,
                        model.AccNo,
                        model.EntryID
                    });
                }

                return calibrationId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving calibration EntryID {EntryID}", model.EntryID);
                throw;
            }
        }

        public async Task<bool> DeleteCalibrationAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM FixedAssets_Calibration WHERE EntryID = @EntryID;";
                int affected = await db.ExecuteAsync(sql, new { EntryID = entryId });
                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting calibration EntryID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<List<FixedAssetLookupDto>> GetFixedAssetsAsync(bool requiresCalibrationOnly = true)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT AssetID, Description, RequiresCalibration, CalibrationFrequency
                    FROM FixedAssets
                    WHERE (@RequiresOnly = 0 OR RequiresCalibration = 1)
                    ORDER BY AssetID;";

                var result = await db.QueryAsync<FixedAssetLookupDto>(sql, new { RequiresOnly = requiresCalibrationOnly ? 1 : 0 });
                var list = result.ToList();

                // If no assets have RequiresCalibration=1, return all assets so the user is never blocked
                if (list.Count == 0 && requiresCalibrationOnly)
                {
                    const string fallbackSql = @"
                        SELECT AssetID, Description, RequiresCalibration, CalibrationFrequency
                        FROM FixedAssets
                        ORDER BY AssetID;";
                    var all = await db.QueryAsync<FixedAssetLookupDto>(fallbackSql);
                    return all.ToList();
                }

                return list;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching fixed assets lookup");
                return new List<FixedAssetLookupDto>();
            }
        }

        public async Task<List<VendorLookupDto>> GetVendorsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT AccNo, AccTitle
                    FROM VVenders
                    ORDER BY AccTitle;";

                var result = await db.QueryAsync<VendorLookupDto>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching vendors lookup");
                return new List<VendorLookupDto>();
            }
        }
    }
}

