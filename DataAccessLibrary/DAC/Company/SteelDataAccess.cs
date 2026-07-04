using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class SteelDataAccess : ISteelDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<SteelDataAccess> _logger;

        public SteelDataAccess(IConfiguration configuration, ILogger<SteelDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<SteelTypeNodeModel>> GetFullSteelTreeAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                var steels = await db.QueryAsync<SteelTypeNodeModel>(
                    "SELECT SteelID, ISNULL(SteelType, '') as SteelType, ISNULL(Description, '') as Description, ISNULL(StripeMinW, 0) as StripeMinW, ISNULL(StripeMaxW, 0) as StripeMaxW FROM SteelTypes ORDER BY SteelType");
                    
                var gages = await db.QueryAsync<SteelGageNodeModel>(
                    "SELECT GageID, SteelID, ISNULL(Gage, '') as Gage, ISNULL(GageUnit, '') as GageUnit FROM SteelGages ORDER BY CAST(SUBSTRING(GageID, CHARINDEX('-', GageID) + 1, 50) AS INT)");
                    
                var allComps = await db.QueryAsync<dynamic>(
                    "SELECT c.ID as CompID, c.Symbol, c.Name, stc.SteelID, stc.LVal as MinValue, stc.UVal as MaxValue FROM Compositions c JOIN SteelTypeComps stc ON c.ID = stc.CompID");

                var steelList = steels.ToList();
                var gageList = gages.ToList();
                var compList = allComps.ToList();

                foreach (var steel in steelList)
                {
                    steel.Gages = gageList.Where(g => g.SteelID == steel.SteelID).ToList();
                    
                    var compsForSteel = compList.Where(c => c.SteelID == steel.SteelID).Select(c => new SteelCompModel
                    {
                        CompID = c.CompID,
                        Symbol = c.Symbol,
                        Name = c.Name,
                        MinValue = (float?)c.MinValue,
                        MaxValue = (float?)c.MaxValue
                    }).ToList();
                    
                    steel.Compositions = compsForSteel;
                }

                return steelList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting steel tree.");
                throw;
            }
        }

        public async Task<List<SteelCompModel>> GetAllCompositionsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                var comps = await db.QueryAsync<SteelCompModel>("SELECT ID as CompID, Symbol, Name FROM Compositions ORDER BY Symbol");
                return comps.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all compositions.");
                throw;
            }
        }

        public async Task<List<string>> GetUnitsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                var units = await db.QueryAsync<string>("SELECT ISNULL(UnitName, '') FROM Units ORDER BY UnitName");
                return units.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting units.");
                throw;
            }
        }

        public async Task<bool> SaveSteelTypeAsync(SteelTypeNodeModel model)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                db.Open();
                using var transaction = db.BeginTransaction();
                
                try
                {
                    if (model.SteelID == 0)
                    {
                        // Insert new
                        var maxIdObj = await db.ExecuteScalarAsync<int?>("SELECT MAX(SteelID) FROM SteelTypes", transaction: transaction);
                        model.SteelID = (maxIdObj ?? 0) + 1;

                        string sql = "INSERT INTO SteelTypes (SteelID, SteelType, Description, StripeMinW, StripeMaxW) VALUES (@SteelID, @SteelType, @Description, @StripeMinW, @StripeMaxW)";
                        await db.ExecuteAsync(sql, model, transaction: transaction);
                    }
                    else
                    {
                        // Update existing
                        string sql = "UPDATE SteelTypes SET SteelType = @SteelType, Description = @Description, StripeMinW = @StripeMinW, StripeMaxW = @StripeMaxW WHERE SteelID = @SteelID";
                        await db.ExecuteAsync(sql, model, transaction: transaction);
                        
                        // Clear old compositions
                        await db.ExecuteAsync("DELETE FROM SteelTypeComps WHERE SteelID = @SteelID", new { SteelID = model.SteelID }, transaction: transaction);
                    }

                    // Insert compositions
                    if (model.Compositions != null && model.Compositions.Any())
                    {
                        foreach (var comp in model.Compositions)
                        {
                            if (comp.MinValue.HasValue || comp.MaxValue.HasValue)
                            {
                                string compSql = "INSERT INTO SteelTypeComps (SteelID, CompID, LVal, UVal) VALUES (@SteelID, @CompID, @LVal, @UVal)";
                                await db.ExecuteAsync(compSql, new { 
                                    SteelID = model.SteelID, 
                                    CompID = comp.CompID, 
                                    LVal = comp.MinValue ?? 0, 
                                    UVal = comp.MaxValue ?? 0 
                                }, transaction: transaction);
                            }
                        }
                    }

                    transaction.Commit();
                    return true;
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving steel type.");
                throw;
            }
        }

        public async Task<bool> DeleteSteelTypeAsync(int steelId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                db.Open();
                using var transaction = db.BeginTransaction();
                
                try
                {
                    await db.ExecuteAsync("DELETE FROM SteelTypeComps WHERE SteelID = @SteelId", new { SteelId = steelId }, transaction: transaction);
                    await db.ExecuteAsync("DELETE FROM SteelGages WHERE SteelID = @SteelId", new { SteelId = steelId }, transaction: transaction);
                    var rows = await db.ExecuteAsync("DELETE FROM SteelTypes WHERE SteelID = @SteelId", new { SteelId = steelId }, transaction: transaction);
                    
                    transaction.Commit();
                    return rows > 0;
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting steel type.");
                throw;
            }
        }

        public async Task<bool> SaveSteelGageAsync(SteelGageNodeModel model, bool isInsert)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                if (isInsert)
                {
                    // Generate new GageID (SteelID-Increment)
                    string maxQuery = "SELECT MAX(CAST(SUBSTRING(GageID, CHARINDEX('-', GageID) + 1, 50) AS INT)) FROM SteelGages WHERE SteelID = @SteelID";
                    var maxObj = await db.ExecuteScalarAsync<int?>(maxQuery, new { SteelID = model.SteelID });
                    int nextId = (maxObj ?? 0) + 1;
                    model.GageID = $"{model.SteelID}-{nextId}";

                    string sql = "INSERT INTO SteelGages (GageID, SteelID, Gage, GageUnit) VALUES (@GageID, @SteelID, @Gage, @GageUnit)";
                    var rows = await db.ExecuteAsync(sql, model);
                    return rows > 0;
                }
                else
                {
                    string sql = "UPDATE SteelGages SET Gage = @Gage, GageUnit = @GageUnit WHERE GageID = @GageID";
                    var rows = await db.ExecuteAsync(sql, model);
                    return rows > 0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving steel gage.");
                throw;
            }
        }

        public async Task<bool> DeleteSteelGageAsync(string gageId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM SteelGages WHERE GageID = @GageId";
                var rows = await db.ExecuteAsync(sql, new { GageId = gageId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting steel gage.");
                throw;
            }
        }
    }
}
