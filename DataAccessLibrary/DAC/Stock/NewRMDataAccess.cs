using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class NewRMDataAccess : INewRMDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<NewRMDataAccess> _logger;

        public NewRMDataAccess(IConfiguration configuration, ILogger<NewRMDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                             ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<NewRMViewModel> GetRMAsync(int rmId)
        {
            var model = new NewRMViewModel();

            try
            {
                string sqlRM = @"
                    SELECT 
                        RMID, GroupID, RMID1, RMName, Unit, RMSize, SizeUnit, Rate, RMUsage, 
                        AssetAccNo, SaleAccNo, SRTAccNo, InActive, ReorderPoint, 
                        MinLevel, MaxLevel, OpeningStock, Description, TechnicalDrawing, Pic, 
                        MakerRate, UrduName, RequiresMaleGrinding, 
                        Returnable, Weight_Length, RM_Maker_Description
                    FROM RM
                    WHERE RMID = @rmId;
                ";

                var rmParams = new DynamicParameters();
                rmParams.Add("@rmId", rmId);

                using var connection = new SqlConnection(_connectionString);
                var rmResult = await connection.QueryAsync<NewRMViewModel>(sqlRM, rmParams);
                var rmRecord = rmResult.FirstOrDefault();

                if (rmRecord != null)
                {
                    model = rmRecord;

                    // Fetch Used In Items (ItemsRMComp)
                    string sqlUsedIn = @"
                        SELECT 
                            c.EntryID, c.RMID, c.ItemID, i.ItemName 
                        FROM ItemsRMComp c
                        LEFT JOIN Items i ON c.ItemID = i.ItemID
                        WHERE c.RMID = @rmId;
                    ";

                    var usedInResult = await connection.QueryAsync<ItemRMCompViewModel>(sqlUsedIn, rmParams);
                    model.UsedInItems = usedInResult.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting RM data for RMID {RMID}", rmId);
                throw;
            }

            return model;
        }

        public async Task<int> InsertRMAsync(NewRMViewModel model)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using IDbTransaction transaction = connection.BeginTransaction();

            try
            {
                // RMID is an IDENTITY column — let SQL Server generate it, retrieve via SCOPE_IDENTITY()
                string sqlRM = @"
                    INSERT INTO RM (
                        GroupID, RMID1, RMName, Unit, RMSize, SizeUnit, Rate, RMUsage, 
                        AssetAccNo, SaleAccNo, SRTAccNo, InActive, ReorderPoint, 
                        MinLevel, MaxLevel, OpeningStock, Description, TechnicalDrawing, Pic, 
                        MakerRate, UrduName, RequiresMaleGrinding, 
                        Returnable, Weight_Length, RM_Maker_Description
                    ) VALUES (
                        @GroupID, @RMID1, @RMName, @Unit, @RMSize, @SizeUnit, @Rate, @RMUsage, 
                        @AssetAccNo, @SaleAccNo, @SRTAccNo, @InActive, @ReorderPoint, 
                        @MinLevel, @MaxLevel, @OpeningStock, @Description, @TechnicalDrawing, @Pic, 
                        @MakerRate, @UrduName, @RequiresMaleGrinding, 
                        @Returnable, @Weight_Length, @RM_Maker_Description
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);
                ";

                int newId = await connection.ExecuteScalarAsync<int>(sqlRM, model, transaction);
                model.RMID = newId;

                // Insert UsedIn items
                if (model.UsedInItems != null && model.UsedInItems.Any())
                {
                    string sqlItems = @"
                        INSERT INTO ItemsRMComp (RMID, ItemID) 
                        VALUES (@RMID, @ItemID);
                    ";

                    foreach (var item in model.UsedInItems.Where(x => !x.IsDeleted))
                    {
                        item.RMID = newId;
                        await connection.ExecuteAsync(sqlItems, item, transaction);
                    }
                }

                transaction.Commit();
                return newId;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error inserting new RM");
                throw;
            }
        }

        public async Task<bool> UpdateRMAsync(NewRMViewModel model)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using IDbTransaction transaction = connection.BeginTransaction();

            try
            {
                string sqlRM = @"
                    UPDATE RM SET
                        GroupID = @GroupID, 
                        RMID1 = @RMID1, 
                        RMName = @RMName, 
                        Unit = @Unit, 
                        RMSize = @RMSize, 
                        SizeUnit = @SizeUnit, 
                        Rate = @Rate, 
                        RMUsage = @RMUsage, 
                        AssetAccNo = @AssetAccNo, 
                        SaleAccNo = @SaleAccNo, 
                        SRTAccNo = @SRTAccNo, 
                        InActive = @InActive, 
                        ReorderPoint = @ReorderPoint, 
                        MinLevel = @MinLevel, 
                        MaxLevel = @MaxLevel, 
                        OpeningStock = @OpeningStock, 
                        Description = @Description, 
                        TechnicalDrawing = @TechnicalDrawing, 
                        Pic = @Pic, 
                        MakerRate = @MakerRate, 
                        UrduName = @UrduName, 
                        RequiresMaleGrinding = @RequiresMaleGrinding, 
                        Returnable = @Returnable, 
                        Weight_Length = @Weight_Length, 
                        RM_Maker_Description = @RM_Maker_Description
                    WHERE RMID = @RMID;
                ";

                await connection.ExecuteAsync(sqlRM, model, transaction);

                // Handle UsedIn Items (Delete/Insert)
                if (model.UsedInItems != null)
                {
                    string sqlDelItems = "DELETE FROM ItemsRMComp WHERE EntryID = @EntryID;";
                    string sqlInsItems = "INSERT INTO ItemsRMComp (RMID, ItemID) VALUES (@RMID, @ItemID);";

                    foreach (var item in model.UsedInItems)
                    {
                        if (item.IsDeleted && item.EntryID > 0)
                        {
                            await connection.ExecuteAsync(sqlDelItems, new { EntryID = item.EntryID }, transaction);
                        }
                        else if (item.IsNew && !item.IsDeleted)
                        {
                            item.RMID = model.RMID;
                            await connection.ExecuteAsync(sqlInsItems, item, transaction);
                        }
                    }
                }

                transaction.Commit();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error updating RM {RMID}", model.RMID);
                throw;
            }
        }

        public async Task<bool> DeleteRMAsync(int id)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = "DELETE FROM RM WHERE RMID = @RMID";
            int rows = await connection.ExecuteAsync(sql, new { RMID = id });
            return rows > 0;
        }

        public async Task<List<ItemLookupModel>> GetAvailableItemsAsync()
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = "SELECT ItemID, ItemName FROM Items ORDER BY ItemName";
            var result = await connection.QueryAsync<ItemLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<RawMaterialLookupModel>> GetAvailableRawMaterialsAsync()
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = "SELECT RMID1, RMName, Unit FROM RM ORDER BY RMName";
            var result = await connection.QueryAsync<RawMaterialLookupModel>(sql);
            return result.ToList();
        }
    }
}
