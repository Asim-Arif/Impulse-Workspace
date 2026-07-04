using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class CustomerCatalogDataAccess : ICustomerCatalogDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<CustomerCatalogDataAccess> _logger;

        public CustomerCatalogDataAccess(IConfiguration configuration, ILogger<CustomerCatalogDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT CustCode, Name FROM ForeignCustomers WHERE Active = 1 ORDER BY Name";
                var result = await db.QueryAsync<CustomerLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching active foreign customers.");
                return new List<CustomerLookupModel>();
            }
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";
                var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching countries for customer: {CustCode}", custCode);
                return new List<string>();
            }
        }

        public async Task<double> GetMaxDiscountAsync(string custCode, string country)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ISNULL(MaxDiscount, 0) FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
                return await db.QueryFirstOrDefaultAsync<double>(sql, new { CustCode = custCode, Country = country });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching MaxDiscount for customer: {CustCode}, country: {Country}", custCode, country);
                return 0.0;
            }
        }

        public async Task<List<ItemGroupLookupModel>> GetItemGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ID, Description FROM ItemGroups ORDER BY Description";
                var result = await db.QueryAsync<ItemGroupLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item groups.");
                return new List<ItemGroupLookupModel>();
            }
        }

        public async Task<List<CustomerCatalogItemViewModel>> GetCatalogItemsAsync(string custCode, string country, int? groupID)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = @"
                    SELECT 
                        CustCode, Country, ItemID, CompItemID, Description, Unit, 
                        FOB, ExWorks, CnISea, CnIAir, CIFSea, CIFAir, CnFSea, CnFAir, 
                        FOBTop, ExWorksTop, CnIAirTop, CnISeaTop, CIFSeaTop, CIFAirTop, CnFSeaTop, CnFAirTop, 
                        PackingMode, BarcodeNo, SpecialInstructions, StampInstructions, PackingInstructions, ItemColor, BarcodeFile,
                        ILO, OCR, OCR2, FinQuality, Temper_Rate, First_Inspection_Rate, MDMA, SFDA_Listing_No, MD_Group,
                        CompDesc, CatID, CatDesc, Running_Lots_No, Running_Lots_Qty, Forging_Stock, ItemGroup, GroupID, InActive, EntryID,
                        ItemSize, SizeUnit, TipSize
                    FROM VFCustomerCatalog_Ex
                    WHERE CustCode = @CustCode AND Country = @Country AND InActive = 0";
                
                if (groupID.HasValue && groupID.Value != 0)
                {
                    sql += " AND GroupID = @GroupID";
                }
                
                sql += " ORDER BY CompItemID";
                
                var result = await db.QueryAsync<CustomerCatalogItemViewModel>(sql, new { CustCode = custCode, Country = country, GroupID = groupID });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching customer catalog items for customer: {CustCode}, country: {Country}, group: {GroupID}", custCode, country, groupID);
                return new List<CustomerCatalogItemViewModel>();
            }
        }

        public async Task<List<ItemLookupModel>> GetUnassignedItemsAsync(string custCode, string country)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName
                    FROM Items
                    WHERE ItemID NOT IN (
                        SELECT CompItemID FROM FCustomerCatalog WHERE CustCode = @CustCode AND Country = @Country
                    ) AND InActive = 0
                    ORDER BY ItemID";
                var result = await db.QueryAsync<ItemLookupModel>(sql, new { CustCode = custCode, Country = country });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching unassigned items for customer: {CustCode}, country: {Country}", custCode, country);
                return new List<ItemLookupModel>();
            }
        }

        public async Task<int> AssignItemAsync(string custCode, string country, string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                var parameters = new DynamicParameters();
                parameters.Add("@CustCode", custCode);
                parameters.Add("@CustCountry", country);
                parameters.Add("@ItemID", itemId);
                return await db.ExecuteAsync("SP_AssItem", parameters, commandType: CommandType.StoredProcedure);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error assigning item {ItemID} to customer: {CustCode}, country: {Country}", itemId, custCode, country);
                throw;
            }
        }

        public async Task<int> AssignAllItemsAsync(string custCode, string country)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                var parameters = new DynamicParameters();
                parameters.Add("@CustCode", custCode);
                parameters.Add("@CustCountry", country);
                parameters.Add("@ItemID", "%");
                return await db.ExecuteAsync("SP_AssItem", parameters, commandType: CommandType.StoredProcedure);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error assigning all items to customer: {CustCode}, country: {Country}", custCode, country);
                throw;
            }
        }

        public async Task<string?> ItemIsInOrderAsync(string compItemId, string custCode, string country)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT TOP 1 OrderNo FROM VrptOrders WHERE CustCode = @CustCode AND Country = @Country AND CompItemID = @CompItemID";
                return await db.QueryFirstOrDefaultAsync<string>(sql, new { CustCode = custCode, Country = country, CompItemID = compItemId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking order status for item: {CompItemID}, customer: {CustCode}, country: {Country}", compItemId, custCode, country);
                return null;
            }
        }

        public async Task<int> UnassignItemAsync(string custCode, string country, string compItemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "DELETE FROM FCustomerCatalog WHERE CustCode = @CustCode AND Country = @Country AND CompItemID = @CompItemID";
                return await db.ExecuteAsync(sql, new { CustCode = custCode, Country = country, CompItemID = compItemId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting catalog item: {CompItemID}, customer: {CustCode}, country: {Country}", compItemId, custCode, country);
                throw;
            }
        }

        public async Task<int> UnassignAllItemsAsync(string custCode, string country)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    DELETE FROM FCustomerCatalog 
                    WHERE CustCode = @CustCode 
                      AND Country = @Country 
                      AND CompItemID NOT IN (
                          SELECT CompItemCode 
                          FROM FOrderItems 
                          INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo = FOrderItems.OrderNo 
                          WHERE CustCode = @CustCode AND Country = @Country
                      )";
                return await db.ExecuteAsync(sql, new { CustCode = custCode, Country = country });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error unassigning all items for customer: {CustCode}, country: {Country}", custCode, country);
                throw;
            }
        }

        public async Task<int> UpdateRatesAsync(string custCode, double percentage, bool isIncrease)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string op = isIncrease ? "+" : "-";
                string sql = $@"
                    UPDATE FCustomerCatalog 
                    SET FOB = ROUND(FOB {op} ((FOB / 100.0) * @Percentage), 2),
                        ExWorks = ROUND(ExWorks {op} ((ExWorks / 100.0) * @Percentage), 2),
                        CnFAir = ROUND(CnFAir {op} ((CnFAir / 100.0) * @Percentage), 2),
                        CnFSea = ROUND(CnFSea {op} ((CnFSea / 100.0) * @Percentage), 2),
                        CIFAir = ROUND(CIFAir {op} ((CIFAir / 100.0) * @Percentage), 2),
                        CIFSea = ROUND(CIFSea {op} ((CIFSea / 100.0) * @Percentage), 2),
                        CnIAir = ROUND(CnIAir {op} ((CnIAir / 100.0) * @Percentage), 2),
                        CnISea = ROUND(CnISea {op} ((CnISea / 100.0) * @Percentage), 2)
                    WHERE CustCode = @CustCode";
                return await db.ExecuteAsync(sql, new { CustCode = custCode, Percentage = percentage });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error globally updating rates for customer: {CustCode} by {Percentage} percent (Increase: {IsIncrease}).", custCode, percentage, isIncrease);
                throw;
            }
        }

        public async Task<bool> UpdateCatalogItemDetailAsync(CustomerCatalogItemViewModel m)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    UPDATE FCustomerCatalog
                    SET ItemID = @ItemID,
                        Unit = @Unit,
                        Description = @Description,
                        FOB = @FOB,
                        ExWorks = @ExWorks,
                        CnFAir = @CnFAir,
                        CnFSea = @CnFSea,
                        CIFAir = @CIFAir,
                        CIFSea = @CIFSea,
                        CnIAir = @CnIAir,
                        CnISea = @CnISea,
                        PackingMode = @PackingMode,
                        BarcodeNo = @BarcodeNo,
                        SpecialInstructions = @SpecialInstructions,
                        StampInstructions = @StampInstructions,
                        PackingInstructions = @PackingInstructions,
                        ItemColor = @ItemColor,
                        BarcodeFile = @BarcodeFile,
                        FOBTop = @FOBTop,
                        ExWorksTop = @ExWorksTop,
                        CnFAirTop = @CnFAirTop,
                        CnFSeaTop = @CnFSeaTop,
                        CIFAirTop = @CIFAirTop,
                        CIFSeaTop = @CIFSeaTop,
                        CnIAirTop = @CnIAirTop,
                        CnISeaTop = @CnISeaTop,
                        ILO = @ILO,
                        OCR = @OCR,
                        OCR2 = @OCR2,
                        FinQuality = @FinQuality,
                        Temper_Rate = @Temper_Rate,
                        First_Inspection_Rate = @First_Inspection_Rate,
                        MDMA = @MDMA,
                        SFDA_Listing_No = @SFDA_Listing_No,
                        MD_Group = @MD_Group,
                        PageNo = @PageNo
                    WHERE EntryID = @EntryID";
                int result = await db.ExecuteAsync(sql, m);
                return result > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating catalog details for EntryID: {EntryID}", m.EntryID);
                return false;
            }
        }

        public async Task<List<CustomerCatalogRevisionViewModel>> GetRevisionsAsync(int entryId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT EntryID, RefID, DT, UserName FROM FCustCatalRevs WHERE RefID = @RefID ORDER BY DT DESC";
                var result = await db.QueryAsync<CustomerCatalogRevisionViewModel>(sql, new { RefID = entryId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching catalog revisions for RefID: {EntryID}", entryId);
                return new List<CustomerCatalogRevisionViewModel>();
            }
        }

        public async Task<byte[]> GetItemPicAsync(string compItemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ItemPic FROM Items WHERE ItemID = @CompItemID";
                return await db.QueryFirstOrDefaultAsync<byte[]>(sql, new { CompItemID = compItemId }) ?? Array.Empty<byte>();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching ItemPic for CompItemID: {CompItemID}", compItemId);
                return Array.Empty<byte>();
            }
        }

        public async Task<List<string>> GetUnitsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT UnitName FROM Units ORDER BY UnitName";
                var result = await db.QueryAsync<string>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching units.");
                return new List<string>();
            }
        }

        public async Task<List<FinQualityLookupModel>> GetFinishingQualitiesAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ID, Description FROM ItemsFinQuality ORDER BY Description";
                var result = await db.QueryAsync<FinQualityLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching finishing qualities.");
                return new List<FinQualityLookupModel>();
            }
        }

        public async Task<List<string>> GetMDGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT DISTINCT MD_Group FROM FCustomerCatalog WHERE ISNULL(MD_Group,'') <> '' ORDER BY MD_Group";
                var result = await db.QueryAsync<string>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching MD groups.");
                return new List<string>();
            }
        }

        public async Task<bool> UploadArtworkPDFAsync(int entryId, byte[] pdfBytes)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "UPDATE FCustomerCatalog SET ArtworkPDF = @ArtworkPDF WHERE EntryID = @EntryID";
                int rows = await db.ExecuteAsync(sql, new { ArtworkPDF = pdfBytes, EntryID = entryId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading ArtworkPDF for EntryID: {EntryID}", entryId);
                return false;
            }
        }

        public async Task<string> GetPageNoAsync(int entryId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ISNULL(PageNo,'') FROM FCustomerCatalog WHERE EntryID = @EntryID";
                return await db.QueryFirstOrDefaultAsync<string>(sql, new { EntryID = entryId }) ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching PageNo for EntryID: {EntryID}", entryId);
                return string.Empty;
            }
        }
    }
}
