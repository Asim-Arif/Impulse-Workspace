using System.Data;
using System.Data.SqlClient;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class ItemDataAccess : IItemDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<ItemDataAccess> _logger;

        public ItemDataAccess(IConfiguration configuration, ILogger<ItemDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                             ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        // ── Lookups ──────────────────────────────────────────────────────────

        public async Task<List<ItemCategoryModel>> GetCategoriesAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT CatID, Description FROM VItemCatagories ORDER BY Description";
                var result = await db.QueryAsync<ItemCategoryModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item categories.");
                return new List<ItemCategoryModel>();
            }
        }

        public async Task<List<ItemGroupModel>> GetItemGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ID, Description, MainGroup_RefID FROM ItemGroups ORDER BY Description";
                var result = await db.QueryAsync<ItemGroupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item groups.");
                return new List<ItemGroupModel>();
            }
        }

        public async Task<List<ItemMainGroupModel>> GetMainGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT MainGroupID, MainGroupName FROM ItemGroupsMain ORDER BY MainGroupName";
                var result = await db.QueryAsync<ItemMainGroupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching main groups.");
                return new List<ItemMainGroupModel>();
            }
        }

        public async Task<List<UnitModel>> GetUnitsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT UnitID, UnitName FROM Units ORDER BY UnitName";
                var result = await db.QueryAsync<UnitModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching units.");
                return new List<UnitModel>();
            }
        }

        public async Task<List<ItemTypeModel>> GetItemTypesAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT ID, Description FROM ItemTypes ORDER BY Description";
                var result = await db.QueryAsync<ItemTypeModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item types.");
                return new List<ItemTypeModel>();
            }
        }

        // ── Read ─────────────────────────────────────────────────────────────

        public async Task<bool> ItemIdExistsAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT COUNT(1) FROM Items WHERE ItemID = @ItemID";
                int count = await db.ExecuteScalarAsync<int>(sql, new { ItemID = itemId });
                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking item existence for {ItemID}.", itemId);
                return false;
            }
        }

        public async Task<ItemDto?> GetItemByIdAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT
                        ItemID, CatID, GroupID, MainGroupID, ItemType,
                        ItemName, CustomDescription, Description, Additional_Detail, MakerDescription,
                        Unit, Type, ItemSize, SizeUnit, TipSize, ItemUsage,
                        FinQuality, ItemColor, SteelUsed, Gage,
                        Tagging, EAN128, GTINBarcodeNo, UMDNSCode,
                        FDAListingNo, FDAProductCode, EuropeanRegNo, FDA510K,
                        SFDA_Name, SFDA_No, SFDA_Listing,
                        Fig_No, Scale, HRC_From, HRC_To,
                        MinLevel, MaxLevel, ReOrderLevel, ReorderPoint,
                        ItemMaxLotSize, ItemLotSizeBuffer, FixedPackingUnit,
                        UnitWeight, ForgingWeight, FinishedWeight, WasteVisible,
                        ItemCurr, FOB, ExWorks, CnISea, CnIAir, CIFSea, CIFAir, CnFSea, CnFAir,
                        FOBTop, ExWorksTop, CnISeaTop, CnIAirTop, CIFSeaTop, CIFAirTop, CnFSeaTop, CnFAirTop,
                        PlantRate, SnaffRate, StampRate,
                        ReadyFinishPrice, FillingPrice, PriceForCost,
                        PackingInstructions, PolyBag, POInustructions,
                        MasterCartonL, MasterCartonW, MasterCartonH,
                        SmallCartonL, SmallCartonW, SmallCartonH,
                        SmallBoxPcs, MasterCartonSmallBoxes,
                        FinishDescription1, FinishDescription2,
                        MakerDescription1, MakerDescription2,
                        InActive, Sample, AvailableForECommerce,
                        ItemPic, TechnicalDrawing,
                        ItemGroup, Category AS CategoryName, MainGroupName
                    FROM VItems
                    WHERE ItemID = @ItemID";
                var item = await db.QuerySingleOrDefaultAsync<ItemDto>(sql, new { ItemID = itemId });
                if (item != null)
                {
                    var fam = await GetItemProcessFamilyInternalAsync(db, itemId);
                    item.ProcessGroupID = fam.GroupId;
                    item.ProcessFamily = fam.GroupName;

                    try
                    {
                        var rate = await db.ExecuteScalarAsync<decimal?>("SELECT TOP 1 EuroToUS FROM EuroToUS");
                        item.EuroRate = rate ?? 0m;
                    }
                    catch
                    {
                        item.EuroRate = 0m;
                    }
                }
                return item;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item {ItemID}.", itemId);
                return null;
            }
        }

        // ── Child collections ────────────────────────────────────────────────

        public async Task<List<ItemProcessModel>> GetItemProcessesAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT
                        vp.ProcessID, vp.Description, vp.Code, vp.Scanning,
                        vp.ItemSNo AS SNO, vp.Rate, vp.MinLevel, vp.MaxLevel,
                        vp.ReOrderLevel, vp.EntryID
                    FROM VItemProcesses vp
                    WHERE vp.IsExist = @ItemID
                    ORDER BY vp.ItemSNo";
                var result = await db.QueryAsync<ItemProcessModel>(sql, new { ItemID = itemId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching processes for item {ItemID}.", itemId);
                return new List<ItemProcessModel>();
            }
        }

        public async Task<List<ItemCatalogRefModel>> GetItemCatalogRefsAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT EntryID, CatalogName, Reference
                    FROM ItemsCatalogReferences
                    WHERE ItemID = @ItemID
                    ORDER BY EntryID";
                var result = await db.QueryAsync<ItemCatalogRefModel>(sql, new { ItemID = itemId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching catalog refs for item {ItemID}.", itemId);
                return new List<ItemCatalogRefModel>();
            }
        }

        public async Task<List<ItemRMComponentModel>> GetItemRMComponentsAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        v.EntryID,
                        v.ItemID,
                        v.RMID1 AS MaterialCode,
                        v.Qty,
                        v.ProcessID,
                        v.RMName AS MaterialName,
                        v.Unit,
                        v.ProcDesc AS ProcessName,
                        CAST(v.RM AS BIT) AS IsRawMaterial,
                        rm.RMID
                    FROM VItemsRMSFUnioned v
                    LEFT JOIN RM rm ON v.RM = 1 AND rm.RMID1 = v.RMID1
                    WHERE v.ItemID = @ItemID
                    ORDER BY v.EntryID";
                var result = await db.QueryAsync<ItemRMComponentModel>(sql, new { ItemID = itemId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching RM components for item {ItemID}.", itemId);
                return new List<ItemRMComponentModel>();
            }
        }

        public async Task<List<MaterialLookupModel>> GetMaterialLookupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        rm.RMID,
                        rm.RMID1 AS MaterialCode,
                        rm.RMName AS MaterialName,
                        CAST(1 AS BIT) AS IsRawMaterial
                    FROM dbo.RM rm
                    UNION ALL
                    SELECT 
                        NULL AS RMID,
                        items.ItemID AS MaterialCode,
                        items.ItemName AS MaterialName,
                        CAST(0 AS BIT) AS IsRawMaterial
                    FROM Items items
                    WHERE items.ItemType = 1
                    ORDER BY MaterialCode";
                var result = await db.QueryAsync<MaterialLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching material lookups.");
                return new List<MaterialLookupModel>();
            }
        }

        public async Task<List<ItemLookAlikeModel>> GetItemLookAlikesAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        EntryID,
                        ItemID,
                        ItemID_LookAlike,
                        ItemName,
                        InHand
                    FROM VItemLookalikes
                    WHERE ItemID = @ItemID
                    ORDER BY EntryID";
                var result = await db.QueryAsync<ItemLookAlikeModel>(sql, new { ItemID = itemId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching lookalike items for {ItemID}.", itemId);
                return new List<ItemLookAlikeModel>();
            }
        }

        public async Task<List<ItemSearchLookupModel>> GetItemSearchLookupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        ItemID,
                        ISNULL(ItemName, '') AS ItemName,
                        ISNULL(InHand, 0) AS InHand
                    FROM Items
                    ORDER BY ItemID";
                var result = await db.QueryAsync<ItemSearchLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item search lookups.");
                return new List<ItemSearchLookupModel>();
            }
        }

        public async Task<List<ItemSetDetailModel>> GetItemSetDetailsAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        EntryID,
                        ItemID,
                        Set_ItemID,
                        Qty,
                        SetItemName,
                        ItemSize,
                        SizeUnit,
                        ExWorks AS Price
                    FROM VItemsSets
                    WHERE ItemID = @ItemID
                    ORDER BY EntryID";
                var result = await db.QueryAsync<ItemSetDetailModel>(sql, new { ItemID = itemId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching set details for {ItemID}.", itemId);
                return new List<ItemSetDetailModel>();
            }
        }

        public async Task<List<ItemSetLookupModel>> GetItemSetLookupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        ItemID,
                        ISNULL(ItemName, '') AS ItemName,
                        ItemSize,
                        ISNULL(ExWorks, 0) AS ExWorks
                    FROM Items
                    ORDER BY ItemID";
                var result = await db.QueryAsync<ItemSetLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching item set lookups.");
                return new List<ItemSetLookupModel>();
            }
        }

        public async Task<List<ProcessGroupLookupModel>> GetProcessGroupsLookupAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT EntryID, ISNULL(GroupName, '') AS GroupName FROM ProcessGroups ORDER BY GroupName";
                var result = await db.QueryAsync<ProcessGroupLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching process groups lookup.");
                return new List<ProcessGroupLookupModel>();
            }
        }

        public async Task<List<ProcessItemLookupModel>> GetProcessesLookupAsync(int? groupId = null)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                if (groupId.HasValue && groupId.Value > 0)
                {
                    const string sql = @"
                        SELECT P.ProcessID, P.Description, P.Code, ISNULL(PGP.Scanning, 0) AS Scanning, PGP.SeqNo
                        FROM ProcessGroupsProcesses PGP
                        INNER JOIN Processes P ON PGP.Process_RefID = P.ProcessID
                        WHERE PGP.Group_RefID = @GroupID
                        ORDER BY PGP.SeqNo";
                    var result = await db.QueryAsync<ProcessItemLookupModel>(sql, new { GroupID = groupId.Value });
                    return result.ToList();
                }
                else
                {
                    const string sql = @"
                        SELECT ProcessID, Description, Code, CAST(0 AS bit) AS Scanning, SNo AS SeqNo
                        FROM Processes
                        ORDER BY SNo";
                    var result = await db.QueryAsync<ProcessItemLookupModel>(sql);
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching processes lookup for group {GroupID}.", groupId);
                return new List<ProcessItemLookupModel>();
            }
        }

        public async Task<(int? GroupId, string? GroupName)> GetItemProcessFamilyAsync(string itemId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                return await GetItemProcessFamilyInternalAsync(db, itemId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching process family for item {ItemID}.", itemId);
                return (null, null);
            }
        }

        public async Task<List<ForeignCatalogLookupModel>> GetForeignCatalogsLookupAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT CatalID, ISNULL(CatalogName, '') AS CatalogName FROM ForeignCatalog ORDER BY CatalogName";
                var result = await db.QueryAsync<ForeignCatalogLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching foreign catalogs lookup.");
                return new List<ForeignCatalogLookupModel>();
            }
        }

        private static async Task<(int? GroupId, string? GroupName)> GetItemProcessFamilyInternalAsync(IDbConnection db, string itemId)
        {
            const string sql = "SELECT TOP 1 PG_RefID AS GroupId, GroupName FROM VItemProcessGroups WHERE ItemID = @ItemID";
            var row = await db.QueryFirstOrDefaultAsync<dynamic>(sql, new { ItemID = itemId });
            if (row != null)
            {
                return ((int?)row.GroupId, (string?)row.GroupName);
            }
            return (null, null);
        }

        // ── Save (Insert / Update) ───────────────────────────────────────────

        public async Task<bool> SaveItemAsync(ItemDto item, bool isAdd)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using IDbTransaction tx = db.BeginTransaction();
            try
            {
                await SaveBaseItemAsync(db, tx, item, isAdd);
                await SaveItemProcessesAsync(db, tx, item);
                await SaveItemCatalogRefsAsync(db, tx, item);
                await SaveItemRMComponentsAsync(db, tx, item);
                await SaveItemLookAlikesAsync(db, tx, item);
                await SaveItemSetDetailsAsync(db, tx, item);

                tx.Commit();
                return true;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving item {ItemID}.", item.ItemID);
                return false;
            }
        }

        // ── Private helpers ──────────────────────────────────────────────────

        private async Task SaveBaseItemAsync(IDbConnection db, IDbTransaction tx, ItemDto item, bool isAdd)
        {
            if (isAdd)
            {
                const string sql = @"
                    INSERT INTO Items (
                        ItemID, CatID, GroupID, MainGroupID, ItemType,
                        ItemName, CustomDescription, Description, Additional_Detail, MakerDescription,
                        Unit, Type, ItemSize, SizeUnit, TipSize, ItemUsage,
                        FinQuality, ItemColor, SteelUsed, Gage,
                        Tagging, EAN128, GTINBarcodeNo, UMDNSCode,
                        FDAListingNo, FDAProductCode, EuropeanRegNo, FDA510K,
                        SFDA_Name, SFDA_No, SFDA_Listing,
                        Fig_No, Scale, HRC_From, HRC_To,
                        MinLevel, MaxLevel, ReOrderLevel, ReorderPoint,
                        ItemMaxLotSize, ItemLotSizeBuffer, FixedPackingUnit,
                        UnitWeight, ForgingWeight, FinishedWeight, WasteVisible,
                        ItemCurr, FOB, ExWorks, CnISea, CnIAir, CIFSea, CIFAir, CnFSea, CnFAir,
                        FOBTop, ExWorksTop, CnISeaTop, CnIAirTop, CIFSeaTop, CIFAirTop, CnFSeaTop, CnFAirTop,
                        PlantRate, SnaffRate, StampRate,
                        ReadyFinishPrice, FillingPrice, PriceForCost,
                        PackingInstructions, PolyBag, POInustructions,
                        MasterCartonL, MasterCartonW, MasterCartonH,
                        SmallCartonL, SmallCartonW, SmallCartonH,
                        SmallBoxPcs, MasterCartonSmallBoxes,
                        FinishDescription1, FinishDescription2,
                        MakerDescription1, MakerDescription2,
                        InActive, Sample, AvailableForECommerce,
                        ItemPic, TechnicalDrawing,
                        CreateDT
                    ) VALUES (
                        @ItemID, @CatID, @GroupID, @MainGroupID, @ItemType,
                        @ItemName, @CustomDescription, @Description, @Additional_Detail, @MakerDescription,
                        @Unit, @Type, @ItemSize, @SizeUnit, @TipSize, @ItemUsage,
                        @FinQuality, @ItemColor, @SteelUsed, @Gage,
                        @Tagging, @EAN128, @GTINBarcodeNo, @UMDNSCode,
                        @FDAListingNo, @FDAProductCode, @EuropeanRegNo, @FDA510K,
                        @SFDA_Name, @SFDA_No, @SFDA_Listing,
                        @Fig_No, @Scale, @HRC_From, @HRC_To,
                        @MinLevel, @MaxLevel, @ReOrderLevel, @ReorderPoint,
                        @ItemMaxLotSize, @ItemLotSizeBuffer, @FixedPackingUnit,
                        @UnitWeight, @ForgingWeight, @FinishedWeight, @WasteVisible,
                        @ItemCurr, @FOB, @ExWorks, @CnISea, @CnIAir, @CIFSea, @CIFAir, @CnFSea, @CnFAir,
                        @FOBTop, @ExWorksTop, @CnISeaTop, @CnIAirTop, @CIFSeaTop, @CIFAirTop, @CnFSeaTop, @CnFAirTop,
                        @PlantRate, @SnaffRate, @StampRate,
                        @ReadyFinishPrice, @FillingPrice, @PriceForCost,
                        @PackingInstructions, @PolyBag, @POInustructions,
                        @MasterCartonL, @MasterCartonW, @MasterCartonH,
                        @SmallCartonL, @SmallCartonW, @SmallCartonH,
                        @SmallBoxPcs, @MasterCartonSmallBoxes,
                        @FinishDescription1, @FinishDescription2,
                        @MakerDescription1, @MakerDescription2,
                        @InActive, @Sample, @AvailableForECommerce,
                        @ItemPic, @TechnicalDrawing,
                        GETDATE()
                    )";
                await db.ExecuteAsync(sql, item, tx);
            }
            else
            {
                const string sql = @"
                    UPDATE Items SET
                        CatID               = @CatID,
                        GroupID             = @GroupID,
                        MainGroupID         = @MainGroupID,
                        ItemType            = @ItemType,
                        ItemName            = @ItemName,
                        CustomDescription   = @CustomDescription,
                        Description         = @Description,
                        Additional_Detail   = @Additional_Detail,
                        MakerDescription    = @MakerDescription,
                        Unit                = @Unit,
                        Type                = @Type,
                        ItemSize            = @ItemSize,
                        SizeUnit            = @SizeUnit,
                        TipSize             = @TipSize,
                        ItemUsage           = @ItemUsage,
                        FinQuality          = @FinQuality,
                        ItemColor           = @ItemColor,
                        SteelUsed           = @SteelUsed,
                        Gage                = @Gage,
                        Tagging             = @Tagging,
                        EAN128              = @EAN128,
                        GTINBarcodeNo       = @GTINBarcodeNo,
                        UMDNSCode           = @UMDNSCode,
                        FDAListingNo        = @FDAListingNo,
                        FDAProductCode      = @FDAProductCode,
                        EuropeanRegNo       = @EuropeanRegNo,
                        FDA510K             = @FDA510K,
                        SFDA_Name           = @SFDA_Name,
                        SFDA_No             = @SFDA_No,
                        SFDA_Listing        = @SFDA_Listing,
                        Fig_No              = @Fig_No,
                        Scale               = @Scale,
                        HRC_From            = @HRC_From,
                        HRC_To              = @HRC_To,
                        MinLevel            = @MinLevel,
                        MaxLevel            = @MaxLevel,
                        ReOrderLevel        = @ReOrderLevel,
                        ReorderPoint        = @ReorderPoint,
                        ItemMaxLotSize      = @ItemMaxLotSize,
                        ItemLotSizeBuffer   = @ItemLotSizeBuffer,
                        FixedPackingUnit    = @FixedPackingUnit,
                        UnitWeight          = @UnitWeight,
                        ForgingWeight       = @ForgingWeight,
                        FinishedWeight      = @FinishedWeight,
                        WasteVisible        = @WasteVisible,
                        ItemCurr            = @ItemCurr,
                        FOB                 = @FOB,
                        ExWorks             = @ExWorks,
                        CnISea              = @CnISea,
                        CnIAir              = @CnIAir,
                        CIFSea              = @CIFSea,
                        CIFAir              = @CIFAir,
                        CnFSea              = @CnFSea,
                        CnFAir              = @CnFAir,
                        FOBTop              = @FOBTop,
                        ExWorksTop          = @ExWorksTop,
                        CnISeaTop           = @CnISeaTop,
                        CnIAirTop           = @CnIAirTop,
                        CIFSeaTop           = @CIFSeaTop,
                        CIFAirTop           = @CIFAirTop,
                        CnFSeaTop           = @CnFSeaTop,
                        CnFAirTop           = @CnFAirTop,
                        PlantRate           = @PlantRate,
                        SnaffRate           = @SnaffRate,
                        StampRate           = @StampRate,
                        ReadyFinishPrice    = @ReadyFinishPrice,
                        FillingPrice        = @FillingPrice,
                        PriceForCost        = @PriceForCost,
                        PackingInstructions = @PackingInstructions,
                        PolyBag             = @PolyBag,
                        POInustructions     = @POInustructions,
                        MasterCartonL       = @MasterCartonL,
                        MasterCartonW       = @MasterCartonW,
                        MasterCartonH       = @MasterCartonH,
                        SmallCartonL        = @SmallCartonL,
                        SmallCartonW        = @SmallCartonW,
                        SmallCartonH        = @SmallCartonH,
                        SmallBoxPcs         = @SmallBoxPcs,
                        MasterCartonSmallBoxes = @MasterCartonSmallBoxes,
                        FinishDescription1  = @FinishDescription1,
                        FinishDescription2  = @FinishDescription2,
                        MakerDescription1   = @MakerDescription1,
                        MakerDescription2   = @MakerDescription2,
                        InActive            = @InActive,
                        Sample              = @Sample,
                        AvailableForECommerce = @AvailableForECommerce,
                        ItemPic             = @ItemPic,
                        TechnicalDrawing    = @TechnicalDrawing
                    WHERE ItemID = @ItemID";
                await db.ExecuteAsync(sql, item, tx);
            }

            if (item.EuroRate.HasValue)
            {
                try
                {
                    await db.ExecuteAsync("DELETE FROM EuroToUS; INSERT INTO EuroToUS (EuroToUS) VALUES (@EuroRate)", new { EuroRate = item.EuroRate.Value }, tx);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not update EuroToUS rate table.");
                }
            }
        }

        private async Task SaveItemProcessesAsync(IDbConnection db, IDbTransaction tx, ItemDto item)
        {
            // Delete existing rows then re-insert from the in-memory list
            await db.ExecuteAsync(
                "DELETE FROM ItemProcesses WHERE ItemID = @ItemID",
                new { item.ItemID }, tx);

            if (!item.Processes.Any()) return;

            const string sql = @"
                INSERT INTO ItemProcesses (ItemID, ProcessID, Rate, SNO, Scanning, MinLevel, MaxLevel, ReOrderLevel)
                VALUES (@ItemID, @ProcessID, @Rate, @SNO, @Scanning, @MinLevel, @MaxLevel, @ReOrderLevel)";

            foreach (var p in item.Processes)
            {
                await db.ExecuteAsync(sql, new
                {
                    item.ItemID,
                    p.ProcessID,
                    p.Rate,
                    p.SNO,
                    p.Scanning,
                    p.MinLevel,
                    p.MaxLevel,
                    p.ReOrderLevel
                }, tx);
            }

            // Save or delete ItemProcessGroups link
            if (item.BreakProcessGroupLink)
            {
                await db.ExecuteAsync(
                    "DELETE FROM ItemProcessGroups WHERE ItemID = @ItemID",
                    new { item.ItemID }, tx);
            }
            else if (item.ProcessGroupID.HasValue && item.ProcessGroupID.Value > 0)
            {
                await db.ExecuteAsync(
                    "DELETE FROM ItemProcessGroups WHERE ItemID = @ItemID",
                    new { item.ItemID }, tx);

                await db.ExecuteAsync(
                    "INSERT INTO ItemProcessGroups (ItemID, PG_RefID) VALUES (@ItemID, @PG_RefID)",
                    new { item.ItemID, PG_RefID = item.ProcessGroupID.Value }, tx);
            }
        }

        private async Task SaveItemCatalogRefsAsync(IDbConnection db, IDbTransaction tx, ItemDto item)
        {
            await db.ExecuteAsync(
                "DELETE FROM ItemsCatalogReferences WHERE ItemID = @ItemID",
                new { item.ItemID }, tx);

            if (!item.CatalogRefs.Any()) return;

            const string sql = @"
                INSERT INTO ItemsCatalogReferences (ItemID, CatalogName, Reference)
                VALUES (@ItemID, @CatalogName, @Reference)";

            foreach (var r in item.CatalogRefs)
            {
                await db.ExecuteAsync(sql, new
                {
                    item.ItemID,
                    r.CatalogName,
                    r.Reference
                }, tx);
            }
        }

        private async Task SaveItemRMComponentsAsync(IDbConnection db, IDbTransaction tx, ItemDto item)
        {
            await db.ExecuteAsync("DELETE FROM ItemsRMComp WHERE ItemID = @ItemID", new { item.ItemID }, tx);
            await db.ExecuteAsync("DELETE FROM ItemsSFMaterial WHERE ItemID = @ItemID", new { item.ItemID }, tx);

            if (item.RMComponents == null || !item.RMComponents.Any()) return;

            const string sqlRm = @"
                INSERT INTO ItemsRMComp (ItemID, RMID, Qty, ProcessID)
                VALUES (@ItemID, @RMID, @Qty, @ProcessID)";

            const string sqlSf = @"
                INSERT INTO ItemsSFMaterial (ItemID, ItemID_SF, Qty, ProcessID)
                VALUES (@ItemID, @MaterialCode, @Qty, @ProcessID)";

            foreach (var c in item.RMComponents)
            {
                if (c.IsRawMaterial)
                {
                    long rmid = c.RMID ?? 0;
                    if (rmid == 0 && !string.IsNullOrEmpty(c.MaterialCode))
                    {
                        rmid = await db.ExecuteScalarAsync<long>(
                            "SELECT TOP 1 RMID FROM RM WHERE RMID1 = @Code",
                            new { Code = c.MaterialCode }, tx);
                    }
                    if (rmid > 0)
                    {
                        await db.ExecuteAsync(sqlRm, new
                        {
                            item.ItemID,
                            RMID = rmid,
                            c.Qty,
                            c.ProcessID
                        }, tx);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(c.MaterialCode))
                    {
                        await db.ExecuteAsync(sqlSf, new
                        {
                            item.ItemID,
                            c.MaterialCode,
                            c.Qty,
                            c.ProcessID
                        }, tx);
                    }
                }
            }
        }

        private async Task SaveItemLookAlikesAsync(IDbConnection db, IDbTransaction tx, ItemDto item)
        {
            await db.ExecuteAsync("DELETE FROM ItemLookalikes WHERE ItemID = @ItemID", new { item.ItemID }, tx);

            if (item.LookAlikes == null || !item.LookAlikes.Any()) return;

            const string sql = @"
                INSERT INTO ItemLookalikes (ItemID, ItemID_LookAlike)
                VALUES (@ItemID, @ItemID_LookAlike)";

            foreach (var l in item.LookAlikes)
            {
                if (!string.IsNullOrWhiteSpace(l.ItemID_LookAlike))
                {
                    await db.ExecuteAsync(sql, new
                    {
                        item.ItemID,
                        l.ItemID_LookAlike
                    }, tx);
                }
            }
        }

        private async Task SaveItemSetDetailsAsync(IDbConnection db, IDbTransaction tx, ItemDto item)
        {
            await db.ExecuteAsync("DELETE FROM ItemsSets WHERE ItemID = @ItemID", new { item.ItemID }, tx);

            if (item.ItemType != 2 || item.SetDetails == null || !item.SetDetails.Any()) return;

            const string sql = @"
                INSERT INTO ItemsSets (ItemID, Set_ItemID, Qty)
                VALUES (@ItemID, @Set_ItemID, @Qty)";

            foreach (var s in item.SetDetails)
            {
                if (!string.IsNullOrWhiteSpace(s.Set_ItemID))
                {
                    await db.ExecuteAsync(sql, new
                    {
                        item.ItemID,
                        s.Set_ItemID,
                        s.Qty
                    }, tx);
                }
            }
        }
    }
}
