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
                        FOB, CIFSea, CIFAir, CnFSea, CnFAir, ExWorks,
                        PlantRate, SnaffRate, StampRate,
                        ReadyFinishPrice, FillingPrice, PriceForCost,
                        PackingInstructions, PolyBag, POInustructions,
                        MasterCartonL, MasterCartonW, MasterCartonH,
                        SmallCartonL, SmallCartonW, SmallCartonH,
                        SmallBoxPcs, MasterCartonSmallBoxes,
                        FinishDescription1, FinishDescription2,
                        MakerDescription1, MakerDescription2,
                        InActive, Sample, AvailableForECommerce,
                        ItemGroup, Category AS CategoryName, MainGroupName
                    FROM VItems
                    WHERE ItemID = @ItemID";
                return await db.QuerySingleOrDefaultAsync<ItemDto>(sql, new { ItemID = itemId });
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
                        FOB, CIFSea, CIFAir, CnFSea, CnFAir, ExWorks,
                        PlantRate, SnaffRate, StampRate,
                        ReadyFinishPrice, FillingPrice, PriceForCost,
                        PackingInstructions, PolyBag, POInustructions,
                        MasterCartonL, MasterCartonW, MasterCartonH,
                        SmallCartonL, SmallCartonW, SmallCartonH,
                        SmallBoxPcs, MasterCartonSmallBoxes,
                        FinishDescription1, FinishDescription2,
                        MakerDescription1, MakerDescription2,
                        InActive, Sample, AvailableForECommerce,
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
                        @FOB, @CIFSea, @CIFAir, @CnFSea, @CnFAir, @ExWorks,
                        @PlantRate, @SnaffRate, @StampRate,
                        @ReadyFinishPrice, @FillingPrice, @PriceForCost,
                        @PackingInstructions, @PolyBag, @POInustructions,
                        @MasterCartonL, @MasterCartonW, @MasterCartonH,
                        @SmallCartonL, @SmallCartonW, @SmallCartonH,
                        @SmallBoxPcs, @MasterCartonSmallBoxes,
                        @FinishDescription1, @FinishDescription2,
                        @MakerDescription1, @MakerDescription2,
                        @InActive, @Sample, @AvailableForECommerce,
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
                        FOB                 = @FOB,
                        CIFSea              = @CIFSea,
                        CIFAir              = @CIFAir,
                        CnFSea              = @CnFSea,
                        CnFAir              = @CnFAir,
                        ExWorks             = @ExWorks,
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
                        AvailableForECommerce = @AvailableForECommerce
                    WHERE ItemID = @ItemID";
                await db.ExecuteAsync(sql, item, tx);
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
    }
}
