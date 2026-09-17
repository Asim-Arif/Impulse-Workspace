using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    // ── Lookup models ────────────────────────────────────────────────────────
    public class ItemCategoryModel
    {
        public int CatID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class ItemGroupModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
        public int? MainGroup_RefID { get; set; }
    }

    public class ItemMainGroupModel
    {
        public int MainGroupID { get; set; }
        public string MainGroupName { get; set; } = string.Empty;
    }

    public class UnitModel
    {
        public int UnitID { get; set; }
        public string UnitName { get; set; } = string.Empty;
    }

    public class ItemTypeModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    // ── Process row (maps to ItemProcesses table + VItemProcesses view) ──────
    public class ItemProcessModel
    {
        public int? EntryID { get; set; }           // PK in ItemProcesses
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public string? Code { get; set; }
        public bool Scanning { get; set; }
        public decimal SNO { get; set; }
        public double? Rate { get; set; }
        public int? MinLevel { get; set; }
        public int? MaxLevel { get; set; }
        public int? ReOrderLevel { get; set; }
    }

    // ── Catalog-reference row (maps to ItemsCatalogReferences table) ─────────
    public class ItemCatalogRefModel
    {
        public int? EntryID { get; set; }           // PK in ItemsCatalogReferences
        public string CatalogName { get; set; } = string.Empty;
        public string Reference { get; set; } = string.Empty;
    }

    // ── Main DTO — properties match Items table columns exactly ─────────────
    public class ItemDto
    {
        // ── Header / identity ────────────────────────────────────────────────
        [Required(ErrorMessage = "Item Code is required.")]
        [StringLength(50)]
        public string ItemID { get; set; } = string.Empty;

        public int? CatID { get; set; }
        public int? GroupID { get; set; }
        public int? MainGroupID { get; set; }

        /// <summary>Maps to Items.ItemType (0 = Finished, 1 = Semi Finished, 2 = Set).</summary>
        public int? ItemType { get; set; } = 0;

        public bool InActive { get; set; }
        public bool Sample { get; set; }
        public bool AvailableForECommerce { get; set; }

        // ── General Info tab ─────────────────────────────────────────────────
        [Required(ErrorMessage = "Item Name is required.")]
        [StringLength(1000)]
        public string? ItemName { get; set; }

        [StringLength(1000)]
        public string? CustomDescription { get; set; }

        [StringLength(1000)]
        public string? Description { get; set; }

        [StringLength(4000)]
        public string? Additional_Detail { get; set; }

        [StringLength(4000)]
        public string? MakerDescription { get; set; }

        /// <summary>Maps to Items.Unit — free-text unit code (e.g. "PCS").</summary>
        public string? Unit { get; set; }

        /// <summary>Maps to Items.Type — item type string (legacy free-text).</summary>
        [StringLength(50)]
        public string? Type { get; set; }

        [StringLength(50)]
        public string? ItemSize { get; set; }

        [StringLength(50)]
        public string? SizeUnit { get; set; }

        [StringLength(50)]
        public string? TipSize { get; set; }

        [StringLength(500)]
        public string? ItemUsage { get; set; }

        [StringLength(50)]
        public string? FinQuality { get; set; }

        [StringLength(50)]
        public string? ItemColor { get; set; }

        /// <summary>Maps to Items.SteelUsed (int FK to a steel-type table).</summary>
        public int? SteelUsed { get; set; }

        [StringLength(50)]
        public string? Gage { get; set; }

        [StringLength(50)]
        public string? Tagging { get; set; }

        [StringLength(50)]
        public string? EAN128 { get; set; }

        [StringLength(50)]
        public string? GTINBarcodeNo { get; set; }

        [StringLength(50)]
        public string? UMDNSCode { get; set; }

        // ── SFDA / Regulatory ────────────────────────────────────────────────
        [StringLength(50)]
        public string? FDAListingNo { get; set; }

        [StringLength(50)]
        public string? FDAProductCode { get; set; }

        [StringLength(50)]
        public string? EuropeanRegNo { get; set; }

        [StringLength(50)]
        public string? FDA510K { get; set; }

        [StringLength(50)]
        public string? SFDA_Name { get; set; }

        [StringLength(50)]
        public string? SFDA_No { get; set; }

        [StringLength(50)]
        public string? SFDA_Listing { get; set; }

        [StringLength(255)]
        public string? Fig_No { get; set; }

        [StringLength(255)]
        public string? Scale { get; set; }

        public int? HRC_From { get; set; }
        public int? HRC_To { get; set; }

        // ── Stock levels ─────────────────────────────────────────────────────
        public int? MinLevel { get; set; }
        public int? MaxLevel { get; set; }
        public int? ReOrderLevel { get; set; }
        public int? ReorderPoint { get; set; }
        public int? ItemMaxLotSize { get; set; }
        public int? ItemLotSizeBuffer { get; set; }
        public int? FixedPackingUnit { get; set; }

        // ── Weight tab ───────────────────────────────────────────────────────
        /// <summary>Maps to Items.UnitWeight (varchar 50) — stored as text in DB.</summary>
        public string? UnitWeight { get; set; }

        public double? ForgingWeight { get; set; }
        public double? FinishedWeight { get; set; }
        public double? WasteVisible { get; set; }

        // ── Pricing tab ──────────────────────────────────────────────────────
        [StringLength(50)]
        public string? ItemCurr { get; set; } = "US$";
        public decimal? EuroRate { get; set; }

        // Regular Pricing
        public decimal FOB { get; set; }
        public decimal? ExWorks { get; set; }
        public decimal? CnISea { get; set; }
        public decimal? CnIAir { get; set; }
        public decimal? CIFSea { get; set; }
        public decimal? CIFAir { get; set; }
        public decimal? CnFSea { get; set; }
        public decimal? CnFAir { get; set; }

        // Top Pricing
        public decimal? FOBTop { get; set; }
        public decimal? ExWorksTop { get; set; }
        public decimal? CnISeaTop { get; set; }
        public decimal? CnIAirTop { get; set; }
        public decimal? CIFSeaTop { get; set; }
        public decimal? CIFAirTop { get; set; }
        public decimal? CnFSeaTop { get; set; }
        public decimal? CnFAirTop { get; set; }

        // Process Rates
        public double? PlantRate { get; set; }
        public double? SnaffRate { get; set; }
        public double? StampRate { get; set; }

        // Cost Prices
        public float? ReadyFinishPrice { get; set; }
        public float? FillingPrice { get; set; }
        public double? PriceForCost { get; set; }

        // ── Shipping tab ─────────────────────────────────────────────────────
        [StringLength(1000)]
        public string? PackingInstructions { get; set; }

        [StringLength(255)]
        public string? PolyBag { get; set; }

        [StringLength(1000)]
        public string? POInustructions { get; set; }

        public float? MasterCartonL { get; set; }
        public float? MasterCartonW { get; set; }
        public float? MasterCartonH { get; set; }
        public float? SmallCartonL { get; set; }
        public float? SmallCartonW { get; set; }
        public float? SmallCartonH { get; set; }
        public int? SmallBoxPcs { get; set; }
        public int? MasterCartonSmallBoxes { get; set; }

        // ── Descriptions (finish / maker) ────────────────────────────────────
        [StringLength(1000)]
        public string? FinishDescription1 { get; set; }

        [StringLength(1000)]
        public string? FinishDescription2 { get; set; }

        [StringLength(1000)]
        public string? MakerDescription1 { get; set; }

        [StringLength(1000)]
        public string? MakerDescription2 { get; set; }

        // ── Picture & Tech Drawing ──────────────────────────────────────────
        public byte[]? ItemPic { get; set; }
        public byte[]? TechnicalDrawing { get; set; }
        public string? TechnicalDrawingFileName { get; set; }

        public string? ItemPicBase64 => ItemPic != null && ItemPic.Length > 0
            ? $"data:image/jpeg;base64,{Convert.ToBase64String(ItemPic)}"
            : null;

        public string? TechnicalDrawingBase64 => TechnicalDrawing != null && TechnicalDrawing.Length > 0
            ? $"data:application/pdf;base64,{Convert.ToBase64String(TechnicalDrawing)}"
            : null;

        // ── Child collections (Processes, Catalog Refs, RM Components, Look Alikes, Set Details) ──
        public List<ItemProcessModel> Processes { get; set; } = new();
        public List<ItemCatalogRefModel> CatalogRefs { get; set; } = new();
        public List<ItemRMComponentModel> RMComponents { get; set; } = new();
        public List<ItemLookAlikeModel> LookAlikes { get; set; } = new();
        public List<ItemSetDetailModel> SetDetails { get; set; } = new();

        // ── Process Family Link (ItemProcessGroups) ───────────────────────────
        public int? ProcessGroupID { get; set; }
        public string? ProcessFamily { get; set; }
        public bool BreakProcessGroupLink { get; set; }

        // ── Display-only (populated when loading via VItems) ─────────────────
        public string? CategoryName { get; set; }
        public string? ItemGroup { get; set; }
        public string? MainGroupName { get; set; }
    }

    public class ItemSetDetailModel
    {
        public int? EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string Set_ItemID { get; set; } = string.Empty;
        public string? SetItemName { get; set; }
        public string? ItemSize { get; set; }
        public string? SizeUnit { get; set; }
        public int Qty { get; set; } = 1;
        public decimal? Price { get; set; }
        public decimal Amount => (decimal)Qty * (Price ?? 0m);
    }

    public class ItemSetLookupModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string? ItemSize { get; set; }
        public decimal? ExWorks { get; set; }
        public string DisplayText => $"[{ItemID}] {ItemName}";
    }

    public class ItemLookAlikeModel
    {
        public int? EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemID_LookAlike { get; set; } = string.Empty;
        public string? ItemName { get; set; }
        public decimal? InHand { get; set; }
    }

    public class ItemSearchLookupModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal? InHand { get; set; }
        public string DisplayText => $"[{ItemID}] {ItemName}";
    }

    public class ItemRMComponentModel
    {
        public int? EntryID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public long? RMID { get; set; }
        public string MaterialCode { get; set; } = string.Empty;
        public string MaterialName { get; set; } = string.Empty;
        public int? ProcessID { get; set; }
        public string? ProcessName { get; set; }
        public decimal? Qty { get; set; }
        public string? Unit { get; set; }
        public bool IsRawMaterial { get; set; } = true;
    }

    public class MaterialLookupModel
    {
        public long? RMID { get; set; }
        public string MaterialCode { get; set; } = string.Empty;
        public string MaterialName { get; set; } = string.Empty;
        public bool IsRawMaterial { get; set; }
        public string DisplayText => $"[{MaterialCode}] {MaterialName} ({(IsRawMaterial ? "RM" : "SF")})";
    }

    public class ProcessGroupLookupModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; } = string.Empty;
    }

    public class ProcessItemLookupModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public string? Code { get; set; }
        public bool Scanning { get; set; }
        public int SeqNo { get; set; }
    }

    public class ForeignCatalogLookupModel
    {
        public int CatalID { get; set; }
        public string CatalogName { get; set; } = string.Empty;
    }
}
