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

        /// <summary>Maps to Items.ItemType (tinyint). Populated from ItemTypes.ID lookup.</summary>
        public byte? ItemType { get; set; }

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
        public decimal FOB { get; set; }
        public decimal? CIFSea { get; set; }
        public decimal? CIFAir { get; set; }
        public decimal? CnFSea { get; set; }
        public decimal? CnFAir { get; set; }
        public decimal? ExWorks { get; set; }
        public double? PlantRate { get; set; }
        public double? SnaffRate { get; set; }
        public double? StampRate { get; set; }
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

        // ── Child collections (Processes & Catalog Refs tabs) ─────────────────
        public List<ItemProcessModel> Processes { get; set; } = new();
        public List<ItemCatalogRefModel> CatalogRefs { get; set; } = new();

        // ── Display-only (populated when loading via VItems) ─────────────────
        public string? CategoryName { get; set; }
        public string? ItemGroup { get; set; }
        public string? MainGroupName { get; set; }
    }
}
