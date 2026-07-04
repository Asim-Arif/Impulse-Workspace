using System;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class CustomerCatalogItemViewModel
    {
        public int EntryID { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty; // Customer Item Code
        public string Unit { get; set; } = string.Empty;
        public string CompItemID { get; set; } = string.Empty; // Company Internal Item ID
        public string Description { get; set; } = string.Empty; // Customer specific description
        
        // Prices
        public double FOB { get; set; }
        public double ExWorks { get; set; }
        public double CnFAir { get; set; }
        public double CnFSea { get; set; }
        public double CIFAir { get; set; }
        public double CIFSea { get; set; }
        public double CnIAir { get; set; }
        public double CnISea { get; set; }

        public int? PackingMode { get; set; }
        public string BarcodeNo { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public string StampInstructions { get; set; } = string.Empty;
        public string PackingInstructions { get; set; } = string.Empty;
        public string ItemColor { get; set; } = string.Empty;
        public string BarcodeFile { get; set; } = string.Empty;

        // Top Quality prices
        public double FOBTop { get; set; }
        public double ExWorksTop { get; set; }
        public double CnFAirTop { get; set; }
        public double CnFSeaTop { get; set; }
        public double CIFAirTop { get; set; }
        public double CIFSeaTop { get; set; }
        public double CnIAirTop { get; set; }
        public double CnISeaTop { get; set; }

        public string ILO { get; set; } = string.Empty;
        public string OCR { get; set; } = string.Empty;
        public string OCR2 { get; set; } = string.Empty;
        public string FinQuality { get; set; } = string.Empty;
        public float? Temper_Rate { get; set; }
        public float? First_Inspection_Rate { get; set; }
        public string MDMA { get; set; } = string.Empty;
        public string SFDA_Listing_No { get; set; } = string.Empty;
        public string MD_Group { get; set; } = string.Empty;
        public string PageNo { get; set; } = string.Empty;

        // Joined view fields
        public string CompDesc { get; set; } = string.Empty;
        public int? CatID { get; set; }
        public string CatDesc { get; set; } = string.Empty;
        public int Running_Lots_No { get; set; }
        public double Running_Lots_Qty { get; set; }
        public double Forging_Stock { get; set; }
        public string ItemGroup { get; set; } = string.Empty;
        public int? GroupID { get; set; }
        public bool InActive { get; set; }
        
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string TipSize { get; set; } = string.Empty;
    }

    public class CustomerCatalogRevisionViewModel
    {
        public int EntryID { get; set; }
        public int RefID { get; set; }
        public DateTime DT { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;

        // Historical prices stored when a change is made
        public double FOB { get; set; }
        public double FOBSialkot { get; set; }   // = ExWorks
        public double CnISea { get; set; }
        public double CnIAir { get; set; }
        public double CIFSea { get; set; }
        public double CIFAir { get; set; }
        public double CnFSea { get; set; }
        public double CnFAir { get; set; }

        // Top quality historical prices
        public double FOBTop { get; set; }
        public double ExWorksTop { get; set; }
        public double CnISeaTop { get; set; }
        public double CnIAirTop { get; set; }
        public double CIFSeaTop { get; set; }
        public double CIFAirTop { get; set; }
        public double CnFSeaTop { get; set; }
        public double CnFAirTop { get; set; }
    }

    public class ItemGroupLookupModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class ItemLookupModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
    }

    public class FinQualityLookupModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
    }
}
