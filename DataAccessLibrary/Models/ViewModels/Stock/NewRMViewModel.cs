using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class NewRMViewModel
    {
        public int RMID { get; set; }
        
        [Required(ErrorMessage = "Category (Group) is required.")]
        public int GroupID { get; set; }
        
        [Required(ErrorMessage = "Device Code is required.")]
        public string RMID1 { get; set; } = string.Empty;
        
        // Global / General Info
        public string RMName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string UrduName { get; set; } = string.Empty;
        public string RM_Maker_Description { get; set; } = string.Empty;
        public string RMUsage { get; set; } = string.Empty;
        
        // Units & Sizes
        public string Unit { get; set; } = string.Empty;
        public float RMSize { get; set; }
        public string SizeUnit { get; set; } = string.Empty;
        public string Weight_Length { get; set; } = string.Empty;
        
        // Rates & Stock Levels
        public double Rate { get; set; }
        public float MakerRate { get; set; }
        public float OpeningStock { get; set; }
        public float MinLevel { get; set; }
        public float MaxLevel { get; set; }
        public int ReorderPoint { get; set; }
        
        // Checkboxes
        public bool InActive { get; set; }
        public bool Returnable { get; set; }
        public bool RequiresMaleGrinding { get; set; }
        
        // Account Info
        public string SaleAccNo { get; set; } = string.Empty;
        public string SRTAccNo { get; set; } = string.Empty;
        public string AssetAccNo { get; set; } = string.Empty;
        
        // Picture & Technical Drawing
        public byte[]? Pic { get; set; }
        public byte[]? TechnicalDrawing { get; set; }
        
        // Used In
        public List<ItemRMCompViewModel> UsedInItems { get; set; } = new();
    }
}
