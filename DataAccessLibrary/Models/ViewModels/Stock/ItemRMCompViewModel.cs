using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class ItemRMCompViewModel
    {
        public long EntryID { get; set; }
        public int RMID { get; set; }
        
        [Required(ErrorMessage = "Item is required.")]
        public string ItemID { get; set; } = string.Empty;
        
        // Joined property for display
        public string ItemName { get; set; } = string.Empty;
        
        // For UI tracking
        public bool IsNew { get; set; } = false;
        public bool IsDeleted { get; set; } = false;
    }
}
