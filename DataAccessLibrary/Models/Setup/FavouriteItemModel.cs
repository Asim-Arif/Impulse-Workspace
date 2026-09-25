namespace DataAccessLibrary.Models.Setup
{
    public class FavouriteItemModel
    {
        public int UserID { get; set; }
        public string OptionID { get; set; } = string.Empty;
        public string OptionName { get; set; } = string.Empty;
        public string ModuleName { get; set; } = string.Empty;
    }
}
