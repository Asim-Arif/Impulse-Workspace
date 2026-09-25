namespace Impulse.Models.Setup
{
    public class FavouriteLinkDto
    {
        public string OptionId { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Href { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
        public string IconClass { get; set; } = "fas fa-link";
        public bool IsAuthorized { get; set; } = true;
    }
}
