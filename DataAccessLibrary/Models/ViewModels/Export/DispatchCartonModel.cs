namespace DataAccessLibrary.Models.ViewModels.Export;

public class DispatchCartonModel
{
    public long DispatchListEntryID { get; set; }
    public long CartonNo { get; set; }
    public string OrderNo { get; set; } = string.Empty;
    public string ItemCode { get; set; } = string.Empty;
    public string LotNo { get; set; } = string.Empty;
    public long CartonQty { get; set; }
}
