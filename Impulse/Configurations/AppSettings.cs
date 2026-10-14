namespace Impulse.Configurations
{
    public class AppSettings
    {
        public string BaseAddress { get; set; }
        public string BaseAddressForReports { get; set; }
        public string UploadPath { get; set; }
        public bool EnableCustomerWiseMakerRates { get; set; } = false;
    }
}
