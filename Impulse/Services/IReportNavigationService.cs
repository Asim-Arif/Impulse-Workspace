namespace Impulse.Services
{
    public interface IReportNavigationService
    {
        Task PrintVoucher(string voucherNo);
        Task PrintReportAsync_Old(string reportName, string selectionFormula);
        Task PrintReportAsync(ReportRequest reportRequest);
       
    }
}
