using Microsoft.AspNetCore.Components;
using System;
namespace Impulse.Services
{
    public class ReportNavigationServiceCA : IReportNavigationServiceCA
    {
        private readonly NavigationManager _navigation;
        public ReportNavigationServiceCA(NavigationManager navigation) 
        {
            _navigation = navigation;
        }
        public void PrintVoucher(string strVchrNo)
        {
            string strSelectionFormula = $"{{VLedger.VchrNo}}='{strVchrNo}'";
            string strEncodedFormula = Uri.EscapeDataString(strSelectionFormula);
            _navigation.NavigateTo($"/myreports/Voucher.rpt/{strEncodedFormula}");
        }
    }
}
