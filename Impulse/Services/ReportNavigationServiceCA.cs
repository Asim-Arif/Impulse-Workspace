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
        public void PrintVoucher(string voucherNo) 
        {
            string strSelectionFormula= "{VLedger.VchrNo}='"+voucherNo+"'";
            string strEncodedFormula = Uri.EscapeDataString(strSelectionFormula);
            _navigation.NavigateTo($"/myreports/PV.rpt/{strEncodedFormula}");
        }
    }
}
