using Impulse.Configurations;
using Microsoft.JSInterop;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Impulse.Services
{
    public class ReportNavigationService : IReportNavigationService
    {
        private readonly HttpClient _httpClient;
        private readonly IJSRuntime _js;
        private readonly AppSettings _appSettings;
        private readonly LoadingService _loader; 
        public ReportNavigationService(HttpClient httpClient, IJSRuntime js,AppSettings appSettings,LoadingService loader) 
        { 
            _httpClient = httpClient;
            _js = js;
            _appSettings = appSettings;
            _loader = loader;
        }
        public async Task PrintReportAsync(ReportRequest reportRequest)
        {
            await GenerateAndOpenReport(reportRequest);
        }
        public async Task PrintReportAsync_Old(string reportName, string selectionFormula)
        {
            await GenerateAndOpenReport_Old(reportName, selectionFormula);
        }
        public async Task PrintVoucher(string strVchrNo) 
        {
            await GenerateAndOpenReport_Old("PV.rpt", $"{{VLedger.VchrNo}}='{strVchrNo}'");
        }
        public async Task PrintVoucherOld(string strVchrNo) 
        {
            var request = new
            {
                ReportName = "PV.rpt",
                SelectionFormula = "{VLedger.VchrNo}='" + strVchrNo + "'"
            };

            var json = JsonSerializer.Serialize(request);
            var content=new StringContent(json,Encoding.UTF8, "application/json");

            //Call the Report API
            var apiUrl = _appSettings.BaseAddressForReports + "/api/report";
            //var response = await _httpClient.PostAsync("api/report", content);
            var response = await _httpClient.PostAsync(apiUrl, content);

            if (!response.IsSuccessStatusCode) 
            {
                var errorMsg = await response.Content.ReadAsStringAsync();
                Console.WriteLine("API Error: " + errorMsg);
                return;
            }

            //response.EnsureSuccessStatusCode();


            var pdfBytes= await response.Content.ReadAsByteArrayAsync();

            var base64=Convert.ToBase64String(pdfBytes);

            var jsCode = @$"
                var byteCharacters = atob('{base64}');
                var byteNumbers = new Array(byteCharacters.length);
                for (var i = 0; i < byteCharacters.length; i++) {{
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }}
                var byteArray = new Uint8Array(byteNumbers);
                var blob = new Blob([byteArray], {{ type: 'application/pdf' }});
                var url = URL.createObjectURL(blob);
                window.open(url, '_blank');";


            /*var jsCode = $"var w = window.open('about:blank');" +
                         $"w.document.write('<iframe width=\"100%\" height=\"100%\" src=\"data:application/pdf;base64,{base64}\"></iframe>');";
            */
            await _js.InvokeVoidAsync("eval", jsCode);
        }
        private async Task GenerateAndOpenReport(ReportRequest request)
        {
            _loader.Show($"Generating {request.ReportName}...");
            //var request = new { ReportName = reportName, SelectionFormula = selectionFormula };
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");


            try
            {
                var apiUrl = $"{_appSettings.BaseAddressForReports}/api/report";
                var response = await _httpClient.PostAsync(apiUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorMsg = await response.Content.ReadAsStringAsync();
                    // You could inject a Toast service here to show the error to the user
                    Console.WriteLine($"API Error for {request.ReportName}: {errorMsg}");
                    return;
                }
                
                var bytes = await response.Content.ReadAsByteArrayAsync();
                var contentType = response.Content.Headers.ContentType?.MediaType;
                if (contentType == "application/pdf")
                {
                    await OpenPdfInNewTab(bytes);
                }
                else 
                {
                    await DownloadFile(
                        bytes,
                        contentType,
                        "Report.xlsx");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical Error printing {request.ReportName}: {ex.Message}");
            }
            finally 
            {
                _loader.Hide();
            }
        }
        private async Task GenerateAndOpenReport_Old(string reportName, string selectionFormula)
        {
            var request = new { ReportName = reportName, SelectionFormula = selectionFormula };
            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                var apiUrl = $"{_appSettings.BaseAddressForReports}/api/report";
                var response = await _httpClient.PostAsync(apiUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorMsg = await response.Content.ReadAsStringAsync();
                    // You could inject a Toast service here to show the error to the user
                    Console.WriteLine($"API Error for {reportName}: {errorMsg}");
                    return;
                }

                var pdfBytes = await response.Content.ReadAsByteArrayAsync();
                await OpenPdfInNewTab(pdfBytes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical Error printing {reportName}: {ex.Message}");
            }
        }
        private async Task OpenPdfInNewTab(byte[] pdfBytes)
        {
            var base64 = Convert.ToBase64String(pdfBytes);
            await _js.InvokeVoidAsync("openPdfFromBase64", base64);
            // Optimized JS: Uses a Helper function instead of raw 'eval' if possible, 
            // but keeping your logic flow:
            //var jsCode = @$"
            //    (function() {{
            //        var byteCharacters = atob('{base64}');
            //        var byteNumbers = new Array(byteCharacters.length);
            //        for (var i = 0; i < byteCharacters.length; i++) {{
            //            byteNumbers[i] = byteCharacters.charCodeAt(i);
            //        }}
            //        var byteArray = new Uint8Array(byteNumbers);
            //        var blob = new Blob([byteArray], {{ type: 'application/pdf' }});
            //        var url = URL.createObjectURL(blob);
            //        window.open(url, '_blank');
            //    }})();";

            //await _js.InvokeVoidAsync("eval", jsCode);
        }

        private async Task DownloadFile(byte[] bytes, string contentType, string fileName)
        {
            // This calls the JavaScript function we added to _Host.cshtml / pdfHelper.js
            await _js.InvokeVoidAsync("downloadFileFromBytes", bytes, contentType, fileName);
        }

    }
    public class ReportRequest
    {
        public string ReportName { get; set; }
        public string SelectionFormula { get; set; }
        public string Format { get; set; } = "pdf";
        // Parameters
        public Dictionary<string, object> Parameters { get; set; }

        // Formula values
        public Dictionary<string, object> FormulaValues { get; set; }
    }

}
