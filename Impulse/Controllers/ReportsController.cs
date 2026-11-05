using Impulse.Configurations;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Impulse.Controllers
{
    [Route("api/reports")]
    [ApiController]
    public class ReportsController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly AppSettings _appSettings;
        public string ConnectionStringName { get; set; } = "Default";
        
        public ReportsController(IWebHostEnvironment webHostEnvironment, AppSettings appSettings) 
        {
            _environment = webHostEnvironment;
            _appSettings = appSettings;
        }

        [HttpPost("export")]
        public async Task<IActionResult> ExportReport()
        {
            string rawJson;
            using (var reader = new StreamReader(Request.Body, Encoding.UTF8))
            {
                rawJson = await reader.ReadToEndAsync();
            }

            if (string.IsNullOrWhiteSpace(rawJson))
                return BadRequest("Request body is required");

            string apiUrl = _appSettings.BaseAddressForReports + "/api/report";

            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(3);
                var content = new StringContent(rawJson, Encoding.UTF8, "application/json");

                var response = await client.PostAsync(apiUrl, content);
                if (!response.IsSuccessStatusCode)
                {
                    var err = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, err);
                }

                var fileBytes = await response.Content.ReadAsByteArrayAsync();
                var contentType = response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
                var fileName = response.Content.Headers.ContentDisposition?.FileName?.Trim('"', '\'');

                // Parse format and report name safely from raw JSON
                string format = "";
                string reportName = "Report";
                try
                {
                    using var doc = JsonDocument.Parse(rawJson);
                    var root = doc.RootElement;
                    if (root.TryGetProperty("Format", out var f1)) format = f1.GetString() ?? "";
                    else if (root.TryGetProperty("format", out var f2)) format = f2.GetString() ?? "";

                    if (root.TryGetProperty("ReportName", out var r1)) reportName = r1.GetString() ?? "";
                    else if (root.TryGetProperty("reportName", out var r2)) reportName = r2.GetString() ?? "";
                }
                catch { }

                var fmt = format.ToLower();
                var baseName = Path.GetFileNameWithoutExtension(reportName);

                if (fmt == "word" || fmt == "doc" || fmt == "docx")
                {
                    contentType = "application/msword";
                }
                else if (fmt == "excel" || fmt == "xlsx")
                {
                    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    if (fmt == "excel" || fmt == "xlsx")
                        fileName = baseName + ".xlsx";
                    else if (fmt == "word" || fmt == "doc" || fmt == "docx")
                        fileName = baseName + ".doc";
                    else
                        fileName = baseName + ".pdf";
                }

                return File(fileBytes, contentType, fileName);
            }
        }
    }
}
