using CIP.Configurations;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.Extensions.Options;
using Microsoft.PowerBI.Api.Models;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Text;

namespace CIP.Controllers
{
    [Route("api/reports")]
    [ApiController]
    public class ReportsController : Controller
    {
        private IWebHostEnvironment _environment;
        private readonly AppSettings _appSettings;
        public string ConnectionStringName { get; set; } = "Default";
        
        public ReportsController(IWebHostEnvironment webHostEnvironment, AppSettings appSettings) 
        {
            _environment = webHostEnvironment;
            _appSettings = appSettings;
        }
        /*public IActionResult Index()
        {
            return View();
        }*/
        public class ReportRequest 
        {
            public string ReportName { get; set; }
            public string SelectionFormula { get; set; }
        }
        [HttpPost]
        public async Task<IActionResult> PrintReport([FromBody] ReportRequest request) 
        {
            if (request == null || string.IsNullOrEmpty(request.ReportName))
                return BadRequest("Report Name is required");

            string apiUrl = _appSettings.BaseAddressForReports + "/api/report";

            using (var client = new HttpClient())
            { 
                var json=JsonConvert.SerializeObject(request);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync(apiUrl, content);

                var pdfBytes=await response.Content.ReadAsByteArrayAsync();
                return File(pdfBytes, "application/pdf", "Report.pdf");

                // Option 2: if your API returns a URL, you can return JSON instead
                /*
                var resultJson = await response.Content.ReadAsStringAsync();
                return Content(resultJson, "application/json");
                */

            }
        }
        
    }
}
