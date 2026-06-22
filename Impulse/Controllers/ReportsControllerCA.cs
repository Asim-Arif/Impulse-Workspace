using CIP.Configurations;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.Extensions.Options;
using Microsoft.PowerBI.Api.Models;
using Newtonsoft.Json;
using System;
using System.Diagnostics;

namespace CIP.Controllers
{
    [Route("api/reports/ca")]
    [ApiController]
    public class ReportsControllerCA : Controller
    {
        private IWebHostEnvironment _environment;
        private readonly AppSettings _appSettings;
        public string ConnectionStringName { get; set; } = "Default";
        
        public ReportsControllerCA(IWebHostEnvironment webHostEnvironment, AppSettings appSettings) 
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
        public IActionResult PrintReport([FromBody] ReportRequest request) 
        {
            if (request == null || string.IsNullOrEmpty(request.ReportName)) 
            {
                return BadRequest("Export parameter is required");
            }
            string returnstatus = "Entered";
            string helperpath = "";
            Guid guid = Guid.NewGuid();
            
            string strExportPath= Path.Combine(_environment.WebRootPath, "Exported", guid + ".pdf");

            string strFilePath = Path.Combine(_environment.WebRootPath, "Reports", request.ReportName);
            /*strFilePath = "//Exported//bill.pdf";
            strFilePath = Path.Combine(_environment.WebRootPath, "Exported","bill.pdf");*/
            //string baseAddress = _appSettings.BaseAddress;
            string baseAddress = _appSettings.BaseAddressForReports;

            //strFilePath = Path.Combine(baseAddress, "Exported", "bill.pdf");
            string strFirstArgument = "Report";
            string strRecordSelectionFormula = request.SelectionFormula; //"{VendOrderDetail.OrderNo}='ORD-090623-022'";
            string arguments = $"\"{strFirstArgument}\" \"{strFilePath}\" \"{strRecordSelectionFormula}\" \"{strExportPath}\"";
            //First:Report,Second:Report Name;Third:Selection Formula;Fourth:Exported PDF Path 
            returnstatus = "Starting Process with "+arguments;

            int exitCode = 0;
            string Error = "Initial";
            try
            {
                Process process = new Process();
                //process.StartInfo.FileName = "\\Report_Helper.exe";
                helperpath = Path.Combine(_environment.WebRootPath, "Report_Helper.exe");
                process.StartInfo.FileName = Path.Combine(_environment.WebRootPath, "Report_Helper.exe");
                process.StartInfo.Arguments = arguments;
                process.StartInfo.RedirectStandardOutput = true;
                string anotherStatus = "Before Start";
                process.Start();
                anotherStatus = "After Start";
                process.WaitForExit();

                exitCode = process.ExitCode;
                Error = process.StandardOutput.ReadToEnd();                
                process.Close();
            }
            catch (Exception ex) 
            {
                returnstatus = ex.Message;
            }
            
            string strPathToShowPDF=Path.Combine(baseAddress, "Exported//"+guid+".pdf");
            string content = JsonConvert.SerializeObject(new { FilePath = strPathToShowPDF, returnstatus=returnstatus, helperpath=helperpath, Error=exitCode, anotherStatus = Error }, Formatting.None, new JsonSerializerSettings()
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            });
            return this.Content(content, "application/json");
            
            //return Json(new { filepath = "/Exported/bill.pdf" });
            /*Guid guid= Guid.NewGuid();
            string FilePath = "\\Exported\\"+guid+".pdf";
            string arguments = "report " + FilePath;
            Process process = new Process();
            process.StartInfo.FileName = this._environment.ContentRootPath + "\\wwwroot\\Report_Helper.exe";
            process.StartInfo.Arguments = arguments;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            process.WaitForExit();
            int exitCode = process.ExitCode;
            string Error=process.StandardOutput.ReadToEnd();
            process.Close();

            string content = JsonConvert.SerializeObject(new { FilePath = FilePath }, Formatting.None, new JsonSerializerSettings()
            {
                ReferenceLoopHandling= ReferenceLoopHandling.Ignore
            });
            return this.Content(content,"application/json");*/
            //return View();
        }
    }
}
