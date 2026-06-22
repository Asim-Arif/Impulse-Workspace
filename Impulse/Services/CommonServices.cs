using ClosedXML.Excel;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components.Forms;
using System.Data;
using System.Xml.Linq;

namespace Impulse.Services
{
    public class CommonServices : ICommonServices
    {
        private readonly ICommonDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly ILogger<VouchersDataAccess> _logger;
        public CommonServices(ICommonDataAccess dataAccess, IAuditService auditService, ILogger<VouchersDataAccess> logger)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _logger = logger;
        }
        public async Task<int> ProcessBODMeetingAttendanceExcel(IBrowserFile file, DateTime meetingDate)
        {

            // 1. Ensure meeting exists
            int meetingID = await _dataAccess.GetOrCreateBODMeetingID(meetingDate);

            // 2. Open Excel
            using var stream = file.OpenReadStream();
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream);

            using var workbook = new XLWorkbook(memoryStream);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip Header

            foreach (var row in rows)
            {
                if (int.TryParse(row.Cell(1).GetValue<string>(), out int folioNo))
                {
                    // Fetch Director info from DB via DataAccess
                    var director = await _dataAccess.GetDirectorByFolio(folioNo);
                    if (director != null)
                    {
                        await _dataAccess.SaveBODMeetingParticipant(meetingID, folioNo, director);
                    }
                    else 
                    {
                        _logger.LogWarning("Import skipped: Folio {Folio} not found in database.", folioNo);
                        continue;
                    }
                }
            }
            return meetingID;
        }
        public string GetValueFromXMLString(string xml, string tagName)
        {
            if (string.IsNullOrWhiteSpace(xml)) return string.Empty;

            try
            {
                var doc = XDocument.Parse(xml);
                // This finds the tag regardless of where it is in the XML hierarchy
                // and ignores namespaces (local name only) to keep it simple
                var element = doc.Descendants().FirstOrDefault(x => x.Name.LocalName == tagName);

                return element?.Value ?? string.Empty;
            }
            catch (Exception ex)
            {
                // Equivalent to your MsgBox in VB6
                Console.WriteLine($"XML Parsing Error for {tagName}: {ex.Message}");
                return string.Empty;
            }
        }

    }
}
