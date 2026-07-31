using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// DTO for posting an employee fine record into EmpFine.
    /// </summary>
    public class PostFineDto
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime FineDate { get; set; } = DateTime.Today;
        public decimal Amount { get; set; }
        public string? Description { get; set; }
        public string? Type { get; set; }
        public bool IsContractor { get; set; } = false;
    }
}
