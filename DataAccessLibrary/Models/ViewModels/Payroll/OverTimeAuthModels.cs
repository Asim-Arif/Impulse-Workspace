using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class OverTimeAuthRowDto
    {
        public int EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Rel { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public decimal StartingSalary { get; set; }
        public decimal OtRate { get; set; }
        public double OtHrs { get; set; }
        public decimal Amount => Math.Round(OtRate * (decimal)OtHrs, 2);
        public bool IsAuthorized { get; set; }
    }
}
