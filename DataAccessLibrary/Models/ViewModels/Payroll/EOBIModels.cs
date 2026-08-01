using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class EOBIRowModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string EOBINo { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string EmpStatus { get; set; } = "Permanent";
        public decimal Salary { get; set; }
        public int DaysWorked { get; set; }
        public decimal NetSalary { get; set; }
        public decimal EOBIAmt { get; set; }
        public bool IsSelected { get; set; } = true;
    }

    public class EOBIFilterDto
    {
        public int Month { get; set; } = DateTime.Today.Month;
        public int Year { get; set; } = DateTime.Today.Year;
        public string? DeptID { get; set; }
    }

    public class EOBISheetModel
    {
        public int Month { get; set; }
        public int Year { get; set; }
        public string? DeptID { get; set; }
        public bool IsSaved { get; set; }
        public List<EOBIRowModel> Rows { get; set; } = new();

        public decimal TotalSalary => Rows?.Where(r => r.IsSelected).Sum(r => r.Salary) ?? 0;
        public decimal TotalNetSalary => Rows?.Where(r => r.IsSelected).Sum(r => r.NetSalary) ?? 0;
        public decimal TotalEOBIAmt => Rows?.Where(r => r.IsSelected).Sum(r => r.EOBIAmt) ?? 0;
    }

    public class SaveEOBIDto
    {
        public int Month { get; set; }
        public int Year { get; set; }
        public string? DeptID { get; set; }
        public List<EOBIRowModel> Rows { get; set; } = new();
    }
}
