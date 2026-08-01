using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class GratuityRowModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public decimal Salary { get; set; }
        public DateTime JoinDate { get; set; }
        public DateTime? DiscontinuedOn { get; set; }
        public int MonthsWorked { get; set; }
        public int DaysWorked { get; set; }
        public int GratuityAmount { get; set; }
        public int? SavedGratuityAmount { get; set; }
        public bool IsModified { get; set; }
    }

    public class GratuitySheetModel
    {
        public int Year { get; set; }
        public bool IsPosted { get; set; }
        public List<GratuityRowModel> Rows { get; set; } = new();

        public decimal TotalSalary => Rows.Sum(r => r.Salary);
        public long TotalGratuityAmount => Rows.Sum(r => (long)r.GratuityAmount);
    }

    public class SaveGratuityDto
    {
        public int Year { get; set; }
        public List<GratuityRowModel> Rows { get; set; } = new();
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }
}
