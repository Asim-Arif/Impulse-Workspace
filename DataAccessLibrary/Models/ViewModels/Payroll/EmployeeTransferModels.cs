using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class EmployeeTransferDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string TransferFromDeptID { get; set; } = string.Empty;
        public string TransferToDeptID { get; set; } = string.Empty;
        public DateTime TransferDT { get; set; } = DateTime.Today;
        public string EntryDescription { get; set; } = string.Empty;
        public string UserName { get; set; } = "System";
    }

    public class EmployeeTransferHistoryRow
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string FromDeptName { get; set; } = string.Empty;
        public string ToDeptName { get; set; } = string.Empty;
        public DateTime? TransferDT { get; set; }
        public string EntryDescription { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
    }
}
