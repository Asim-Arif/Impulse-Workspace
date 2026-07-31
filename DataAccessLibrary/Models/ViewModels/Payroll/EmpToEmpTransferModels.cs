using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// DTO passed from the Blazor page to the service/DAC layer when transferring
    /// Long Term loans or Short Term advances from one employee to another.
    /// </summary>
    public class EmpToEmpTransferDto
    {
        /// <summary>Source Employee ID (Transfer From / Outgoing)</summary>
        public string FromEmpID { get; set; } = string.Empty;

        /// <summary>Destination Employee ID (Transfer To / Incoming)</summary>
        public string ToEmpID { get; set; } = string.Empty;

        /// <summary>Transfer Date (stored in Advances.DT and AmtCleared.DT)</summary>
        public DateTime TransferDate { get; set; } = DateTime.Today;

        /// <summary>First month in which monthly deductions begin for Destination employee</summary>
        public DateTime DeductionStartDT { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

        /// <summary>Long Term loan amount to transfer</summary>
        public decimal LTAmt { get; set; }

        /// <summary>Monthly deduction for Destination employee's Long Term loan</summary>
        public decimal LTDeduct { get; set; }

        /// <summary>Short Term advance amount to transfer</summary>
        public decimal STAmt { get; set; }

        /// <summary>Description / Narration</summary>
        public string Description { get; set; } = "Employee to Employee Loan Transfer";

        /// <summary>Contractor flag</summary>
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Summary model displaying current outstanding LT & ST balances for the Source Employee.
    /// </summary>
    public class EmpTransferBalanceSummaryModel
    {
        /// <summary>Outstanding Long-Term loan balance (Type = 1)</summary>
        public decimal LTBalance { get; set; }

        /// <summary>Outstanding Short-Term advance balance (Type = 0)</summary>
        public decimal STBalance { get; set; }
    }
}
