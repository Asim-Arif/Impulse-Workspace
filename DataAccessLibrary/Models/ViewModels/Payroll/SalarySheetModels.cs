using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Filter criteria for loading or calculating the monthly Salary Sheet.
    /// </summary>
    public class SalarySheetFilterDto
    {
        public int Year { get; set; } = DateTime.Today.Year;
        public int Month { get; set; } = DateTime.Today.Month;
        public string DeptID { get; set; } = "0";
        public string EmpID { get; set; } = "0";
        public bool IsExternal { get; set; } = false;
        public string? SearchText { get; set; }
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Individual employee row in the Salary Sheet grid.
    /// </summary>
    public class SalarySheetRowModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string? Designation { get; set; }

        public decimal BasicSalary { get; set; }
        public decimal DailyRate { get; set; }
        public double MonthDays { get; set; }

        public double PresentDays { get; set; }
        public double AbsentDays { get; set; }
        public double LeaveDays { get; set; }
        public double EarnedDays { get; set; }

        public decimal EarnedSalary { get; set; }
        public double OTHours { get; set; }
        public decimal OTAmount { get; set; }
        public decimal Allowances { get; set; }

        public double ShortHours { get; set; }
        public decimal ShortHoursDeduction { get; set; }

        public decimal GrossEarnings { get; set; }

        public decimal TaxDeduction { get; set; }
        public decimal NetEarnings { get; set; }

        public decimal ShortTermAdvanceDeduction { get; set; }
        public decimal LongTermLoanDeduction { get; set; }
        public decimal AdvSalaryPaid { get; set; }
        public decimal UnionFundDeduction { get; set; }
        public decimal FineDeduction { get; set; }
        public decimal EOBIDeduction { get; set; }
        public decimal GPHrsDeduction { get; set; }

        public decimal NetPayable { get; set; }

        public bool IsPaid { get; set; }
        public decimal PrevLTLoanBalance { get; set; }
        public bool IsFinalized { get; set; }

        // Extra details for finalization audit
        public double CasualLeaves { get; set; }
        public double SickLeaves { get; set; }
        public double AnnualLeaves { get; set; }
        public double CompensatoryLeaves { get; set; }
        public double WPLeaves { get; set; }
        public double MaternityLeaves { get; set; }

        // Additional MonthlySalaries schema properties for full 45-column persistence
        public double AbsentAllowed { get; set; }
        public decimal AbsentAllowedAmount { get; set; }
        public decimal AbsentDeductionAmount { get; set; }
        public double HrsPerDay { get; set; } = 8.0;
        public int TotalMonthHrs { get; set; }
        public double FakeWorkingHrs { get; set; }
        public decimal FakeRate { get; set; }
        public decimal FakeSalary { get; set; }
        public double SundayOTHrs { get; set; }
        public double SundayOTRate { get; set; }
        public decimal FixAllowance { get; set; }
        public decimal HoldSalaryAmt { get; set; }
        public double LateComingHrs { get; set; }
        public decimal AmtPaid { get; set; }
        public double GPHrs { get; set; }
        public int OTDinnerCount { get; set; }
        public decimal OTDinnerAmount { get; set; }
        public decimal DedOnePercent { get; set; }
        public decimal ZeroAbsentBonus { get; set; }
        public double OTHrs_Original { get; set; }
        public double LateHrs_Original { get; set; }
        public double OTHrs_Net { get; set; }
        public double LateHrs_Net { get; set; }

        // Legacy-layout persistence values (frmSalarySheet.frm) — populated during the
        // unfinalized calculation and used ONLY when saving to MonthlySalaries.
        // They have no effect on any salary calculation or on the grid.
        public double DisplayAbsentDays { get; set; }  // legacy display absent days (ADays = this x 8)
        public double LegacySalHrs { get; set; }       // legacy SalHrs (stored in SDays)
        public double LegacyLateHrs { get; set; }      // legacy dLateHrs, unclamped (stored as 'HH:MM' in LHrs)
        public double ShortHoursOnly { get; set; }     // legacy dShortHrs only (stored in ShortHrs)
        public decimal LateAmount { get; set; }        // legacy dLateAmt only (stored in LAmt)
        public decimal ShortAmount { get; set; }       // legacy dShortAmt only (stored in ShortHrsAmt)
    }

    /// <summary>
    /// Summary totals for KPI cards in Salary Sheet.
    /// </summary>
    public class SalarySheetKpiSummaryModel
    {
        public int TotalEmployees { get; set; }
        public decimal GrossEarnedSalary { get; set; }
        public decimal TotalOvertime { get; set; }
        public decimal TotalDeductions { get; set; }
        public decimal NetPayableSalary { get; set; }
    }

    /// <summary>
    /// Aggregated department totals and account mappings for salary Journal Voucher posting.
    /// </summary>
    public class DepartmentSalaryVoucherSummaryDto
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string AccruedAccNo { get; set; } = string.Empty;
        public string SalaryAccNo { get; set; } = string.Empty;
        public string AccruedFoodAccNo { get; set; } = string.Empty;
        public string AccruedUnionAccNo { get; set; } = string.Empty;
        public string SchAccNo { get; set; } = string.Empty;
        public string AccruedEOBIAccNo { get; set; } = string.Empty;
        public string TaxAccNo { get; set; } = string.Empty;
        public string OnePercentAccNo { get; set; } = string.Empty;

        public decimal TotalAmt { get; set; }
        public decimal SchAmt { get; set; }
        public decimal LunchAmt { get; set; }
        public decimal UnionAmt { get; set; }
        public decimal BalanceAmt { get; set; }
        public decimal ShortAmt { get; set; }
        public decimal LongAmt { get; set; }
        public decimal EOBIAmt { get; set; }
        public decimal TaxAmt { get; set; }
        public decimal OTAmt { get; set; }
        public decimal DedOnePercentAmt { get; set; }
        public decimal ZeroAbsentBonus { get; set; }
        public decimal GPHrsAmt { get; set; }
    }
}
