using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class DepartmentListItemModel
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool Active { get; set; }
        public int? Color { get; set; }
        public bool? TempDept { get; set; }
    }

    public class EmployeeListItemModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public float StartingSalary { get; set; }
        public DateTime? JoinDate { get; set; }
        public bool Active { get; set; }
        public bool TempEmp { get; set; }
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public byte[]? EmpPic { get; set; }
    }

    public class EmployeeInActiveDetailsModel
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string InActiveRemarks { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class EmployeeDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int? EmpType { get; set; }
        public string Rel { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public DateTime? DOB { get; set; }
        public string Religion { get; set; } = string.Empty;
        public int? Age { get; set; }
        public string Sex { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string NIC { get; set; } = string.Empty;
        public string NTN { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string RefName { get; set; } = string.Empty;
        public string RefFName { get; set; } = string.Empty;
        public string RefPhone1 { get; set; } = string.Empty;
        public string RefPhone2 { get; set; } = string.Empty;
        public string RefAddress { get; set; } = string.Empty;
        public string TempAddress { get; set; } = string.Empty;
        public string IdentityMark { get; set; } = string.Empty;
        public string LastEmpFirm { get; set; } = string.Empty;
        public string LastEmpDesig { get; set; } = string.Empty;
        public string LastEmpReason { get; set; } = string.Empty;
        public DateTime? JoinDate { get; set; }
        public string Retirement { get; set; } = string.Empty;
        public float? Salary { get; set; }
        public float? PProfit { get; set; }
        public float? StartingSalary { get; set; }
        public bool OldAge { get; set; }
        public bool Social { get; set; }
        public bool TaxPayee { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public string GInfo { get; set; } = string.Empty;
        public string SocialNo { get; set; } = string.Empty;
        public float? SocialSal { get; set; }
        public string OldAgeNo { get; set; } = string.Empty;
        public float? OldAgeSal { get; set; }
        public string Status { get; set; } = string.Empty;
        public bool Active { get; set; }
        public byte[]? EmpPic { get; set; }
        public bool Bonus { get; set; }
        public DateTime? DiscontinuedOn { get; set; }
        public byte[]? ThumbExp { get; set; }
        public bool Transport { get; set; }
        public bool? ExemptSettings { get; set; }
        public int EmployeeID { get; set; }
        public string BloodGroup { get; set; } = string.Empty;
        public string ProcessID { get; set; } = string.Empty;
        public bool? BankPymt { get; set; }
        public bool? ShowInExternal { get; set; }
        public string EmergencyName { get; set; } = string.Empty;
        public string EmergencyFName { get; set; } = string.Empty;
        public string EmergencyPhone1 { get; set; } = string.Empty;
        public string EmergencyPhone2 { get; set; } = string.Empty;
        public string EmergencyAddress { get; set; } = string.Empty;
        public bool? OTByQty { get; set; }
        public float? MinQty { get; set; }
        public bool? UnionFund { get; set; }
        public bool? Lunch { get; set; }
        public int? TaxAmt { get; set; }
        public DateTime? InsurableDT { get; set; }
        public bool? ArmyRetired { get; set; }
        public string EmpAccNo { get; set; } = string.Empty;
        public bool? PerPieceRate { get; set; }
        public float? Rate { get; set; }
        public string MainContractorID { get; set; } = string.Empty;
        public DateTime? EntryDT { get; set; }
        public string EmpIDVision { get; set; } = string.Empty;
        public string EmpIDOld { get; set; } = string.Empty;
        public int? ProbationPeriod { get; set; }
        public int? FixAllowance { get; set; }
        public int? EOBIAmt { get; set; }
        public int? SocialSecurityAmt { get; set; }
        public int? EmpLunchAmt { get; set; }
        public float? BondPeriod { get; set; }
        public string RelationWithCompEmp { get; set; } = string.Empty;
        public string RelativeEmpID { get; set; } = string.Empty;
        public int? Sr { get; set; }
        public int? IntDepartmentId { get; set; }
        public int? IntJobTitleId { get; set; }
        public int? IntGenderId { get; set; }
        public int? IntGuarantorId { get; set; }
        public int? EmpCodes { get; set; }
        public int? IntMeritalStatusId { get; set; }
        public int? IntBloodGroupId { get; set; }
        public int? IntReligionId { get; set; }
        public string StrTitle { get; set; } = string.Empty;
        public string StrCity { get; set; } = string.Empty;
        public string StrProvince { get; set; } = string.Empty;
        public string StrEmail { get; set; } = string.Empty;
        public string StrFax { get; set; } = string.Empty;
        public string StrCast { get; set; } = string.Empty;
        public bool? TempEmp { get; set; }
        public bool Pay_Full_Salary { get; set; }
        public byte[]? CNIC_PDF { get; set; }
        public string CNIC_PDF_FileName { get; set; } = string.Empty;
        public string Caste { get; set; } = string.Empty;
        public string Maslak { get; set; } = string.Empty;
    }

    public class DepartmentModel
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public bool Active { get; set; }
        public int? Color { get; set; }
        public string? Dept_ID { get; set; }
        public bool? TempDept { get; set; }
        public string? AccruedAccNo { get; set; }
        public string? SalaryAccNo { get; set; }
        public string? AccruedFoodAccNo { get; set; }
        public string? AccruedUnionAccNo { get; set; }
        public string? SchAccNo { get; set; }
        public string? AccruedEOBIAccNo { get; set; }
        public string? TaxAccNo { get; set; }
        public string? OnePercentAccNo { get; set; }
    }
}
