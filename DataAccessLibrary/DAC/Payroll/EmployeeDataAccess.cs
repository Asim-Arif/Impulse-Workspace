using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Payroll
{
    public class EmployeeDataAccess : IEmployeeDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<EmployeeDataAccess> _logger;

        public EmployeeDataAccess(IConfiguration configuration, ILogger<EmployeeDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<DepartmentListItemModel>> GetDepartmentsAsync(bool showInactive, string? selectedDeptId = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT deptid AS DeptID, name AS Name, Active, Color, TempDept FROM Departments WHERE IsNull(TempDept, 0) = 0";
                    
                    if (!showInactive)
                    {
                        sql += " AND Active = 1";
                    }

                    if (!string.IsNullOrEmpty(selectedDeptId) && selectedDeptId != "0")
                    {
                        sql += " AND deptid = @DeptID";
                    }

                    sql += " ORDER BY name";

                    var result = await db.QueryAsync<DepartmentListItemModel>(sql, new { DeptID = selectedDeptId });
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching departments");
                throw;
            }
        }

        public async Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, bool showInactive, bool tempEmpOnly, string? selectedEmpId = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT EmpID, Name, FName, Designation, StartingSalary, JoinDate, Active, TempEmp, Phone1, Phone2, DeptID, DeptName FROM VEmp WHERE DeptID = @DeptID";

                    if (!showInactive)
                    {
                        sql += " AND Active = 1";
                    }

                    if (tempEmpOnly)
                    {
                        sql += " AND TempEmp = 1";
                    }
                    else
                    {
                        sql += " AND TempEmp = 0";
                    }

                    if (!string.IsNullOrEmpty(selectedEmpId) && selectedEmpId != "0")
                    {
                        sql += " AND EmpID = @EmpID";
                    }

                    sql += " ORDER BY EmpID";

                    var result = await db.QueryAsync<EmployeeListItemModel>(sql, new { DeptID = deptId, EmpID = selectedEmpId });
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employees for department {DeptID}", deptId);
                throw;
            }
        }

        public async Task<byte[]?> GetEmployeePictureAsync(string empId)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT EmpPic FROM Employees WHERE EmpID = @EmpID";
                    var pic = await db.ExecuteScalarAsync<byte[]>(sql, new { EmpID = empId });
                    return pic;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching picture for employee {EmpID}", empId);
                throw;
            }
        }

        public async Task SetEmployeeActiveStateAsync(string empId, bool active, DateTime? date = null, string? remarks = null, string? userName = null, string? machineName = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            if (!active)
                            {
                                // Deactivate employee and set DiscontinuedOn date
                                string updateSql = "UPDATE Employees SET Active = 0, DiscontinuedOn = @DT WHERE EmpID = @EmpID";
                                await db.ExecuteAsync(updateSql, new { EmpID = empId, DT = date ?? DateTime.Today }, transaction: trans);

                                // Insert inactive audit trail
                                string insertSql = @"
                                    INSERT INTO EmployeeInActiveDetails (EmpID, DT, InAciveRemarks, UserName, MachineName) 
                                    VALUES (@EmpID, @DT, @Remarks, @UserName, @MachineName)";
                                await db.ExecuteAsync(insertSql, new
                                {
                                    EmpID = empId,
                                    DT = date ?? DateTime.Today,
                                    Remarks = remarks ?? "Discontinued",
                                    UserName = userName ?? "System",
                                    MachineName = machineName ?? "Web"
                                }, transaction: trans);
                            }
                            else
                            {
                                // Re-activate employee, clear DiscontinuedOn
                                string updateSql = "UPDATE Employees SET Active = 1, DiscontinuedOn = NULL WHERE EmpID = @EmpID";
                                await db.ExecuteAsync(updateSql, new { EmpID = empId }, transaction: trans);
                            }

                            trans.Commit();
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating active status of employee {EmpID} to {Active}", empId, active);
                throw;
            }
        }

        public async Task SaveEmployeeAsync(EmployeeDto employee, bool isAdd, string userName, string machineName)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            if (isAdd)
                            {
                                string insertSql = @"
                                    INSERT INTO Employees (
                                        empid, deptid, name, EmpType, Rel, fname, dob, Religion, Age, sex,
                                        Address, NIC, NTN, Phone1, Phone2, Designation, refName, refFName,
                                        refPhone1, refphone2, refAddress, TempAddress, IdentityMark, lastEmpFirm,
                                        lastEmpDesig, lastempReason, JoinDate, retirement, Salary, PProfit,
                                        StartingSalary, oldage, social, taxPayee, remarks, GInfo, SocialNo,
                                        SocialSal, OldAgeNo, OldAgeSal, Status, Active, EmpPic, Bonus,
                                        ThumbExp, Transport, ExemptSettings, BloodGroup, ProcessID, BankPymt,
                                        ShowInExternal, EmergencyName, EmergencyFName, EmergencyPhone1,
                                        EmergencyPhone2, EmergencyAddress, OTByQty, MinQty, UnionFund, Lunch,
                                        TaxAmt, InsurableDT, ArmyRetired, EmpAccNo, PerPieceRate, Rate,
                                        MainContractorID, EntryDT, ProbationPeriod, FixAllowance, EOBIAmt,
                                        SocialSecurityAmt, EmpLunchAmt, BondPeriod, RelationWithCompEmp,
                                        RelativeEmpID, Sr, intDepartmentId, intJobTitleId, intGenderId,
                                        intGuarantorId, [Emp codes], intMeritalStatusId, intBloodGroupId,
                                        intReligionId, strTitle, strCity, strProvince, strEmail, strFax,
                                        strCast, TempEmp, Pay_Full_Salary, CNIC_PDF, CNIC_PDF_FileName,
                                        Caste, Maslak
                                    ) VALUES (
                                        @EmpID, @DeptID, @Name, @EmpType, @Rel, @FName, @DOB, @Religion, @Age, @Sex,
                                        @Address, @NIC, @NTN, @Phone1, @Phone2, @Designation, @RefName, @RefFName,
                                        @RefPhone1, @RefPhone2, @RefAddress, @TempAddress, @IdentityMark, @LastEmpFirm,
                                        @LastEmpDesig, @LastEmpReason, @JoinDate, @Retirement, @Salary, @PProfit,
                                        @StartingSalary, @OldAge, @Social, @TaxPayee, @Remarks, @GInfo, @SocialNo,
                                        @SocialSal, @OldAgeNo, @OldAgeSal, @Status, @Active, @EmpPic, @Bonus,
                                        @ThumbExp, @Transport, @ExemptSettings, @BloodGroup, @ProcessID, @BankPymt,
                                        @ShowInExternal, @EmergencyName, @EmergencyFName, @EmergencyPhone1,
                                        @EmergencyPhone2, @EmergencyAddress, @OTByQty, @MinQty, @UnionFund, @Lunch,
                                        @TaxAmt, @InsurableDT, @ArmyRetired, @EmpAccNo, @PerPieceRate, @Rate,
                                        @MainContractorID, @EntryDT, @ProbationPeriod, @FixAllowance, @EOBIAmt,
                                        @SocialSecurityAmt, @EmpLunchAmt, @BondPeriod, @RelationWithCompEmp,
                                        @RelativeEmpID, @Sr, @IntDepartmentId, @IntJobTitleId, @IntGenderId,
                                        @IntGuarantorId, @EmpCodes, @IntMeritalStatusId, @IntBloodGroupId,
                                        @IntReligionId, @StrTitle, @StrCity, @StrProvince, @StrEmail, @StrFax,
                                        @StrCast, @TempEmp, @Pay_Full_Salary, @CNIC_PDF, @CNIC_PDF_FileName,
                                        @Caste, @Maslak
                                    )";
                                await db.ExecuteAsync(insertSql, employee, transaction: trans);
                            }
                            else
                            {
                                string updateSql = @"
                                    UPDATE Employees SET
                                        deptid = @DeptID, name = @Name, EmpType = @EmpType, Rel = @Rel, fname = @FName,
                                        dob = @DOB, Religion = @Religion, Age = @Age, sex = @Sex, Address = @Address,
                                        NIC = @NIC, NTN = @NTN, Phone1 = @Phone1, Phone2 = @Phone2, Designation = @Designation,
                                        refName = @RefName, refFName = @RefFName, refPhone1 = @RefPhone1, refphone2 = @RefPhone2,
                                        refAddress = @RefAddress, TempAddress = @TempAddress, IdentityMark = @IdentityMark,
                                        lastEmpFirm = @LastEmpFirm, lastEmpDesig = @LastEmpDesig, lastempReason = @LastEmpReason,
                                        JoinDate = @JoinDate, retirement = @Retirement, Salary = @Salary, PProfit = @PProfit,
                                        StartingSalary = @StartingSalary, oldage = @OldAge, social = @Social, taxPayee = @TaxPayee,
                                        remarks = @Remarks, GInfo = @GInfo, SocialNo = @SocialNo, SocialSal = @SocialSal,
                                        OldAgeNo = @OldAgeNo, OldAgeSal = @OldAgeSal, Status = @Status, Active = @Active,
                                        EmpPic = @EmpPic, Bonus = @Bonus, ThumbExp = @ThumbExp, Transport = @Transport,
                                        ExemptSettings = @ExemptSettings, BloodGroup = @BloodGroup, ProcessID = @ProcessID,
                                        BankPymt = @BankPymt, ShowInExternal = @ShowInExternal, EmergencyName = @EmergencyName,
                                        EmergencyFName = @EmergencyFName, EmergencyPhone1 = @EmergencyPhone1,
                                        EmergencyPhone2 = @EmergencyPhone2, EmergencyAddress = @EmergencyAddress,
                                        OTByQty = @OTByQty, MinQty = @MinQty, UnionFund = @UnionFund, Lunch = @Lunch,
                                        TaxAmt = @TaxAmt, InsurableDT = @InsurableDT, ArmyRetired = @ArmyRetired,
                                        EmpAccNo = @EmpAccNo, PerPieceRate = @PerPieceRate, Rate = @Rate,
                                        MainContractorID = @MainContractorID, EntryDT = @EntryDT,
                                        ProbationPeriod = @ProbationPeriod, FixAllowance = @FixAllowance, EOBIAmt = @EOBIAmt,
                                        SocialSecurityAmt = @SocialSecurityAmt, EmpLunchAmt = @EmpLunchAmt,
                                        BondPeriod = @BondPeriod, RelationWithCompEmp = @RelationWithCompEmp,
                                        RelativeEmpID = @RelativeEmpID, Sr = @Sr, intDepartmentId = @IntDepartmentId,
                                        intJobTitleId = @IntJobTitleId, intGenderId = @IntGenderId, intGuarantorId = @IntGuarantorId,
                                        [Emp codes] = @EmpCodes, intMeritalStatusId = @IntMeritalStatusId,
                                        intBloodGroupId = @IntBloodGroupId, intReligionId = @IntReligionId, strTitle = @StrTitle,
                                        strCity = @StrCity, strProvince = @StrProvince, strEmail = @StrEmail, strFax = @StrFax,
                                        strCast = @StrCast, TempEmp = @TempEmp, Pay_Full_Salary = @Pay_Full_Salary,
                                        CNIC_PDF = @CNIC_PDF, CNIC_PDF_FileName = @CNIC_PDF_FileName, Caste = @Caste, Maslak = @Maslak
                                    WHERE empid = @EmpID";
                                await db.ExecuteAsync(updateSql, employee, transaction: trans);
                            }
                            trans.Commit();
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving employee {EmpID}", employee.EmpID);
                throw;
            }
        }
    }
}
