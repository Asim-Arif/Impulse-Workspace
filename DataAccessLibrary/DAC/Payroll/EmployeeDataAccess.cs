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

        public async Task<string> GetNextEmpIDAsync(string deptId)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "SELECT MAX(CAST(RIGHT(empid, LEN(empid) - CHARINDEX('-', empid)) AS INT)) FROM Employees WHERE deptid = @DeptID AND empid LIKE @DeptID + '-%'";
                    var maxId = await db.ExecuteScalarAsync<int?>(sql, new { DeptID = deptId });

                    int nextNum = (maxId ?? 0) + 1;
                    return $"{deptId}-{nextNum:D3}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting next employee ID for department {DeptID}", deptId);
                return string.Empty;
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
                            await SaveBaseEmployeeAsync(db, trans, employee, isAdd);
                            await SaveSettingsAsync(db, trans, employee);
                            await SaveQualificationsAsync(db, trans, employee);
                            await SaveGuarantorsAsync(db, trans, employee);
                            await SaveMedicalCheckUpsAsync(db, trans, employee);
                            await SaveMedicalHistoryAsync(db, trans, employee);
                            await SaveEmploymentHistoriesAsync(db, trans, employee);
                            await SaveJobsAsync(db, trans, employee);
                            await SaveDependantsAsync(db, trans, employee);
                            await SaveSuccessionsAsync(db, trans, employee);
                            await SaveReviewsAsync(db, trans, employee, userName, machineName);
                            await SaveNoticesAsync(db, trans, employee, userName, machineName);
                            await SaveDailyTargetsAsync(db, trans, employee, userName, machineName);
                            await SaveAssetsAsync(db, trans, employee);

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

        private async Task SaveBaseEmployeeAsync(IDbConnection db, IDbTransaction trans, EmployeeDto emp, bool isAdd)
        {
            var parameters = new
            {
                emp.EmpID, emp.DeptID, emp.Name, emp.EmpType, emp.Rel, emp.FName, emp.DOB, emp.Religion, emp.Age, emp.Sex,
                emp.Address, emp.NIC, emp.NTN, emp.Phone1, emp.Phone2, emp.Designation, emp.RefName, emp.RefFName,
                emp.RefPhone1, emp.RefPhone2, emp.RefAddress, emp.TempAddress, emp.IdentityMark, emp.LastEmpFirm,
                emp.LastEmpDesig, emp.LastEmpReason, emp.JoinDate, emp.Retirement, emp.Salary, emp.PProfit,
                emp.StartingSalary, emp.OldAge, emp.Social, emp.TaxPayee, emp.Remarks, emp.GInfo, emp.SocialNo,
                emp.SocialSal, emp.OldAgeNo, emp.OldAgeSal, emp.Status, emp.Active, emp.EmpPic, emp.Bonus,
                emp.ThumbExp, emp.Transport, emp.ExemptSettings, emp.BloodGroup, emp.ProcessID, emp.BankPymt,
                emp.ShowInExternal, emp.EmergencyName, emp.EmergencyFName, emp.EmergencyPhone1,
                emp.EmergencyPhone2, emp.EmergencyAddress, emp.OTByQty, emp.MinQty, emp.UnionFund, emp.Lunch,
                emp.TaxAmt, emp.InsurableDT, emp.ArmyRetired, emp.EmpAccNo, emp.PerPieceRate, emp.Rate,
                emp.MainContractorID, emp.EntryDT, emp.ProbationPeriod, emp.FixAllowance, emp.EOBIAmt,
                emp.SocialSecurityAmt, emp.EmpLunchAmt, emp.BondPeriod, emp.RelationWithCompEmp,
                emp.RelativeEmpID, emp.Sr, emp.IntDepartmentId, emp.IntJobTitleId, emp.IntGenderId,
                emp.IntGuarantorId, emp.EmpCodes, emp.IntMeritalStatusId, emp.IntBloodGroupId,
                emp.IntReligionId, emp.StrTitle, emp.StrCity, emp.StrProvince, emp.StrEmail, emp.StrFax,
                emp.StrCast, emp.TempEmp, emp.Pay_Full_Salary, emp.CNIC_PDF, emp.CNIC_PDF_FileName,
                emp.Caste, emp.Maslak
            };

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
                await db.ExecuteAsync(insertSql, parameters, transaction: trans);
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
                await db.ExecuteAsync(updateSql, parameters, transaction: trans);
            }
        }

        private async Task SaveSettingsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Settings != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpSettings WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                string sql = @"
                    INSERT INTO EmpSettings (
                        EmpID, EnterTime, ExitTime, LunchTime, AbsAllowed, TotalAbsAllowed,
                        TimeAllowed, NTimes, OverTime, SocialAmt, DinnerAmt, ShortTermPer,
                        AdvSalPer, LongTermTime, LongTermLimit, BGHoliday, AGHoliday,
                        BPHoliday, APHoliday, WorkingHrs, AllowMultiple, DeductAbsents,
                        DeductLHrs, LunchInTime, LunchOutTime
                    ) VALUES (
                        @EmpID, @EnterTime, @ExitTime, @LunchTime, @AbsAllowed, @TotalAbsAllowed,
                        @TimeAllowed, @NTimes, @OverTime, @SocialAmt, @DinnerAmt, @ShortTermPer,
                        @AdvSalPer, @LongTermTime, @LongTermLimit, @BGHoliday, @AGHoliday,
                        @BPHoliday, @APHoliday, @WorkingHrs, @AllowMultiple, @DeductAbsents,
                        @DeductLHrs, @LunchInTime, @LunchOutTime
                    )";
                var p = new
                {
                    employee.EmpID,
                    employee.Settings.EnterTime, employee.Settings.ExitTime, employee.Settings.LunchTime,
                    employee.Settings.AbsAllowed, employee.Settings.TotalAbsAllowed, employee.Settings.TimeAllowed,
                    employee.Settings.NTimes, employee.Settings.OverTime, employee.Settings.SocialAmt,
                    employee.Settings.DinnerAmt, employee.Settings.ShortTermPer, employee.Settings.AdvSalPer,
                    employee.Settings.LongTermTime, employee.Settings.LongTermLimit, employee.Settings.BGHoliday,
                    employee.Settings.AGHoliday, employee.Settings.BPHoliday, employee.Settings.APHoliday,
                    employee.Settings.WorkingHrs, employee.Settings.AllowMultiple, employee.Settings.DeductAbsents,
                    employee.Settings.DeductLHrs, employee.Settings.LunchInTime, employee.Settings.LunchOutTime
                };
                await db.ExecuteAsync(sql, p, transaction: trans);
            }
        }

        private async Task SaveQualificationsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Qualifications != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpQualification WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var q in employee.Qualifications)
                {
                    string sql = "INSERT INTO EmpQualification (EmpID, SrNo, Diploma, Institute, Year, Grad, Division) VALUES (@EmpID, @SrNo, @Diploma, @Institute, @Year, @Grad, @Division)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, q.SrNo, q.Diploma, q.Institute, q.Year, q.Grad, q.Division }, transaction: trans);
                }
            }
        }

        private async Task SaveGuarantorsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Guarantors != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpGuarantors WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var g in employee.Guarantors)
                {
                    string sql = "INSERT INTO EmpGuarantors (EmpID, Name, Address, PhoneNo, KnownBy) VALUES (@EmpID, @Name, @Address, @PhoneNo, @KnownBy)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, g.Name, g.Address, g.PhoneNo, g.KnownBy }, transaction: trans);
                }
            }
        }

        private async Task SaveMedicalCheckUpsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.MedicalCheckUps != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpMedicalCheckUps WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var m in employee.MedicalCheckUps)
                {
                    string sql = "INSERT INTO EmpMedicalCheckUps (EmpID, SrNo, MedicalTest, Lab, DT, Status, Treatment, Remarks) VALUES (@EmpID, @SrNo, @MedicalTest, @Lab, @DT, @Status, @Treatment, @Remarks)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, m.SrNo, m.MedicalTest, m.Lab, m.DT, m.Status, m.Treatment, m.Remarks }, transaction: trans);
                }
            }
        }

        private async Task SaveMedicalHistoryAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.MedicalHistory != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpMedicalHistory WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                string sql = "INSERT INTO EmpMedicalHistory (EmpID, HBsAg, AntiHcv, BloodSugar, BloodPressure, Allergy, AllergyDescription) VALUES (@EmpID, @HBsAg, @AntiHcv, @BloodSugar, @BloodPressure, @Allergy, @AllergyDescription)";
                await db.ExecuteAsync(sql, new { employee.EmpID, employee.MedicalHistory.HBsAg, employee.MedicalHistory.AntiHcv, employee.MedicalHistory.BloodSugar, employee.MedicalHistory.BloodPressure, employee.MedicalHistory.Allergy, employee.MedicalHistory.AllergyDescription }, transaction: trans);
            }
        }

        private async Task SaveEmploymentHistoriesAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.EmploymentHistories != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpEmployementHistory WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var h in employee.EmploymentHistories)
                {
                    string sql = "INSERT INTO EmpEmployementHistory (EmpID, FromDT, ToDT, Company, Designation, Industry, Responsibilities) VALUES (@EmpID, @FromDT, @ToDT, @Company, @Designation, @Industry, @Responsibilities)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, h.FromDT, h.ToDT, h.Company, h.Designation, h.Industry, h.Responsibilities }, transaction: trans);
                }
            }
        }

        private async Task SaveJobsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Jobs != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpJobs WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var j in employee.Jobs)
                {
                    string sql = "INSERT INTO EmpJobs (EmpID, SrNo, JobDescription) VALUES (@EmpID, @SrNo, @JobDescription)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, j.SrNo, j.JobDescription }, transaction: trans);
                }
            }
        }

        private async Task SaveDependantsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Dependants != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpDependants WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var d in employee.Dependants)
                {
                    string sql = "INSERT INTO EmpDependants (EmpID, Name, Relation, NICNo, MaritalStatus, DOB, Education, Class, Institute, Occupation, Male, Dependent, DOM, DOD) VALUES (@EmpID, @Name, @Relation, @NICNo, @MaritalStatus, @DOB, @Education, @Class, @Institute, @Occupation, @Male, @Dependent, @DOM, @DOD)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, d.Name, d.Relation, d.NICNo, d.MaritalStatus, d.DOB, d.Education, d.Class, d.Institute, d.Occupation, d.Male, d.Dependent, d.DOM, d.DOD }, transaction: trans);
                }
            }
        }

        private async Task SaveSuccessionsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Successions != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpSuccessions WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var s in employee.Successions)
                {
                    string sql = "INSERT INTO EmpSuccessions (EmpID, SuccessionEmpID) VALUES (@EmpID, @SuccessionEmpID)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, s.SuccessionEmpID }, transaction: trans);
                }
            }
        }

        private async Task SaveReviewsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee, string userName, string machineName)
        {
            if (employee.Reviews != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpReviews WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var r in employee.Reviews)
                {
                    string sql = "INSERT INTO EmpReviews (EmpID, DT, ReviewedBy, Review, UserName, MachineName, DTEntry, ReviewType, Rating) VALUES (@EmpID, @DT, @ReviewedBy, @Review, @UserName, @MachineName, GETDATE(), @ReviewType, @Rating)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, r.DT, r.ReviewedBy, r.Review, UserName = userName, MachineName = machineName, r.ReviewType, r.Rating }, transaction: trans);
                }
            }
        }

        private async Task SaveNoticesAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee, string userName, string machineName)
        {
            if (employee.Notices != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpNotices WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var n in employee.Notices)
                {
                    string sql = "INSERT INTO EmpNotices (EmpID, DT, NoticeBy, Notice, UserName, MachineName, DTEntry) VALUES (@EmpID, @DT, @NoticeBy, @Notice, @UserName, @MachineName, GETDATE())";
                    await db.ExecuteAsync(sql, new { employee.EmpID, n.DT, n.NoticeBy, n.Notice, UserName = userName, MachineName = machineName }, transaction: trans);
                }
            }
        }

        private async Task SaveDailyTargetsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee, string userName, string machineName)
        {
            if (employee.DailyTargets != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpDailyTargets WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var t in employee.DailyTargets)
                {
                    string sql = "INSERT INTO EmpDailyTargets (EmpID, GroupID, ProcessID, Qty, DTEntry, DTLastUpdated, UserName, MachineName, OverTime) VALUES (@EmpID, @GroupID, @ProcessID, @Qty, GETDATE(), GETDATE(), @UserName, @MachineName, @OverTime)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, t.GroupID, t.ProcessID, t.Qty, UserName = userName, MachineName = machineName, t.OverTime }, transaction: trans);
                }
            }
        }

        private async Task SaveAssetsAsync(IDbConnection db, IDbTransaction trans, EmployeeDto employee)
        {
            if (employee.Assets != null)
            {
                await db.ExecuteAsync("DELETE FROM EmpAssets WHERE EmpID = @EmpID", new { employee.EmpID }, transaction: trans);
                foreach (var a in employee.Assets)
                {
                    string sql = "INSERT INTO EmpAssets (EmpID, SrNo, AssetDescription, AssetDT, Remarks) VALUES (@EmpID, @SrNo, @AssetDescription, @AssetDT, @Remarks)";
                    await db.ExecuteAsync(sql, new { employee.EmpID, a.SrNo, a.AssetDescription, a.AssetDT, a.Remarks }, transaction: trans);
                }
            }
        }
    }
}
