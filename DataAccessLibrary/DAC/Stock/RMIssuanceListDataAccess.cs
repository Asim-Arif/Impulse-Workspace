using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Dapper;
using Microsoft.Extensions.Configuration;
using DataAccessLibrary.Models.ViewModels.Stock;
using DataAccessLibrary.Interface.Stock;

namespace DataAccessLibrary.DAC.Stock
{
    public class RMIssuanceListDataAccess : IRMIssuanceListDataAccess
    {
        private readonly string _connectionString;

        public RMIssuanceListDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        public async Task<IEnumerable<RMGroupLookupModel>> GetRMGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT ID, Description FROM RMGroups";
            return await db.QueryAsync<RMGroupLookupModel>(sql);
        }

        public async Task<IEnumerable<DepartmentLookupModel>> GetDepartmentsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT DeptID, Name FROM Departments";
            return await db.QueryAsync<DepartmentLookupModel>(sql);
        }

        public async Task<IEnumerable<RMIssuanceListRowModel>> GetRMIssuanceListAsync(
            DateTime dtFrom, 
            DateTime dtTo, 
            string? materialId = null, 
            int? groupId = null, 
            string? employeeId = null, 
            int? vendorId = null, 
            string? deptId = null)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            string sql = @"
                SELECT l.IssNo, l.DT,
                       ISNULL(NULLIF(l.VenderName, ''), vm.AccTitle) AS VenderName,
                       l.IssdToEmpID, l.RM, l.Total_Issued_Qty
                FROM VRawMaterialIssuanceList l
                INNER JOIN RawMaterialIssuance rmi ON l.IssNo = rmi.IssNo
                LEFT JOIN VVendersAndMakers vm ON rmi.Vendor_AccNo = vm.AccNo
                WHERE l.DT BETWEEN @DtFrom AND @DtTo";
            
            if (!string.IsNullOrEmpty(materialId))
            {
                sql += " AND l.IssNo IN (SELECT IssNo FROM RawMaterialIssuanceDetail WHERE RMID1 = @MaterialId)";
            }
            else if (groupId.HasValue)
            {
                sql += " AND l.IssNo IN (SELECT IssNo FROM RawMaterialIssuanceDetail INNER JOIN RM ON RM.RMID1 = RawMaterialIssuanceDetail.RMID1 WHERE GroupID = @GroupId)";
            }

            if (!string.IsNullOrEmpty(employeeId))
            {
                sql += " AND l.IssdToEmpIDOnly = @EmployeeId";
            }
            else if (vendorId.HasValue)
            {
                sql += " AND (l.VendID = @VendorId OR vm.VendID = @VendorId)";
            }
            else if (!string.IsNullOrEmpty(deptId))
            {
                sql += " AND l.deptid = @DeptId";
            }

            sql += " ORDER BY l.DT, l.IssNo";

            return await db.QueryAsync<RMIssuanceListRowModel>(sql, new 
            { 
                DtFrom = dtFrom, 
                DtTo = dtTo, 
                MaterialId = materialId, 
                GroupId = groupId, 
                EmployeeId = employeeId, 
                VendorId = vendorId, 
                DeptId = deptId 
            });
        }

        public async Task<bool> DeleteIssuanceAsync(string issNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            int rows = await db.ExecuteAsync("UpdateMaterialLocation_Deletion_SP", new { IssNo = issNo }, commandType: CommandType.StoredProcedure);
            return rows > 0;
        }
    }
}
