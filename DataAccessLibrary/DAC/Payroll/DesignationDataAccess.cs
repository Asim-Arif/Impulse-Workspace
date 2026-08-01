using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Payroll
{
    public interface IDesignationDataAccess
    {
        Task<List<DesignationModel>> GetDesignationsAsync();
        Task<DesignationModel?> GetDesignationByIdAsync(int entryId);
        Task<bool> SaveDesignationAsync(DesignationModel model);
        Task<bool> DeleteDesignationAsync(int entryId);
    }

    public class DesignationDataAccess : IDesignationDataAccess
    {
        private readonly string _connectionString;

        public DesignationDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<DesignationModel>> GetDesignationsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    EntryID,
                    Designation,
                    ISNULL(Duties, '') AS Duties,
                    ISNULL(RequiredSkills, '') AS RequiredSkills,
                    ISNULL(RequiredTrainings, '') AS RequiredTrainings
                FROM Designations
                ORDER BY Designation";

            var list = await db.QueryAsync<DesignationModel>(sql);
            return list.ToList();
        }

        public async Task<DesignationModel?> GetDesignationByIdAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    EntryID,
                    Designation,
                    ISNULL(Duties, '') AS Duties,
                    ISNULL(RequiredSkills, '') AS RequiredSkills,
                    ISNULL(RequiredTrainings, '') AS RequiredTrainings
                FROM Designations
                WHERE EntryID = @entryId";

            return await db.QueryFirstOrDefaultAsync<DesignationModel>(sql, new { entryId });
        }

        public async Task<bool> SaveDesignationAsync(DesignationModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (model.EntryID == 0)
            {
                const string insertSql = @"
                    INSERT INTO Designations (Designation, Duties, RequiredSkills, RequiredTrainings)
                    VALUES (@Designation, @Duties, @RequiredSkills, @RequiredTrainings);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int newId = await db.ExecuteScalarAsync<int>(insertSql, model);
                model.EntryID = newId;
                return newId > 0;
            }
            else
            {
                const string updateSql = @"
                    UPDATE Designations
                    SET 
                        Designation = @Designation,
                        Duties = @Duties,
                        RequiredSkills = @RequiredSkills,
                        RequiredTrainings = @RequiredTrainings
                    WHERE EntryID = @EntryID";

                int rowsAffected = await db.ExecuteAsync(updateSql, model);
                return rowsAffected > 0;
            }
        }

        public async Task<bool> DeleteDesignationAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string deleteSql = "DELETE FROM Designations WHERE EntryID = @entryId";
            int rows = await db.ExecuteAsync(deleteSql, new { entryId });
            return rows > 0;
        }
    }
}
