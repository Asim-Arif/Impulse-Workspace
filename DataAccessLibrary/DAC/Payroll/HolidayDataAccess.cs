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
    public interface IHolidayDataAccess
    {
        Task<List<HolidayModel>> GetHolidaysByYearAsync(int year);
        Task<bool> AddHolidayAsync(HolidayModel model);
        Task<bool> UpdateHolidayAsync(DateTime dt, string description);
        Task<bool> DeleteHolidayAsync(DateTime dt);
    }

    public class HolidayDataAccess : IHolidayDataAccess
    {
        private readonly string _connectionString;

        public HolidayDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<HolidayModel>> GetHolidaysByYearAsync(int year)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT DT, Description 
                FROM Holidays 
                WHERE YEAR(DT) = @year 
                ORDER BY DT";

            var list = await db.QueryAsync<HolidayModel>(sql, new { year });
            return list.ToList();
        }

        public async Task<bool> AddHolidayAsync(HolidayModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                INSERT INTO Holidays (DT, Description)
                VALUES (@DT, @Description)";

            int rows = await db.ExecuteAsync(sql, model);
            return rows > 0;
        }

        public async Task<bool> UpdateHolidayAsync(DateTime dt, string description)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                UPDATE Holidays 
                SET Description = @description 
                WHERE CONVERT(VARCHAR(10), DT, 120) = CONVERT(VARCHAR(10), @dt, 120)";

            int rows = await db.ExecuteAsync(sql, new { dt, description });
            return rows > 0;
        }

        public async Task<bool> DeleteHolidayAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                DELETE FROM Holidays 
                WHERE CONVERT(VARCHAR(10), DT, 120) = CONVERT(VARCHAR(10), @dt, 120)";

            int rows = await db.ExecuteAsync(sql, new { dt });
            return rows > 0;
        }
    }
}
