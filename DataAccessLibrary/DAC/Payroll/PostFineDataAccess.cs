using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class PostFineDataAccess : IPostFineDataAccess
    {
        private readonly string _connectionString;

        public PostFineDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task SaveFineAsync(PostFineDto dto)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                INSERT INTO EmpFine
                    (EmpID, FineDate, Amount, Description, Type)
                VALUES
                    (@EmpID, @FineDate, @Amount, @Description, @Type)";

            await db.ExecuteAsync(sql, new
            {
                EmpID       = dto.EmpID,
                FineDate    = dto.FineDate,
                Amount      = dto.Amount,
                Description = dto.Description ?? string.Empty,
                Type        = dto.Type
            });
        }
    }
}
