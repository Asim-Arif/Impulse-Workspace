using Dapper;
using DataAccessLibrary.Models;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC
{
    //START FOR CUSTOMERS
    public class CustomersData : ICustData
    {
        private readonly IConfiguration _config;
        public CustomersData(IConfiguration config)
        {
            _config = config;
        }
        private string GetConnectionString()
        {
            return _config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<Customers>> GetCustomersAsync()
        {
            using (IDbConnection db = new SqlConnection(GetConnectionString()))
            {
                string sql = "SELECT CustCode,Name As CustomerName FROM ForeignCustomers";
                return (await db.QueryAsync<Customers>(sql)).ToList();
            }
        }
    }
}
