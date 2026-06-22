using Dapper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary
{
    public class DBHelper : IDBHelper
    {

        private readonly IConfiguration _config;
        public string ConnectionStringName { get; set; } = "Default";
        public DBHelper(IConfiguration config)
        {
            _config = config;
        }
        public int getSingleIntValue(string strFieldName, string strTableName, string strCondition = "", object? parameters=null)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");
            int iResult = 0;
            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = "SELECT " + strFieldName + " FROM " + strTableName + " " + strCondition;
                var obj = connection.ExecuteScalar(strSQL,parameters);
                if (obj != null)
                {
                    iResult = Convert.ToInt32(obj);
                }
                else
                {
                    iResult = 0;
                }
            }

            return iResult;
        }
        public async Task<DateTime?> getSingleDateValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");
            DateTime? myDate = null;
            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = "SELECT " + strFieldName + " FROM " + strTableName + " " + strCondition;
                var obj =await connection.ExecuteScalarAsync(strSQL, parameters);
                if (obj != null && obj != DBNull.Value)
                {
                    myDate = Convert.ToDateTime(obj);
                }
                else
                {
                    myDate = null;
                }
            }

            return myDate;
        }
        public int ExecuteDirectInsertUpdate(string strQuery)
        {
            string connectionString = _config.GetConnectionString(ConnectionStringName);
            int iRowsAffected = 0;

            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                iRowsAffected = connection.Execute(strQuery);
            }

            return iRowsAffected;
        }

        public async Task<string> getSingleStringValue(string strFieldName, string strTableName, string strCondition = "")
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");

            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";
                var result = await connection.ExecuteScalarAsync<string>(strSQL);

                return result ?? string.Empty; // return empty string if null
            }
        }

        public async Task<int> getSingleIntValueasync(string strFieldName, string strTableName, string strCondition = "")
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");

            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";

                var result = await connection.ExecuteScalarAsync<object>(strSQL);

                if (result != null && int.TryParse(result.ToString(), out int count))
                {
                    return count;
                }
                return 0;
            }
        }

        public async Task<T> GetSingleValueAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null,IDbTransaction? existingTrans = null)
        {
            string strSQL = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";
            bool isInternalConn = (existingDb == null);
            IDbConnection db = existingDb ?? new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            try
            {
                if (db.State != ConnectionState.Open) db.Open();
                return await db.ExecuteScalarAsync<T>(strSQL, transaction: existingTrans);
            }
            finally 
            {
                if (isInternalConn) db.Dispose();
            }
        }
        public async Task<List<T>> GetListAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null, IDbTransaction? existingTrans = null)
        {
            string strSQL = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";
            bool isInternalConn = (existingDb == null);
            IDbConnection db = existingDb ?? new SqlConnection(_config.GetConnectionString("DefaultConnection"));

            try
            {
                if (db.State != ConnectionState.Open) db.Open();

                // QueryAsync returns the entire result set as the specified type T
                var result = await db.QueryAsync<T>(strSQL, transaction: existingTrans);
                return result.ToList();
            }
            finally
            {
                if (isInternalConn) db.Dispose();
            }
        }
        public async Task<string> GetMaxBillingInvoiceNo(DateTime InvDT, string strinvtype)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");

            using (var connection = new SqlConnection(connectionString))
            {
                string query = $"SELECT dbo.GetMaxBillingInvoiceNo(@InvDT, @strinvtype)";

                try
                {
                    var result = await connection.ExecuteScalarAsync<string>(query, new
                    {
                        InvDT,
                        strinvtype
                    });

                    return result;
                }
                catch (Exception ex)
                {
                    throw new Exception("Error fetching max invoice number", ex);
                }
            }
        }
        public async Task<string> getGeneralDataValueAsync(string strDataSetting)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");
            string strResult = string.Empty;
            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = "SELECT DataValue FROM GeneralData WHERE DataName=@Setting";
                strResult = await connection.ExecuteScalarAsync<string>(strSQL,new {Setting=strDataSetting});

                return strResult ?? string.Empty;
            }
        }
        public async Task<string> getDatabasePasswordAsync(string strSetting)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");
            string strResult = string.Empty;
            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                string strSQL = "SELECT Password FROM Security WHERE SettingName=@Setting";
                strResult = await connection.ExecuteScalarAsync<string>(strSQL, new { Setting = strSetting });

                return strResult ?? string.Empty;
            }
        }
        public async Task<int> ExecuteAsync(string sql, object? param = null)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");
            using (IDbConnection connection = new SqlConnection(connectionString))
            {                
                return await connection.ExecuteAsync(sql, param);
            }
            
        }

        public async Task<bool> getSingleBoolValueasync(string strFieldName, string strTableName, string strCondition = "")
        {
            using IDbConnection db = new SqlConnection(_config.GetConnectionString("DefaultConnection"));

            string sql = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";

            var result = await db.ExecuteScalarAsync<bool?>(sql);

            return result ?? false;
        }

        public async Task<double> getSingleDoubleValueasync(string strFieldName, string strTableName, string strCondition = "")
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");

            using (IDbConnection connection = new SqlConnection(connectionString))
            {
             
                string strSQL = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";
                var result = await connection.ExecuteScalarAsync<object>(strSQL);

                if (result != null && result != DBNull.Value)
                {
                    if (double.TryParse(result.ToString(), out double value))
                    {
                        return value;
                    }
                }

                return 0;
            }
        }

        public async Task<IEnumerable<T>> getListasync<T>(string sql, object parameters = null)
        {
            string connectionString = _config.GetConnectionString("DefaultConnection");

            using (IDbConnection connection = new SqlConnection(connectionString))
            {
                return await connection.QueryAsync<T>(sql, parameters);
            }
        }
    }
}
