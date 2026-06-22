using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using Microsoft.PowerBI.Api.Models;
using System.Data;

namespace CIP.Services
{
    public class DBHelperService : IDBHelperService
    {
        private readonly IDBHelper _dbHelper;
        private Encryption encryption = new Encryption();
        public DBHelperService(IDBHelper dbHelper)
        {
            _dbHelper = dbHelper;            
        }
        public int getSingleIntValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null)
        { 
            return _dbHelper.getSingleIntValue(strFieldName, strTableName, strCondition,parameters);
        }
        public async Task<DateTime?> getSingleDateValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null)
        {
            return await _dbHelper.getSingleDateValue(strFieldName, strTableName, strCondition, parameters);
        }
        public async Task<string> getGeneralDataValueAsync(string strDataSetting)
        {
            return await _dbHelper.getGeneralDataValueAsync(strDataSetting);
        }
        public async Task<string> getDatabasePasswordAsync(string strDataSetting)
        {
            string strRawPassword= await _dbHelper.getDatabasePasswordAsync(strDataSetting);
            string strPassword= encryption.CryptRC4(strRawPassword, "awmsial");

            return strPassword;
        }
        public int ExecuteDirectInsertUpdate(string strQuery)
        { 
            return _dbHelper.ExecuteDirectInsertUpdate(strQuery);
        }
        public async Task<string> getSingleStringValue(string strFieldName, string strTableName, string strCondition = "")
        { 
            return await _dbHelper.getSingleStringValue(strFieldName, strTableName, strCondition);
        }
        public async Task<int> getSingleIntValueasync(string strFieldName, string strTableName, string strCondition = "")
        { 
            return await _dbHelper.getSingleIntValueasync(strFieldName, strTableName, strCondition);
        }
        public async Task<T> GetSingleValueAsync<T>(string strFieldName, string strTableName, string strCondition = "")
        {
            return await _dbHelper.GetSingleValueAsync<T>(strFieldName, strTableName, strCondition);
        }
        public async Task<List<T>> GetListAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null, IDbTransaction? existingTrans = null) 
        {
            return await _dbHelper.GetListAsync<T>(strFieldName,strTableName,strCondition, existingDb,existingTrans);
        }
        public async Task<string> GetMaxBillingInvoiceNo(DateTime InvDT, string strinvtype)
        {
            return await _dbHelper.GetMaxBillingInvoiceNo(InvDT, strinvtype);
        }
        public async Task<int> ExecuteAsync(string sql, object? param = null) 
        {
            return await _dbHelper.ExecuteAsync(sql, param);
        }
        public async Task<double> getSingleDoubleValueasync(string strFieldName, string strTableName, string strCondition = "")
        {
            return await _dbHelper.getSingleDoubleValueasync(strFieldName, strTableName, strCondition);
        }

    }
}
