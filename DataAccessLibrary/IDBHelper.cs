using System.Data;

namespace DataAccessLibrary
{
    public interface IDBHelper
    {
        string ConnectionStringName { get; set; }
        
        int getSingleIntValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null);
        Task<DateTime?> getSingleDateValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null);
        int ExecuteDirectInsertUpdate(string strQuery);
        Task<string> getSingleStringValue(string strFieldName, string strTableName, string strCondition = "");
        Task<int> getSingleIntValueasync(string strFieldName, string strTableName, string strCondition = "");
        Task<string > GetMaxBillingInvoiceNo(DateTime DT, string strinvtype);
        Task<T> GetSingleValueAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null, IDbTransaction? existingTrans = null);
        Task<List<T>> GetListAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null, IDbTransaction? existingTrans = null);
        Task<string> getGeneralDataValueAsync(string strDataSetting);
        Task<string> getDatabasePasswordAsync(string strSetting);
        Task<int> ExecuteAsync(string sql, object? param = null);
        Task<bool> getSingleBoolValueasync(string strFieldName, string strTableName, string strCondition = "");

        Task<double> getSingleDoubleValueasync(string strFieldName, string strTableName, string strCondition = "");
        Task<IEnumerable<T>> getListasync<T>(string sql, object parameters = null);

    }

}