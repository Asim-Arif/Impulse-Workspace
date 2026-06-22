using System.Data;

namespace CIP.Services
{
    public interface IDBHelperService
    {
        int getSingleIntValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null);
        Task<DateTime?> getSingleDateValue(string strFieldName, string strTableName, string strCondition = "", object? parameters = null);
        int ExecuteDirectInsertUpdate(string strQuery);
        Task<string> getSingleStringValue(string strFieldName, string strTableName, string strCondition = "");
        Task<int> getSingleIntValueasync(string strFieldName, string strTableName, string strCondition = "");
        Task<string> GetMaxBillingInvoiceNo(DateTime DT, string strinvtype);
        Task<T> GetSingleValueAsync<T>(string strFieldName, string strTableName, string strCondition = "");
        Task<List<T>> GetListAsync<T>(string strFieldName, string strTableName, string strCondition = "", IDbConnection? existingDb = null, IDbTransaction? existingTrans = null);
        Task<string> getGeneralDataValueAsync(string strDataSetting);
        Task<string> getDatabasePasswordAsync(string strSetting);
        Task<int> ExecuteAsync(string sql, object? param = null);
        Task<double> getSingleDoubleValueasync(string strFieldName, string strTableName, string strCondition = "");

    }
}
