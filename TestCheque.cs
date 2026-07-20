using System;
using System.Data.SqlClient;
using Dapper;

namespace TestCheque
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=.;Initial Catalog=Impulse;Integrated Security=True;TrustServerCertificate=True";
            double sNo = 2310190001;
            try
            {
                using (var db = new SqlConnection(connectionString))
                {
                    string sql = @"
                    SELECT CDate, chequeno AS ChequeNo, Bank, Branch, Description, Amount, 
                           chequeType AS ChequeType, chequeDate AS ChequeDate, Posted, 
                           Bounced, ClearanceDT, Payment
                    FROM VChqLedger
                    WHERE Vouchers_SNo = @sNo OR SNo = @sNo";

                    var result = db.QueryFirstOrDefault(sql, new { sNo });
                    Console.WriteLine(result != null ? "Success" : "Not Found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.ToString());
            }
        }
    }
}
