using Dapper;
using DataAccessLibrary.Models;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data.SqlClient;
using OfficeOpenXml.Configuration;

namespace DataAccessLibrary
{
    public class PhotoCopyDataAccess : IPhotocopyData
    {
        private readonly string _connectionString;

        public PhotoCopyDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }
        public async Task<string> GetmaxInvoiceNumberAsync(DateTime InvDT)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT dbo.GetMaxPhotoCopyInvoiceNo(@InvDT)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { InvDT });
            }
        }

        public async Task<string> GetmaxInvoiceNumberAsync_ForInvoice(DateTime InvDT)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT [dbo].[GetMaxPhotocopyInvoiceNo_Invoice](@InvDT)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { InvDT });
            }
        }

        public async Task<List<CopyEntry>> GetCopyTypes()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT DISTINCT Description AS PhotoCopyType FROM PhotoCopyTypes ORDER BY PhotoCopyTypes.Description ASC";
                return (await connection.QueryAsync<CopyEntry>(sql)).ToList();
            }
        }

        public async Task InsertPhotoCopyAsync(CopyEntry item)
        {
            using (var con = new SqlConnection(_connectionString))
            {
                string sql = @"INSERT INTO PhotoCopies (DocNo,DT, Type, Qty, Rate,Remarks,UserName,MachineName) 
                                OUTPUT INSERTED.EntryID
                               VALUES (@CopyInvoiceNo,@PCDT, @CopyType, @Qty, @Rate,@Remarks,'PhotoCopy','PhotoCopy')";

                //int iEntryID = await con.QuerySingleAsync<int>(sql, item);
                await con.ExecuteAsync(sql, item);

            }
        }
        //public async Task<List<CopyEntry>> GetPhotoCopyEntriesListAsync(DateTime dt)
        //{
        //    try
        //    {
        //        using (IDbConnection db = new SqlConnection(_connectionString))
        //        {
        //            var dateOnly = dt.Date;

        //            string sql = @"
        //            SELECT DocNo AS CopyInvoiceNo,DT AS PCDT,Type AS CopyType,Qty AS Qty,Rate AS Rate,Remarks As Remarks 
        //            FROM PhotoCopies WHERE CAST(DT AS DATE)=@dt ORDER BY DT ASC";

        //            Console.WriteLine($"Error fetching CIP entries: {dt}");

        //            var entriesList = await db.QueryAsync<CopyEntry>(sql, new { DT= dateOnly });
        //            return entriesList.ToList();

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error fetching CIP entries: {ex.Message}");
        //        throw; // Rethrow or handle appropriately
        //    }
        //}

        public async Task<List<CopyEntry>> GetPhotoCopyEntriesListAsync(DateTime fromDate, DateTime toDate, int i_bshowprocessed)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    var condition = "";
                    if (i_bshowprocessed ==1)
                    {
                        condition = "AND EntryID IN (SELECT LD_RefID FROM PhotoCopy_Invoices_Detail) ORDER BY DT ASC";
                    }
                    else if (i_bshowprocessed == 2)
                    {
                        condition = "AND EntryID NOT IN (SELECT LD_RefID FROM PhotoCopy_Invoices_Detail) ORDER BY DT ASC";
                    }
                    else
                    {
                        condition = " ORDER BY DT ASC";
                    }

                    string sql = @"
                    SELECT EntryID,DocNo AS CopyInvoiceNo,DT AS PCDT,Type AS CopyType,Qty AS Qty,Rate AS Rate,Qty*Rate AS Amount,Remarks As Remarks 
                    FROM PhotoCopies 
                    WHERE DT BETWEEN @fromDate AND @toDate " + @condition;

                    var entriesList = await db.QueryAsync<CopyEntry>(sql, new { fromDate, toDate });
                    return entriesList.ToList();
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error fetching CIP entries: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        public async Task SavePhotoCopyInvoicesAsync(CopyEntry item)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = @"
                        INSERT INTO PhotoCopy_Invoices (InvoiceNo, InvoiceDT, DTFrom, DTTo, Remarks, AttachmentType, UserName, MachineName) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@CopyInvoiceNo, @PCDT, @DTFrom, @DTTo, @Remarks, @AttType, 'Developer', 'Developer')";

                    var loungeParams = new
                    {
                        CopyInvoiceNo = item.CopyInvoiceNo,
                        PCDT = item.PCDT,
                        DTFrom = item.DTFrom,
                        DTTo = item.DTTo,
                        Remarks = item.Remarks,
                        AttType = item.AttType,
                    };

                    // Execute the query using Dapper and get the generated EntryID
                    int iEntryID = await db.QuerySingleAsync<int>(sql, loungeParams);
                    foreach (var Detail in item.CIPDetails)
                    {
                        sql = @"
                                INSERT INTO PhotoCopy_Invoices_Detail(PC_RefID, LD_RefID) VALUES (@CIP_TW_RefID, @LD_RefID)";

                        var detailParams = new
                        {
                            CIP_TW_RefID = iEntryID,
                            LD_RefID = Detail.EntryID,
                        };
                        await db.ExecuteAsync(sql, detailParams);
                    }

                    // SAVE RECORDs FOR APPROVAL
                    sql = @"
                        INSERT INTO PhotoCopy_Invoices_Approval(RefID,ForwardBy,ForwardTo,MachineName) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@RefID,@ForwardBy,@ForwardTo,'Developer')";

                    // Parameters for Lounge table
                    var approvalentry = new
                    {
                        RefID = iEntryID,
                        ForwardTo = item.ForwardTo,
                        ForwardBy = item.ForwardBy ,
                    };
                    await db.QuerySingleAsync<int>(sql, approvalentry);
                    
                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine($"SQL Error: {sqlEx.Message}");
                    throw;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    throw;
                }
            }
        }

    }
}