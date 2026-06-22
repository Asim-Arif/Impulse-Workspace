using Dapper;
using DataAccessLibrary.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SqlClient;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Threading.Tasks;

namespace DataAccessLibrary
{
    public class ChartOfAccounts_DAC : ICIPData
    {
        private readonly string _connectionString;

        public ChartOfAccounts_DAC(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }
        public async Task<IEnumerable<CipLoungeEntry>> GetCIPsAsync()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT * FROM Lounge";
                return await connection.QueryAsync<CipLoungeEntry>(query);
            }
        }

        public async Task<int> UpdateCIPAsync(CipLoungeEntry cip)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = @"UPDATE CIP_Lounge_Cash_Entry 
                                       SET Date = @Date, PassengerName = @PassengerName, SeatNo = @SeatNo, 
                                           FlightNo = @FlightNo, Amount = @Amount 
                                       WHERE CIPInvoiceNo = @CIPInvoiceNo";
                return await connection.ExecuteAsync(query, cip);
            }
        }

        public async Task<int> DeleteCIPAsync(string cipInvoiceNo)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "DELETE FROM Lounge WHERE CIPInvoiceNo = @cipInvoiceNo";
                return await connection.ExecuteAsync(query, new { CIPInvoiceNo = cipInvoiceNo });
            }
        }

        public async Task<IEnumerable<CipLoungeEntry>> GetCIPByInvoiceNosAsync(string cipInvoiceNo)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT * FROM Lounge WHERE CIPInvoiceNo = @cipInvoiceNo";
                return await connection.QueryAsync<CipLoungeEntry>(query, new { cipInvoiceNo });
            }
        }

        public async Task<string> GetmaxCIPInvoiceNumberAsync(DateTime InvDT)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT dbo.GetMaxCIPInvoiceNo(@InvDT)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { InvDT });
            }
        }

        Task<CipLoungeEntry> ICIPData.GetCIPByInvoiceNosAsync(string invoiceno)
        {
            throw new NotImplementedException();
        }

        /// NOW INSERT 
        public async Task InsertCIPAsync(CipLoungeEntry item)
        {
            // Example logic for saving to the database
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    // SQL query to insert into Lounge table
                    string sql = @"
                        INSERT INTO Lounge (InvoiceNo, InvoiceDT, DT, UserName, MachineName, NoOfPerson, NetAmount, RcvdAmount,Source,Type,Remarks,BankRefID,DirectorRefID,Airline_Code,EmpID,AttachmentType) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@CIPInvoiceNo, @CIPDT, @CIPDT, @usernameforinvoices, @machineIP, @Totalperson, @TotalAmount, @RcvdAmount,@CipSourceType,0,@CipRemarks,@BankID,@Dir_RefID,@Airline_Code,@EmpID,@AttType)";

                    // Parameters for Lounge table
                    var loungeParams = new
                    {

                        CIPInvoiceNo = item.CIPInvoiceNo,
                        CIPDT = item.CIPDT,
                        Totalperson = item.CIPDetails.Count,
                        TotalAmount = item.TotalAmount,
                        RcvdAmount = item.RcvdAmount,
                        CipSourceType = item.CipSourceType,
                        CipRemarks = item.CipRemarks,
                        BankID = item.BankID,
                        Dir_RefID = item.Dir_RefID,
                        Airline_Code = item.Code,
                        EmpID = item.EmpID,
                        AttType = item.AttType,
                        usernameforinvoices = item.usernameforinvoices,
                        machineIP=item.machineIP,

                    };

                    // Execute the query using Dapper and get the generated EntryID
                    int iEntryID = await db.QuerySingleAsync<int>(sql, loungeParams);
                    // Insert each CIPDetail into the LoungeDetail table
                    var sno = 0;
                    foreach (var cipDetail in item.CIPDetails)
                    {
                        if (cipDetail == null)
                        {
                            Console.WriteLine("CIP Detail is null, skipping.");
                            continue;
                        }

                        sno++; // Increment serial number (SNo)

                        // SQL query to insert into LoungeDetail table
                        sql = @"
                                INSERT INTO LoungeDetail (RefID, SNo, PersonName, FlightNo_P, SeatNo_P, Amount,CardNo,authcode,BCardNo) 
                                VALUES (@CIP_RefID, @SNo, @PassengerName, @FlightNo, @SeatNo, @Amount, @CardNo,@authcode,@bankcardno)";

                        // Parameters for LoungeDetail table
                        var detailParams = new
                        {
                            CIP_RefID = iEntryID,
                            SNo = sno,  // Serial number for each entry
                            PassengerName = cipDetail.PassengerName,
                            FlightNo = cipDetail.FlightNo,
                            SeatNo = cipDetail.SeatNo,
                            Amount = cipDetail.Amount,
                            CardNo = cipDetail.CardNo,
                            authcode=cipDetail.authcode,
                            bankcardno=cipDetail.AuthCardNo,

                        };

                        // Execute the detail insert query
                        await db.ExecuteAsync(sql, detailParams);
                    }

                    Console.WriteLine("All CIPDetails inserted successfully.");
                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine($"SQL Error: {sqlEx.Message}");
                    throw; // Rethrow exception for further handling
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    throw; // Rethrow exception for further handling
                }
            }
        }

        public Task<List<CipLoungeEntry>> GetCustomersAsync()
        {
            throw new NotImplementedException();
        }

        public async Task<List<CipLoungeEntry>> GetCIPEntriesListAsync(DateTime fromDate, DateTime toDate, string sourcetype, string searchByField, string txtsearch,string selectsource, string sourcetypevalue)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string strcond = "";
                    if (txtsearch != "")
                    {
                        strcond = " AND " + searchByField + " LIKE '%" + txtsearch + "%' ";
                        //strcond = " AND " + searchByField + "='" + txtsearch + "' ";
                    }
                    //else 
                    //{
                    //    strcond = " ORDER BY DT ASC";
                    //}

                    if (sourcetypevalue !="") 
                    {
                        if (selectsource == "Airline")
                        {
                            strcond = strcond + " AND (@sourcetypevalue IS NULL OR Code = @sourcetypevalue) ";
                        }
                        else if (selectsource == "Bank")
                        {
                            strcond = strcond + " AND (@sourcetypevalue IS NULL OR BankID = @sourcetypevalue) ";
                        }
                        else if (selectsource == "Director Booking" || selectsource == "Director Personally")
                        {
                            strcond = strcond + " AND (@sourcetypevalue IS NULL OR SD_EntryID = @sourcetypevalue) ";
                        }
                    }

                    strcond = strcond + " ORDER BY DT ASC";

                    string sql = @"
                        SELECT 
                        DT AS CIPDT,
                        InvoiceNo AS CIPInvoiceNo,
                        ISNULL(Source, '') AS CipSource,
                        CASE 
                            WHEN ISNULL(Type, 0) = 0 THEN 'Cash' 
                            ELSE 'Other' 
                        END AS CipType,
                        NoOfPerson AS CipPersons,
                        NetAmount AS CipNetAmount,
                        RcvdAmount As TotalRcvdAmount,
                        AttachmentType as AttType,
                        PersonName AS PassengerName,
                        SeatNo,
                        FlightNo

                        FROM VLounge 
                
                        WHERE DT BETWEEN @fromDate AND @toDate
                        AND (@sourcetype IS NULL OR Source = @sourcetype) " + strcond ;

                    Console.WriteLine($"Error fetching CIP entries: {sql}");

                    var cipList = await db.QueryAsync<CipLoungeEntry>(sql, new { fromDate, toDate ,sourcetype, sourcetypevalue});
                    return cipList.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching CIP entries: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        private List<CipLoungeEntry> ConvertCipListToItemList(List<CIPList> cipList)
        {
            return cipList.Select(sb => new CipLoungeEntry
            {
                CIPInvoiceNo = sb.CIPInvoiceNo,
                CIPDT = sb.CIPDT,
                CipSource = sb.CipSource,
                CipType = sb.CipType,
                CipPersons = sb.CipPersons,
                CipNetAmount = sb.CipNetAmount,
                TotalRcvdAmount = sb.TotalRcvdAmount,

            }).ToList();
        }

        public async Task<string> getBankName(int BankID)
        {
            //Console.WriteLine($"Fetching bank name for BankID: {BankID}");
            const string query = "SELECT Bank AS BankName FROM BankList WHERE BankID = @BankID ";

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { BankID });

            }
        }

        public async Task<List<CipLoungeEntry>> GetDirectors()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT EntryID As Dir_RefID,CAST(FolioNo  AS VARCHAR(10)) + ' {' + SIALDirectors.DirectorName + '}' AS DirectorName,CAST(FolioNo  AS VARCHAR(10)) + ' {' + SIALDirectors.DirectorName + '}' AS CipSourceTypesForFilter_values FROM SIALDirectors WHERE Active=1 ORDER BY SIALDirectors.DirectorName ASC";

                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
                
            }
        }

        public async Task<List<CipLoungeEntry>> GetEmployees()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT EmpID,EmpID + ' ' + Name + ' {' + Designation + '}' as EmpName,EmpID + ' ' + Name + ' {' + Designation + '}' AS CipSourceTypesForFilter_values FROM Employees WHERE Active=1 ORDER BY Employees.Name ASC";
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }

        public async Task<List<CipLoungeEntry>> GetAirlines()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT Code,Name +'{ ' + Code +' }' As AirlineName,Name +'{ ' + Code +' }' AS CipSourceTypesForFilter_values FROM Airlines WHERE Customer=0 ORDER BY Airlines.Name ASC";
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }

        public async Task<List<CipLoungeEntry>> GetBanks()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT BankID,Bank +'{ ' + Branch +' }' As BankName,Bank +'{ ' + Branch +' }' AS CipSourceTypesForFilter_values FROM BankList WHERE ShowInCIP=1 ORDER BY BankList.Bank ASC";
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }

        public async Task<List<CipLoungeEntry>> GetSourceTypes()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT DISTINCT Source AS CipSourceTypesForFilter FROM Lounge WHERE ISNULL(Source,'')<>'' ORDER BY Lounge.Source ASC";
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }

        public async Task<List<CipLoungeEntry>> getseats()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = "SELECT DISTINCT SeatNo FROM Seats WHERE ISNULL(SeatNo,'')<>'' ORDER BY Seats.SeatNo ASC";
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }
        public async Task<List<CipLoungeEntry>> getflights()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                //string sql = "SELECT DISTINCT FlightNo FROM Flights WHERE ISNULL(FlightNo,'')<>'' ORDER BY Flights.FlightNo ASC";
                string sql = "SELECT DISTINCT FlightNo FROM VAirlines_DepartureFlightNos As Flights WHERE ISNULL(FlightNo,'')<>'' ORDER BY Flights.FlightNo ASC";
                
                return (await connection.QueryAsync<CipLoungeEntry>(sql)).ToList();
            }
        }

        public async Task<List<CipLoungeEntry>> getbankcards(int bankid)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string sql = @"
            SELECT DISTINCT CardNo 
            FROM BankCards 
            WHERE ISNULL(CardNo, '') <> '' 
              AND BankID = @bankid 
            ORDER BY CardNo ASC";

                // Pass the parameter to the query
                var parameters = new { bankid };

                return (await connection.QueryAsync<CipLoungeEntry>(sql, parameters)).ToList();
            }
        }
        public async Task<List<CipLoungeEntry>> GetCIPEntriesAsync_cash_credit(DateTime fromDate, DateTime toDate, string sourcetype,bool p_bshowprocessed)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    var condition = "";
                    if (p_bshowprocessed == true)
                    {
                        condition = "AND EntryID IN (SELECT LD_RefID FROM CIP_Invoices_Typewise_Detail) ORDER BY DT ASC";
                    }
                    else
                    {
                        condition = "AND EntryID NOT IN (SELECT LD_RefID FROM CIP_Invoices_Typewise_Detail) ORDER BY DT ASC";
                    }
                    string sql = @"
                    SELECT 
                        DT AS CIPDT,
                        InvoiceNo AS CIPInvoiceNo,
                        ISNULL(Source, '') AS CipSource,
                        NoOfPerson AS CipPersons,
                        NetAmount AS CipNetAmount,
                        RcvdAmount As TotalRcvdAmount,
                        AttachmentType as AttType,

                        PersonName as PassengerName,
                        SeatNo as SeatNo,
                        FlightNo as FlightNo,
                        Amount as Amount,
                        EntryID as EntryID

                        FROM VLoungeDetail 
                        WHERE DT BETWEEN @fromDate AND @toDate AND (Source = @sourcetype)" + @condition;
                    
                    //Console.WriteLine($"Error fetching CIP entries: {sourcetype}");

                    var cipList = await db.QueryAsync<CipLoungeEntry>(sql, new { fromDate, toDate, sourcetype });
                    return cipList.ToList();
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error fetching CIP entries: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        private List<CipLoungeEntry> ConvertCipListToItemList_Type(List<CIPList> cipList)
        {
            return cipList.Select(sb => new CipLoungeEntry
            {
                CIPInvoiceNo = sb.CIPInvoiceNo,
                CIPDT = sb.CIPDT,
                CipSource = sb.CipSource,
                CipType = sb.CipType,
                CipPersons = sb.CipPersons,
                CipNetAmount = sb.CipNetAmount,
                TotalRcvdAmount = sb.TotalRcvdAmount,

                AttType=sb.AttachmentType,

                PassengerName = sb.PassengerName,
                SeatNo = sb.SeatNo,
                FlightNo =sb.FlightNo,
                Amount =sb.Amount,
                EntryID = sb.EntryID


            }).ToList();
        }
        public async Task<string> GetmaxCIPInvoiceNumberAsync_typewise(DateTime InvDT, string sourcetype)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                const string query = "SELECT dbo.GetMaxCIPInvoiceNo_Typewise(@InvDT,@sourcetype)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { InvDT, sourcetype });
            }
        }
        /// NOW INSERT 
        public async Task SaveCIPInvoicesTypewiseAsync(CipLoungeEntry item)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = @"
                        INSERT INTO CIP_Invoices_Typewise (InvoiceNo, InvoiceDT, DTFrom, DTTo, Remarks, AttachmentType, EntryType, UserName, MachineName) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@CIPInvoiceNo, @CIPDT, @CIPDTFrom, @CIPDTTo, @CipRemarks, @AttType, @CipSourceType, @usernameforinvoices, @machineIP)";
                    
                    // Parameters for Lounge table
                    var loungeParams = new
                    {
                        CIPInvoiceNo = item.CIPInvoiceNo,
                        CIPDT = item.CIPDT,
                        CIPDTFrom = item.CIPDTFrom,
                        CIPDTTo = item.CIPDTTo,
                        CipRemarks = item.CipRemarks,
                        AttType = item.AttType,
                        CipSourceType = item.CipSourceType,
                        usernameforinvoices= item.CIP_Approval_ForwardBy,
                        machineIP = item.machineIP

                    };

                    // Execute the query using Dapper and get the generated EntryID
                    int iEntryID = await db.QuerySingleAsync<int>(sql, loungeParams);
                    foreach (var cipDetail in item.CIPDetails_tw)
                    {
                        sql = @"
                                INSERT INTO CIP_Invoices_Typewise_Detail (CIP_TW_RefID, LD_RefID) 
                                VALUES (@CIP_TW_RefID, @LD_RefID)";

                        var detailParams = new
                        {
                            CIP_TW_RefID = iEntryID,
                            LD_RefID = cipDetail.LD_RefID,
                        };
                        await db.ExecuteAsync(sql, detailParams);
                    }

                    // SAVE RECORDs FOR APPROVAL
                    sql = @"
                        INSERT INTO CIP_Invoices_Typewise_Approval(RefID,ForwardBy,ForwardTo,MachineName) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@CIP_Approval_RefID,@CIP_Approval_ForwardBy,@CIP_Approval_ForwardTo,'Developer')";

                    // Parameters for Lounge table
                    var loungeapp = new
                    {
                        CIP_Approval_RefID = iEntryID,
                        CIP_Approval_ForwardTo = item.CIP_Approval_ForwardTo,
                        CIP_Approval_ForwardBy = item.CIP_Approval_ForwardBy,
                    };
                    await db.QuerySingleAsync<int>(sql, loungeapp);
                    ////////////////////////////

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
        public async Task<List<CipLoungeEntry>> GetCipDetailsByInvoiceNoAsync(string invoiceno)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    var condition = "";

                    string sql = @"
                        SELECT 
                        
                        DT AS CIPDT,
                        InvoiceNo AS CIPInvoiceNo,
                        ISNULL(Source, '') AS CipSource,
                        NoOfPerson AS CipPersons,
                        NetAmount AS CipNetAmount,
                        RcvdAmount As TotalRcvdAmount,
                        AttachmentType as AttType,

                        PersonName as PassengerName,
                        SeatNo as SeatNo,
                        FlightNo as FlightNo,
                        Amount as Amount,
                        EntryID as EntryID,
                        Remarks As Remarks

                        FROM VLoungeDetail 
                        WHERE (InvoiceNo = @invoiceno)";

                    //Console.WriteLine($"Error fetching CIP entries: {sourcetype}");

                    var cipList = await db.QueryAsync<CipLoungeEntry>(sql, new { invoiceno });
                    return cipList.ToList();
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error fetching CIP entries: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        public async Task<string> GetSingleStringValueAsync(string strFieldName, string strTableName, string strCondition = "")
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                string query = $"SELECT {strFieldName} FROM {strTableName} {strCondition}";

                var result = await connection.ExecuteScalarAsync<object>(query);

                return result != null ? result.ToString() : "0";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return "0";
            }
        }

        public async Task<string> ExecuteMonthlyCIPSummaryAsync(DateTime dtFrom, DateTime dtTo)
        {
            using var connection = new SqlConnection(_connectionString);

            var parameters = new DynamicParameters();
            parameters.Add("@DTFrom", dtFrom);
            parameters.Add("@DTTo", dtTo);

            await connection.ExecuteAsync(
                "Print_Monthly_CIPUsageSummary_SP",
                parameters,
                commandType: CommandType.StoredProcedure
            );

            return "Stored procedure executed successfully.";
        }
        public async Task<List<CIPBanks>> getbanks()
        {
            using var connection = new SqlConnection(_connectionString);
            var sql = "SELECT BankID,Bank AS BankName,COA AS bCOA FROM BankList T1 WHERE ISNULL(ShowInCIP,0)=1 ORDER BY T1.COA DESC"; 
            var result = await connection.QueryAsync<CIPBanks>(sql);
            return result.ToList();
        }

        public async Task InsertImportedDirector(CipLoungeEntry item)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {

                    string sql = @"
                        INSERT INTO SIALDirectors (DirectorName,NICNo,FolioNo) 
                        OUTPUT INSERTED.EntryID
                        VALUES (@DirectorName, @CnicNo, @FolioNo)";

                    var ExcelParams = new 
                    {
                        DirectorName = item.DirectorName,
                        CnicNo = item.CnicNo,
                        FolioNo = item.FolioNo,

                    };
                    await db.QuerySingleAsync<int>(sql, ExcelParams);

                }
                catch (SqlException sqlEx)
                {
                    Console.WriteLine($"SQL Error: {sqlEx.Message}");
                    throw;
                }

            }
        }

        public async Task<int> UpdateDirector(int foliono=0)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                string query;

                if (foliono == 0)
                    query = "UPDATE SIALDirectors SET Active =0";
                else
                    query = "UPDATE SIALDirectors SET Active =1 WHERE FolioNo = @foliono";

                return await connection.ExecuteAsync(query, new { foliono });
            }
        }

        Task ICIPData.UpdateDirector(int foliono=0)
        {
            return UpdateDirector(foliono);
        }

        public async Task<List<CipLoungeEntry>> GetCIPEntriesListAsync_typewise(DateTime fromDate, DateTime toDate, string sourcetype="")
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string strcond = "";
                    if (sourcetype != "") {
                        strcond = strcond + "AND (@sourcetype IS NULL OR EntryType = @sourcetype)";
                    }

                    string sql = @"

                    SELECT 
                    DT AS CIPDT,
                    InvoiceNo AS CIPInvoiceNo,
                    EntryType AS CipType,
                    DTFrom AS CIPDTFrom,
                    DTTo AS CIPDTTo,
                    TotalAmount As TotalAmount,
                    ForwardBy as CIP_Approval_ForwardBy,
                    InvoiceRemarks as CipRemarks,
                    AttachmentType as AttType

                    FROM VCIP_Invoices_Typewise
                
                    WHERE DT BETWEEN @fromDate AND @toDate " + strcond;

                    //Console.WriteLine($"Error fetching CIP entries: {sourcetype}");
                    var cipList = await db.QueryAsync<CipLoungeEntry>(sql, new { fromDate, toDate, sourcetype });
                    return cipList.ToList();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching CIP Invoices: {ex.Message}");
                throw;
            }
        }

        public async Task<List<CipLoungeEntry>> GetCIPInvoiceTypewiseDetail(string invoiceno)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    var condition = "";

                    string sql = @"
                        SELECT 
                        
                        DT AS CIPDT,
                        CIPInvoiceNo AS CIPInvoiceNo,
                        ISNULL(Source, '') AS CipSource,
                        NoOfPerson AS CipPersons,
                        NetAmount AS CipNetAmount,
                        RcvdAmount As TotalRcvdAmount,
                        AttachmentType as AttType,

                        PersonName as PassengerName,
                        SeatNo as SeatNo,
                        FlightNo as FlightNo,
                        NetAmount as Amount,
                        EntryID as EntryID,
                        Remarks As Remarks

                        FROM VCIP_Invoices_Typewise_Detail 
                        WHERE (InvoiceNo = @invoiceno)";

                    var cipList = await db.QueryAsync<CipLoungeEntry>(sql, new { invoiceno });
                    return cipList.ToList();
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

    }
}