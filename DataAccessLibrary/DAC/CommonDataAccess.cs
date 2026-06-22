using Dapper;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary.DAC
{
    public class CommonDataAccess : ICommonDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<VouchersDataAccess> _logger;
        private readonly IDBHelper _db;
        public CommonDataAccess(IConfiguration configuration, ILogger<VouchersDataAccess> logger, IDBHelper db)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");

            _logger = logger;
            _db = db;
        }
        public async Task<int> GetOrCreateBODMeetingID(DateTime meetingDate)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Check if exists
                string selectSql = "SELECT EntryID FROM BOD_Meetings WHERE DT = @DT";
                int id = await db.ExecuteScalarAsync<int>(selectSql, new { DT = meetingDate });

                if (id > 0) return id;

                // Otherwise Insert and return new ID
                string insertSql = "INSERT INTO BOD_Meetings (DT) VALUES (@DT); SELECT SCOPE_IDENTITY();";
                return await db.ExecuteScalarAsync<int>(insertSql, new { DT = meetingDate });
            }
        }

        public async Task SaveBODMeetingParticipant(int meetingID, int folioNo, SIALDirectorViewModel director)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
            IF NOT EXISTS (SELECT 1 FROM BOD_Meetings_Participants WHERE BOD_Meetings_RefID=@MID AND SD_RefID=@SDID)
            BEGIN
                INSERT INTO BOD_Meetings_Participants 
                (BOD_Meetings_RefID, Folio_No, SD_RefID, DirectorName, Bank_Account_No, Bank_Account_Title, Bank_Code, Bank_Name)
                VALUES (@MID, @Folio, @SDID, @Name, @AccNo, @AccTitle, @BCode, @BName)
            END";

                await db.ExecuteAsync(sql, new
                {
                    MID = meetingID,
                    Folio = folioNo,
                    SDID = director.EntryID,
                    Name = director.DirectorName,
                    AccNo = director.Bank_Account_No,
                    AccTitle = director.Bank_Account_Title,
                    BCode = director.Bank_Code,
                    BName = director.Bank_Name
                });
            }
        }
        public async Task<SIALDirectorViewModel?> GetDirectorByFolio(int folioNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // We select the columns needed for the BOD_Meetings_Participants table
                string sql = @"SELECT EntryID, DirectorName, Bank_Account_No, 
                              Bank_Account_Title, Bank_Code, Bank_Name 
                       FROM SIALDirectors 
                       WHERE FolioNo = @Folio";

                // QueryFirstOrDefault returns null if no director is found
                return await db.QueryFirstOrDefaultAsync<SIALDirectorViewModel>(sql, new { Folio = folioNo });
            }
        }
    }
}
