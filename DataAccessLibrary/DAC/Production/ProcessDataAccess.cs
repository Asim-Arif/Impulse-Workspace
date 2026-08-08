using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Production
{
    public class ProcessDataAccess : IProcessDataAccess
    {
        private readonly IConfiguration _configuration;

        public ProcessDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? string.Empty;

        public async Task<List<ProcessListItem>> GetProcessesAsync(bool isInspectionProcess = false)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 
                    vp.ProcessID,
                    vp.SNO,
                    ISNULL(vp.Code, '') AS Code,
                    ISNULL(vp.Description, '') AS Description,
                    ISNULL(vp.Supervisor, '') AS Supervisor,
                    ISNULL(vp.Operation, 0) AS Operation,
                    vp.AuthRequired,
                    ISNULL(vp.ProcessNameUrdu, '') AS ProcessNameUrdu,
                    ISNULL(vp.ProcessNameUrduOther, '') AS ProcessNameUrduOther,
                    vp.Insp_RefID,
                    vp.Fix_Maker_RefID,
                    ISNULL(vp.VenderName, '') AS VenderName,
                    ISNULL(vp.Insp_Code, '') AS Insp_Code,
                    ISNULL(vp.Insp_ProcessName, '') AS Insp_ProcessName,
                    ISNULL(vp.Insp_ProcessNameUrdu, '') AS Insp_ProcessNameUrdu,
                    vp.InspectionProcess,
                    vp.BillingProcessID,
                    ISNULL(bp.Description, '') AS BillingProcessName
                FROM VProcesses vp
                LEFT JOIN Processes bp ON vp.BillingProcessID = bp.ProcessID
                WHERE ISNULL(vp.InspectionProcess, 0) = @IsInsp
                ORDER BY vp.Code";

            var list = (await db.QueryAsync<ProcessListItem>(sql, new { IsInsp = isInspectionProcess ? 1 : 0 })).ToList();
            return list;
        }

        public async Task<ProcessModel?> GetProcessByIdAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 
                    ProcessID,
                    SNo,
                    ISNULL(Code, '') AS Code,
                    ISNULL(Description, '') AS Description,
                    ISNULL(Supervisor, '') AS Supervisor,
                    ISNULL(Operation, 0) AS Operation,
                    ISNULL(AuthRequired, 0) AS AuthRequired,
                    ISNULL(ProcessNameUrdu, '') AS ProcessNameUrdu,
                    ISNULL(ProcessNameUrduOther, '') AS ProcessNameUrduOther,
                    Insp_RefID,
                    Fix_Maker_RefID,
                    ISNULL(InspectionProcess, 0) AS InspectionProcess,
                    BillingProcessID
                FROM Processes
                WHERE ProcessID = @ProcessID";

            return await db.QueryFirstOrDefaultAsync<ProcessModel>(sql, new { ProcessID = processId });
        }

        public async Task<ProcessLookups> GetLookupsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string makerSql = "SELECT VendID AS Id, VenderName AS Name FROM Makers WHERE ISNULL(Active, 1) = 1 ORDER BY VenderName";
            const string inspSql = "SELECT EntryID AS Id, '{' + Code + '} ' + ISNULL(ProcessName, '') AS Name FROM InspectionProcesses ORDER BY Code";
            const string billSql = "SELECT ProcessID AS Id, '{' + Code + '} ' + ISNULL(Description, '') AS Name FROM Processes ORDER BY Code";

            var lookups = new ProcessLookups();
            lookups.Makers = (await db.QueryAsync<LookupItemInt>(makerSql)).ToList();
            lookups.InspectionProcesses = (await db.QueryAsync<LookupItemInt>(inspSql)).ToList();
            lookups.BillingProcesses = (await db.QueryAsync<LookupItemInt>(billSql)).ToList();

            return lookups;
        }

        public async Task<int> SaveProcessAsync(ProcessModel model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string maxSql = "SELECT ISNULL(MAX(SNo), 0) + 1 FROM Processes";
            int nextSno = await db.ExecuteScalarAsync<int>(maxSql);

            const string insertSql = @"
                INSERT INTO Processes (
                    SNo, Description, Supervisor, Operation, AuthRequired, Code,
                    ProcessNameUrdu, Insp_RefID, Fix_Maker_RefID, InspectionProcess,
                    ProcessNameUrduOther, BillingProcessID
                ) VALUES (
                    @SNo, @Description, @Supervisor, @Operation, @AuthRequired, @Code,
                    @ProcessNameUrdu, @Insp_RefID, @Fix_Maker_RefID, @InspectionProcess,
                    @ProcessNameUrduOther, @BillingProcessID
                );
                SELECT CAST(SCOPE_IDENTITY() AS INT);";

            return await db.ExecuteScalarAsync<int>(insertSql, new
            {
                SNo = nextSno,
                Description = model.Description ?? string.Empty,
                Supervisor = model.Supervisor ?? string.Empty,
                Operation = model.Operation,
                AuthRequired = model.AuthRequired,
                Code = model.Code ?? string.Empty,
                ProcessNameUrdu = model.ProcessNameUrdu ?? string.Empty,
                Insp_RefID = model.Insp_RefID,
                Fix_Maker_RefID = model.Fix_Maker_RefID,
                InspectionProcess = model.InspectionProcess,
                ProcessNameUrduOther = model.ProcessNameUrduOther ?? string.Empty,
                BillingProcessID = model.BillingProcessID
            });
        }

        public async Task<bool> UpdateProcessAsync(ProcessModel model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string updateSql = @"
                UPDATE Processes SET
                    Description = @Description,
                    Supervisor = @Supervisor,
                    Operation = @Operation,
                    AuthRequired = @AuthRequired,
                    Code = @Code,
                    ProcessNameUrdu = @ProcessNameUrdu,
                    Insp_RefID = @Insp_RefID,
                    Fix_Maker_RefID = @Fix_Maker_RefID,
                    ProcessNameUrduOther = @ProcessNameUrduOther,
                    BillingProcessID = @BillingProcessID
                WHERE ProcessID = @ProcessID";

            int rows = await db.ExecuteAsync(updateSql, new
            {
                ProcessID = model.ProcessID,
                Description = model.Description ?? string.Empty,
                Supervisor = model.Supervisor ?? string.Empty,
                Operation = model.Operation,
                AuthRequired = model.AuthRequired,
                Code = model.Code ?? string.Empty,
                ProcessNameUrdu = model.ProcessNameUrdu ?? string.Empty,
                Insp_RefID = model.Insp_RefID,
                Fix_Maker_RefID = model.Fix_Maker_RefID,
                ProcessNameUrduOther = model.ProcessNameUrduOther ?? string.Empty,
                BillingProcessID = model.BillingProcessID
            });

            return rows > 0;
        }

        public async Task<(bool CanDelete, string Reason)> CanDeleteProcessAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int makerCount = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM MakerProcesses WHERE ProcessID = @ProcessID",
                new { ProcessID = processId });

            if (makerCount > 0)
                return (false, "Cannot delete process: It is currently assigned to one or more Makers/Contractors.");

            int itemCount = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM ItemProcesses WHERE ProcessID = @ProcessID",
                new { ProcessID = processId });

            if (itemCount > 0)
                return (false, "Cannot delete process: It is currently assigned to one or more Items.");

            return (true, string.Empty);
        }

        public async Task<bool> DeleteProcessAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "DELETE FROM Processes WHERE ProcessID = @ProcessID";
            int rows = await db.ExecuteAsync(sql, new { ProcessID = processId });
            return rows > 0;
        }

        public async Task<List<ProcessInspectionPointModel>> GetInspectionPointsAsync(int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntryID, ProcessID, ISNULL(Point_Description, '') AS Point_Description
                FROM Process_Inspection_Points
                WHERE ProcessID = @ProcessID
                ORDER BY EntryID";

            var list = (await db.QueryAsync<ProcessInspectionPointModel>(sql, new { ProcessID = processId })).ToList();
            return list;
        }

        public async Task<bool> AddInspectionPointAsync(ProcessInspectionPointModel point)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                INSERT INTO Process_Inspection_Points (ProcessID, Point_Description)
                VALUES (@ProcessID, @Point_Description)";

            int rows = await db.ExecuteAsync(sql, new
            {
                ProcessID = point.ProcessID,
                Point_Description = point.Point_Description ?? string.Empty
            });
            return rows > 0;
        }

        public async Task<bool> UpdateInspectionPointAsync(ProcessInspectionPointModel point)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                UPDATE Process_Inspection_Points
                SET Point_Description = @Point_Description
                WHERE EntryID = @EntryID";

            int rows = await db.ExecuteAsync(sql, new
            {
                EntryID = point.EntryID,
                Point_Description = point.Point_Description ?? string.Empty
            });
            return rows > 0;
        }

        public async Task<bool> DeleteInspectionPointAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "DELETE FROM Process_Inspection_Points WHERE EntryID = @EntryID";
            int rows = await db.ExecuteAsync(sql, new { EntryID = entryId });
            return rows > 0;
        }
    }
}
