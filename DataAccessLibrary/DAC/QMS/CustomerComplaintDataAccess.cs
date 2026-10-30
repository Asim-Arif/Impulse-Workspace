using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.QMS
{
    public class CustomerComplaintDataAccess : ICustomerComplaintDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<CustomerComplaintDataAccess> _logger;

        public CustomerComplaintDataAccess(IConfiguration configuration, ILogger<CustomerComplaintDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<CustomerComplaintListItemModel>> GetComplaintsAsync(CustomerComplaintFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();

                var sql = new StringBuilder(@"
                    SELECT 
                        c.EntryID,
                        ISNULL(c.CustCode, '') AS CustCode,
                        ISNULL(c.Country, '') AS Country,
                        ISNULL(c.ComplaintNo, '') AS ComplaintNo,
                        c.DT,
                        c.Type,
                        c.NCNo,
                        c.CAPANo,
                        c.CreditNoteNo,
                        CASE WHEN c.AttachmentPDF IS NOT NULL THEN 1 ELSE 0 END AS PDFAttached,
                        ISNULL((
                            SELECT COUNT(1) 
                            FROM EntityAttachments ea 
                            WHERE ea.EntityType = 'CustomerComplaint' 
                              AND ea.EntityId = CAST(c.EntryID AS VARCHAR(50)) 
                              AND ea.IsDeleted = 0
                        ), 0) AS AttachmentCount
                    FROM CustomerComplaints c
                    WHERE 1 = 1");

                var parameters = new DynamicParameters();

                if (!string.IsNullOrWhiteSpace(filter.ComplaintNo))
                {
                    sql.Append(" AND c.ComplaintNo LIKE @ComplaintNo");
                    parameters.Add("ComplaintNo", $"%{filter.ComplaintNo.Trim()}%");
                }
                else
                {
                    if (filter.DateFrom.HasValue && filter.DateTo.HasValue)
                    {
                        sql.Append(" AND c.DT >= @DateFrom AND c.DT <= @DateTo");
                        parameters.Add("DateFrom", filter.DateFrom.Value.Date);
                        parameters.Add("DateTo", filter.DateTo.Value.Date.AddDays(1).AddTicks(-1));
                    }

                    if (!string.IsNullOrWhiteSpace(filter.CustCode) && filter.CustCode != "0")
                    {
                        sql.Append(" AND c.CustCode = @CustCode");
                        parameters.Add("CustCode", filter.CustCode.Trim());
                    }

                    if (!string.IsNullOrWhiteSpace(filter.Country) && filter.Country != "<All Countries>")
                    {
                        sql.Append(" AND c.Country = @Country");
                        parameters.Add("Country", filter.Country.Trim());
                    }

                    var includedTypes = new List<byte>();
                    if (filter.IncludeComplaints) includedTypes.Add(0);
                    if (filter.IncludeFeedback) includedTypes.Add(1);
                    if (filter.IncludeAudits) includedTypes.Add(2);

                    if (includedTypes.Count > 0 && includedTypes.Count < 3)
                    {
                        sql.Append(" AND c.Type IN @IncludedTypes");
                        parameters.Add("IncludedTypes", includedTypes);
                    }
                    else if (includedTypes.Count == 0)
                    {
                        sql.Append(" AND 1 = 0");
                    }
                }

                sql.Append(" ORDER BY c.DT DESC, c.EntryID DESC");

                var results = await db.QueryAsync<CustomerComplaintListItemModel>(sql.ToString(), parameters);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching customer complaints list");
                throw;
            }
        }

        public async Task<List<string>> GetCustomerCodesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT DISTINCT CustCode FROM ForeignCustomers WHERE CustCode IS NOT NULL AND CustCode <> '' ORDER BY CustCode";
                var list = await db.QueryAsync<string>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching distinct customer codes");
                return new List<string>();
            }
        }

        public async Task<List<string>> GetCountriesAsync(string? custCode = null)
        {
            try
            {
                using var db = CreateConnection();
                string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE Country IS NOT NULL AND Country <> ''";
                var parameters = new DynamicParameters();

                if (!string.IsNullOrWhiteSpace(custCode) && custCode != "0")
                {
                    sql += " AND CustCode = @CustCode";
                    parameters.Add("CustCode", custCode.Trim());
                }

                sql += " ORDER BY Country";

                var list = await db.QueryAsync<string>(sql, parameters);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching distinct countries");
                return new List<string>();
            }
        }

        public async Task<string?> GetNCNoAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT NCNo FROM CustomerComplaints WHERE EntryID = @EntryID";
                return await db.QueryFirstOrDefaultAsync<string>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching NCNo for EntryID {EntryID}", entryId);
                return null;
            }
        }

        public async Task<string?> GetCAPANoAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT CAPANo FROM CustomerComplaints WHERE EntryID = @EntryID";
                return await db.QueryFirstOrDefaultAsync<string>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CAPANo for EntryID {EntryID}", entryId);
                return null;
            }
        }

        public async Task<int> GetNCEntryIdAsync(string ncNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT ISNULL(EntryID, 0) FROM NonCompliance WHERE NCNo = @NCNo";
                return await db.QueryFirstOrDefaultAsync<int>(sql, new { NCNo = ncNo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching EntryID from NonCompliance for NCNo {NCNo}", ncNo);
                return 0;
            }
        }

        public async Task<int> GetCAPAEntryIdAsync(string capaNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT ISNULL(EntryID, 0) FROM CAPA WHERE CAPANo = @CAPANo";
                return await db.QueryFirstOrDefaultAsync<int>(sql, new { CAPANo = capaNo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching EntryID from CAPA for CAPANo {CAPANo}", capaNo);
                return 0;
            }
        }

        public async Task<CustomerComplaintModel?> GetComplaintByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        EntryID, Company_RefID, CustCode, Country, ComplaintNo, DT,
                        Unit, ComplaintBy, Designation, Penalty, PenaltyQty, PenaltyAmount, PenaltyCurrency,
                        EmailTo, EmailCc, Type, UserName, MachineName, DTEntry,
                        Description, Reason, ActionRequested, InternalAudit, DTNextAudit, Findings,
                        NCNo, CAPANo, CreditNoteNo
                    FROM CustomerComplaints
                    WHERE EntryID = @EntryID";

                var model = await db.QueryFirstOrDefaultAsync<CustomerComplaintModel>(sql, new { EntryID = entryId });
                if (model != null)
                {
                    model.Details = await GetComplaintDetailsAsync(entryId);
                }
                return model;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CustomerComplaint by ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<List<CustomerComplaintDetailModel>> GetComplaintDetailsAsync(int refId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        d.EntryID, d.RefID, d.InvoiceNo, d.ItemID, d.Qty,
                        d.EmailReference, d.ItemReason_Main, d.ItemReason, d.Price,
                        T1.ItemCode
                    FROM CustomerComplaints_Detail d
                    LEFT JOIN (
                        SELECT CustomInvoice, ItemCode, CompItemCode
                        FROM VCustomInvoiceItems
                        GROUP BY CustomInvoice, ItemCode, CompItemCode
                    ) T1 ON d.InvoiceNo = T1.CustomInvoice AND d.ItemID = T1.CompItemCode
                    WHERE d.RefID = @RefID
                    ORDER BY d.EntryID";

                var details = await db.QueryAsync<CustomerComplaintDetailModel>(sql, new { RefID = refId });
                return details.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CustomerComplaint details for RefID {RefID}", refId);
                throw;
            }
        }

        public async Task<int> SaveComplaintAsync(CustomerComplaintModel model, List<CustomerComplaintDetailModel> details, string currentUserName, string? machineName)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                int complaintId = model.EntryID;

                if (complaintId == 0)
                {
                    const string insertSql = @"
                        INSERT INTO CustomerComplaints (
                            Company_RefID, CustCode, Country, ComplaintNo, DT,
                            Unit, ComplaintBy, Designation, Penalty, PenaltyQty, PenaltyAmount, PenaltyCurrency,
                            EmailTo, EmailCc, Type, UserName, MachineName, DTEntry,
                            Description, Reason, ActionRequested, InternalAudit, DTNextAudit, Findings
                        ) VALUES (
                            @Company_RefID, @CustCode, @Country, @ComplaintNo, @DT,
                            @Unit, @ComplaintBy, @Designation, @Penalty, @PenaltyQty, @PenaltyAmount, @PenaltyCurrency,
                            @EmailTo, @EmailCc, @Type, @UserName, @MachineName, GETDATE(),
                            @Description, @Reason, @ActionRequested, @InternalAudit, @DTNextAudit, @Findings
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    complaintId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.Company_RefID,
                        model.CustCode,
                        model.Country,
                        model.ComplaintNo,
                        model.DT,
                        model.Unit,
                        model.ComplaintBy,
                        model.Designation,
                        model.Penalty,
                        model.PenaltyQty,
                        model.PenaltyAmount,
                        model.PenaltyCurrency,
                        model.EmailTo,
                        model.EmailCc,
                        Type = model.Type ?? 0,
                        UserName = currentUserName,
                        MachineName = machineName ?? Environment.MachineName,
                        model.Description,
                        model.Reason,
                        model.ActionRequested,
                        model.InternalAudit,
                        model.DTNextAudit,
                        model.Findings
                    }, tx);
                }
                else
                {
                    const string updateSql = @"
                        UPDATE CustomerComplaints SET
                            Company_RefID = @Company_RefID,
                            CustCode = @CustCode,
                            Country = @Country,
                            ComplaintNo = @ComplaintNo,
                            DT = @DT,
                            Unit = @Unit,
                            ComplaintBy = @ComplaintBy,
                            Designation = @Designation,
                            Penalty = @Penalty,
                            PenaltyQty = @PenaltyQty,
                            PenaltyAmount = @PenaltyAmount,
                            PenaltyCurrency = @PenaltyCurrency,
                            EmailTo = @EmailTo,
                            EmailCc = @EmailCc,
                            Type = @Type,
                            Description = @Description,
                            Reason = @Reason,
                            ActionRequested = @ActionRequested,
                            InternalAudit = @InternalAudit,
                            DTNextAudit = @DTNextAudit,
                            Findings = @Findings
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.Company_RefID,
                        model.CustCode,
                        model.Country,
                        model.ComplaintNo,
                        model.DT,
                        model.Unit,
                        model.ComplaintBy,
                        model.Designation,
                        model.Penalty,
                        model.PenaltyQty,
                        model.PenaltyAmount,
                        model.PenaltyCurrency,
                        model.EmailTo,
                        model.EmailCc,
                        Type = model.Type ?? 0,
                        model.Description,
                        model.Reason,
                        model.ActionRequested,
                        model.InternalAudit,
                        model.DTNextAudit,
                        model.Findings,
                        model.EntryID
                    }, tx);
                }

                // Delete removed detail rows
                var deletedIds = details.Where(d => d.IsDeleted && d.EntryID > 0).Select(d => d.EntryID).ToList();
                if (deletedIds.Count > 0)
                {
                    const string deleteSql = "DELETE FROM CustomerComplaints_Detail WHERE EntryID IN @DeletedIds";
                    await db.ExecuteAsync(deleteSql, new { DeletedIds = deletedIds }, tx);
                }

                // Upsert active detail rows
                foreach (var detail in details.Where(d => !d.IsDeleted))
                {
                    if (detail.EntryID == 0)
                    {
                        const string insertDetailSql = @"
                            INSERT INTO CustomerComplaints_Detail (
                                RefID, InvoiceNo, ItemID, Qty, EmailReference, ItemReason_Main, ItemReason, Price
                            ) VALUES (
                                @RefID, @InvoiceNo, @ItemID, @Qty, @EmailReference, @ItemReason_Main, @ItemReason, @Price
                            );";

                        await db.ExecuteAsync(insertDetailSql, new
                        {
                            RefID = complaintId,
                            detail.InvoiceNo,
                            detail.ItemID,
                            detail.Qty,
                            detail.EmailReference,
                            detail.ItemReason_Main,
                            detail.ItemReason,
                            detail.Price
                        }, tx);
                    }
                    else
                    {
                        const string updateDetailSql = @"
                            UPDATE CustomerComplaints_Detail SET
                                InvoiceNo = @InvoiceNo,
                                ItemID = @ItemID,
                                Qty = @Qty,
                                EmailReference = @EmailReference,
                                ItemReason_Main = @ItemReason_Main,
                                ItemReason = @ItemReason,
                                Price = @Price
                            WHERE EntryID = @EntryID;";

                        await db.ExecuteAsync(updateDetailSql, new
                        {
                            detail.InvoiceNo,
                            detail.ItemID,
                            detail.Qty,
                            detail.EmailReference,
                            detail.ItemReason_Main,
                            detail.ItemReason,
                            detail.Price,
                            detail.EntryID
                        }, tx);
                    }
                }

                tx.Commit();
                return complaintId;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving CustomerComplaint EntryID {EntryID}", model.EntryID);
                throw;
            }
        }

        public async Task<string> GetNextComplaintNoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(ComplaintNo, 4, CHARINDEX('/', ComplaintNo) - 4) AS INT)), 0)
                    FROM CustomerComplaints
                    WHERE Type = 0 AND LEFT(ComplaintNo, 3) = 'CC-' AND CHARINDEX('/', ComplaintNo) > 4";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"CC-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next complaint number");
                return $"CC-1/{date.Year}";
            }
        }

        public async Task<string> GetNextFeedbackNoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(ComplaintNo, 4, CHARINDEX('/', ComplaintNo) - 4) AS INT)), 0)
                    FROM CustomerComplaints
                    WHERE Type = 1 AND LEFT(ComplaintNo, 3) = 'FB-' AND CHARINDEX('/', ComplaintNo) > 4";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"FB-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next feedback number");
                return $"FB-1/{date.Year}";
            }
        }

        public async Task<string> GetNextAuditNoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(ComplaintNo, 4, CHARINDEX('/', ComplaintNo) - 4) AS INT)), 0)
                    FROM CustomerComplaints
                    WHERE Type = 2 AND LEFT(ComplaintNo, 3) = 'AU-' AND CHARINDEX('/', ComplaintNo) > 4";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"AU-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next audit number");
                return $"AU-1/{date.Year}";
            }
        }

        public async Task<NonComplianceModel?> GetNCByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        nc.EntryID, nc.NCNo, nc.NCType, nc.DT, nc.Source, nc.Department, nc.Description,
                        cc.EntryID as ComplaintEntryID, cc.CustCode, cc.Country, cc.ComplaintNo, cc.DT as ComplaintDate,
                        cc.ComplaintBy, cc.Designation, cc.Penalty, cc.PenaltyQty, cc.PenaltyAmount, cc.PenaltyCurrency
                    FROM NonCompliance nc
                    LEFT JOIN CustomerComplaints cc ON nc.NCNo = cc.NCNo
                    WHERE nc.EntryID = @EntryID";

                return await db.QueryFirstOrDefaultAsync<NonComplianceModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching NonCompliance by ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<NonComplianceModel?> GetNCByComplaintIdAsync(int complaintId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        nc.EntryID, nc.NCNo, nc.NCType, nc.DT, nc.Source, nc.Department, nc.Description,
                        cc.EntryID as ComplaintEntryID, cc.CustCode, cc.Country, cc.ComplaintNo, cc.DT as ComplaintDate,
                        cc.ComplaintBy, cc.Designation, cc.Penalty, cc.PenaltyQty, cc.PenaltyAmount, cc.PenaltyCurrency
                    FROM CustomerComplaints cc
                    INNER JOIN NonCompliance nc ON cc.NCNo = nc.NCNo
                    WHERE cc.EntryID = @ComplaintId";

                return await db.QueryFirstOrDefaultAsync<NonComplianceModel>(sql, new { ComplaintId = complaintId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching NonCompliance by Complaint ID {ComplaintId}", complaintId);
                throw;
            }
        }

        public async Task<string> GetNextNCNoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(NCNo, 4, CHARINDEX('/', NCNo) - 4) AS INT)), 0)
                    FROM NonCompliance
                    WHERE LEFT(NCNo, 3) = 'NC-' AND CHARINDEX('/', NCNo) > 4";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"NC-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next NC number");
                return $"NC-1/{date.Year}";
            }
        }

        public async Task<int> SaveNCAsync(NonComplianceModel model, int? complaintId, string currentUserName)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                int ncId = model.EntryID;

                if (ncId == 0)
                {
                    const string insertSql = @"
                        INSERT INTO NonCompliance (
                            NCNo, NCType, DT, Source, Department, Description
                        ) VALUES (
                            @NCNo, @NCType, @DT, @Source, @Department, @Description
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    ncId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.NCNo,
                        model.NCType,
                        model.DT,
                        model.Source,
                        model.Department,
                        model.Description
                    }, tx);
                }
                else
                {
                    const string updateSql = @"
                        UPDATE NonCompliance SET
                            NCNo = @NCNo,
                            NCType = @NCType,
                            DT = @DT,
                            Source = @Source,
                            Department = @Department,
                            Description = @Description
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.NCNo,
                        model.NCType,
                        model.DT,
                        model.Source,
                        model.Department,
                        model.Description,
                        model.EntryID
                    }, tx);
                }

                // Link with Customer Complaint if complaintId provided
                if (complaintId.HasValue && complaintId.Value > 0)
                {
                    const string updateComplaintSql = "UPDATE CustomerComplaints SET NCNo = @NCNo WHERE EntryID = @ComplaintId;";
                    await db.ExecuteAsync(updateComplaintSql, new { model.NCNo, ComplaintId = complaintId.Value }, tx);
                }

                tx.Commit();
                return ncId;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving NonCompliance EntryID {EntryID}", model.EntryID);
                throw;
            }
        }

        public async Task<CAPAModel?> GetCAPAByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        c.EntryID, c.CAPANo, c.DT, c.Reason, c.System, c.Description,
                        c.Acceptance, c.TeamNominated, c.ResultOfInvestigation, c.CorrectiveAction,
                        c.CAReview, c.Remarks, c.PreventiveActions, c.ResponseEvaluation, c.FollowUp,
                        c.System_QM, c.System_EM, c.System_cGMP, c.System_OHSM, c.System_ITSM, c.System_MDD,
                        c.ResultDT, c.CorrectiveActionDT, c.PreventiveActionDT,
                        c.CorrectiveActionFollowupDT, c.PreventiveActionFollowUpDT, c.CAPACloseDT,
                        cc.EntryID as ComplaintEntryID, cc.CustCode, cc.Country, cc.ComplaintNo, cc.DT as ComplaintDate,
                        cc.ComplaintBy, cc.Designation, cc.Penalty, cc.PenaltyQty, cc.PenaltyAmount, cc.PenaltyCurrency, cc.Reason as ComplaintReason,
                        nc.EntryID as NCEntryID, nc.NCNo, nc.NCType, nc.DT as NCDate, nc.Source as NCSource, nc.Department as NCDepartment, nc.Description as NCDescription
                    FROM CAPA c
                    LEFT JOIN CustomerComplaints cc ON c.CAPANo = cc.CAPANo
                    LEFT JOIN NonCompliance nc ON cc.NCNo = nc.NCNo
                    WHERE c.EntryID = @EntryID;";

                return await db.QueryFirstOrDefaultAsync<CAPAModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CAPA by ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<CAPAModel?> GetCAPAByComplaintIdAsync(int complaintId)
        {
            try
            {
                using var db = CreateConnection();
                // Check if complaint already has a linked CAPA
                const string sqlWithCapa = @"
                    SELECT 
                        c.EntryID, c.CAPANo, c.DT, c.Reason, c.System, c.Description,
                        c.Acceptance, c.TeamNominated, c.ResultOfInvestigation, c.CorrectiveAction,
                        c.CAReview, c.Remarks, c.PreventiveActions, c.ResponseEvaluation, c.FollowUp,
                        c.System_QM, c.System_EM, c.System_cGMP, c.System_OHSM, c.System_ITSM, c.System_MDD,
                        c.ResultDT, c.CorrectiveActionDT, c.PreventiveActionDT,
                        c.CorrectiveActionFollowupDT, c.PreventiveActionFollowUpDT, c.CAPACloseDT,
                        cc.EntryID as ComplaintEntryID, cc.CustCode, cc.Country, cc.ComplaintNo, cc.DT as ComplaintDate,
                        cc.ComplaintBy, cc.Designation, cc.Penalty, cc.PenaltyQty, cc.PenaltyAmount, cc.PenaltyCurrency, cc.Reason as ComplaintReason,
                        nc.EntryID as NCEntryID, nc.NCNo, nc.NCType, nc.DT as NCDate, nc.Source as NCSource, nc.Department as NCDepartment, nc.Description as NCDescription
                    FROM CustomerComplaints cc
                    INNER JOIN CAPA c ON cc.CAPANo = c.CAPANo
                    LEFT JOIN NonCompliance nc ON cc.NCNo = nc.NCNo
                    WHERE cc.EntryID = @ComplaintId;";

                var capa = await db.QueryFirstOrDefaultAsync<CAPAModel>(sqlWithCapa, new { ComplaintId = complaintId });
                if (capa != null) return capa;

                // Otherwise load Complaint and NC information for creating a new CAPA
                const string sqlNew = @"
                    SELECT 
                        cc.EntryID as ComplaintEntryID, cc.CustCode, cc.Country, cc.ComplaintNo, cc.DT as ComplaintDate,
                        cc.ComplaintBy, cc.Designation, cc.Penalty, cc.PenaltyQty, cc.PenaltyAmount, cc.PenaltyCurrency, cc.Reason as ComplaintReason,
                        nc.EntryID as NCEntryID, nc.NCNo, nc.NCType, nc.DT as NCDate, nc.Source as NCSource, nc.Department as NCDepartment, nc.Description as NCDescription
                    FROM CustomerComplaints cc
                    LEFT JOIN NonCompliance nc ON cc.NCNo = nc.NCNo
                    WHERE cc.EntryID = @ComplaintId;";

                var newCapa = await db.QueryFirstOrDefaultAsync<CAPAModel>(sqlNew, new { ComplaintId = complaintId });
                if (newCapa != null)
                {
                    newCapa.Description = newCapa.ComplaintReason;
                }
                return newCapa;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CAPA by Complaint ID {ComplaintId}", complaintId);
                throw;
            }
        }

        public async Task<string> GetNextCAPANoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(CAPANo, 6, CHARINDEX('/', CAPANo) - 6) AS INT)), 0)
                    FROM CAPA
                    WHERE LEFT(CAPANo, 5) = 'CAPA-' AND CHARINDEX('/', CAPANo) > 6";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"CAPA-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next CAPA number");
                return $"CAPA-1/{date.Year}";
            }
        }

        public async Task<int> SaveCAPAAsync(CAPAModel model, int? complaintId, string currentUserName)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                int capaId = model.EntryID;

                if (capaId == 0)
                {
                    const string insertSql = @"
                        INSERT INTO CAPA (
                            CAPANo, DT, Reason, System, Description, Acceptance, TeamNominated,
                            ResultOfInvestigation, CorrectiveAction, CAReview, Remarks, PreventiveActions,
                            ResponseEvaluation, FollowUp, System_QM, System_EM, System_cGMP,
                            System_OHSM, System_ITSM, System_MDD, ResultDT, CorrectiveActionDT,
                            PreventiveActionDT, CorrectiveActionFollowupDT, PreventiveActionFollowUpDT, CAPACloseDT
                        ) VALUES (
                            @CAPANo, @DT, @Reason, @System, @Description, @Acceptance, @TeamNominated,
                            @ResultOfInvestigation, @CorrectiveAction, @CAReview, @Remarks, @PreventiveActions,
                            @ResponseEvaluation, @FollowUp, @System_QM, @System_EM, @System_cGMP,
                            @System_OHSM, @System_ITSM, @System_MDD, @ResultDT, @CorrectiveActionDT,
                            @PreventiveActionDT, @CorrectiveActionFollowupDT, @PreventiveActionFollowUpDT, @CAPACloseDT
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    capaId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.CAPANo,
                        model.DT,
                        model.Reason,
                        model.System,
                        model.Description,
                        model.Acceptance,
                        model.TeamNominated,
                        model.ResultOfInvestigation,
                        model.CorrectiveAction,
                        model.CAReview,
                        model.Remarks,
                        model.PreventiveActions,
                        model.ResponseEvaluation,
                        model.FollowUp,
                        model.System_QM,
                        model.System_EM,
                        model.System_cGMP,
                        model.System_OHSM,
                        model.System_ITSM,
                        model.System_MDD,
                        model.ResultDT,
                        model.CorrectiveActionDT,
                        model.PreventiveActionDT,
                        model.CorrectiveActionFollowupDT,
                        model.PreventiveActionFollowUpDT,
                        model.CAPACloseDT
                    }, tx);
                }
                else
                {
                    const string updateSql = @"
                        UPDATE CAPA SET
                            CAPANo = @CAPANo,
                            DT = @DT,
                            Reason = @Reason,
                            System = @System,
                            Description = @Description,
                            Acceptance = @Acceptance,
                            TeamNominated = @TeamNominated,
                            ResultOfInvestigation = @ResultOfInvestigation,
                            CorrectiveAction = @CorrectiveAction,
                            CAReview = @CAReview,
                            Remarks = @Remarks,
                            PreventiveActions = @PreventiveActions,
                            ResponseEvaluation = @ResponseEvaluation,
                            FollowUp = @FollowUp,
                            System_QM = @System_QM,
                            System_EM = @System_EM,
                            System_cGMP = @System_cGMP,
                            System_OHSM = @System_OHSM,
                            System_ITSM = @System_ITSM,
                            System_MDD = @System_MDD,
                            ResultDT = @ResultDT,
                            CorrectiveActionDT = @CorrectiveActionDT,
                            PreventiveActionDT = @PreventiveActionDT,
                            CorrectiveActionFollowupDT = @CorrectiveActionFollowupDT,
                            PreventiveActionFollowUpDT = @PreventiveActionFollowUpDT,
                            CAPACloseDT = @CAPACloseDT
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.CAPANo,
                        model.DT,
                        model.Reason,
                        model.System,
                        model.Description,
                        model.Acceptance,
                        model.TeamNominated,
                        model.ResultOfInvestigation,
                        model.CorrectiveAction,
                        model.CAReview,
                        model.Remarks,
                        model.PreventiveActions,
                        model.ResponseEvaluation,
                        model.FollowUp,
                        model.System_QM,
                        model.System_EM,
                        model.System_cGMP,
                        model.System_OHSM,
                        model.System_ITSM,
                        model.System_MDD,
                        model.ResultDT,
                        model.CorrectiveActionDT,
                        model.PreventiveActionDT,
                        model.CorrectiveActionFollowupDT,
                        model.PreventiveActionFollowUpDT,
                        model.CAPACloseDT,
                        model.EntryID
                    }, tx);
                }

                // Link with Customer Complaint if complaintId provided
                if (complaintId.HasValue && complaintId.Value > 0)
                {
                    const string updateComplaintSql = "UPDATE CustomerComplaints SET CAPANo = @CAPANo WHERE EntryID = @ComplaintId;";
                    await db.ExecuteAsync(updateComplaintSql, new { model.CAPANo, ComplaintId = complaintId.Value }, tx);
                }

                tx.Commit();
                return capaId;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving CAPA EntryID {EntryID}", model.EntryID);
                throw;
            }
        }

        public async Task<string> GetNextCreditNoteNoAsync(DateTime date)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT ISNULL(MAX(CAST(SUBSTRING(CreditNoteNo, 4, CHARINDEX('/', CreditNoteNo) - 4) AS INT)), 0)
                    FROM CustomerComplaints
                    WHERE LEFT(CreditNoteNo, 3) = 'CR-' AND CHARINDEX('/', CreditNoteNo) > 4";

                int maxNum = await db.ExecuteScalarAsync<int>(sql);
                int nextNum = maxNum + 1;
                return $"CR-{nextNum}/{date.Year}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next Credit Note number");
                return $"CR-1/{date.Year}";
            }
        }

        public async Task<bool> SaveCreditNoteAsync(int complaintId, string creditNoteNo, List<CustomerComplaintDetailModel> details, string currentUserName)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                const string updateComplaintSql = @"
                    UPDATE CustomerComplaints 
                    SET CreditNoteNo = @CreditNoteNo 
                    WHERE EntryID = @ComplaintId;";

                await db.ExecuteAsync(updateComplaintSql, new { CreditNoteNo = creditNoteNo, ComplaintId = complaintId }, tx);

                const string updatePriceSql = @"
                    UPDATE CustomerComplaints_Detail 
                    SET Price = @Price 
                    WHERE EntryID = @EntryID;";

                foreach (var item in details.Where(d => !d.IsDeleted && d.EntryID > 0))
                {
                    await db.ExecuteAsync(updatePriceSql, new { Price = item.Price, EntryID = item.EntryID }, tx);
                }

                tx.Commit();
                return true;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving Credit Note for Complaint ID {ComplaintId}", complaintId);
                throw;
            }
        }

        public async Task<List<CompanyLookupDto>> GetCompaniesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT EntryID, CompanyName FROM Companies ORDER BY CompanyName";
                var list = await db.QueryAsync<CompanyLookupDto>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Companies");
                return new List<CompanyLookupDto>();
            }
        }

        public async Task<List<DepartmentLookupDto>> GetDepartmentsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT DeptID, Name FROM Departments WHERE Active = 1 ORDER BY Name";
                var list = await db.QueryAsync<DepartmentLookupDto>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Departments");
                return new List<DepartmentLookupDto>();
            }
        }

        public async Task<List<QMSReasonLookupDto>> GetQMSReasonsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT EntryID, Reason_QMS FROM QMS_Reasons_List ORDER BY EntryID";
                var list = await db.QueryAsync<QMSReasonLookupDto>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS reasons");
                return new List<QMSReasonLookupDto>();
            }
        }

        public async Task<List<EmployeeLookupDto>> GetEmployeesByDeptAsync(string deptId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT EmpID, Name, DeptID 
                    FROM Employees 
                    WHERE DeptID = @DeptID AND Active = 1 
                    ORDER BY Name";

                var list = await db.QueryAsync<EmployeeLookupDto>(sql, new { DeptID = deptId });
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Employees for DeptID {DeptID}", deptId);
                return new List<EmployeeLookupDto>();
            }
        }

        public async Task<List<string>> GetInvoicesByCustomerAsync(string custCode, string? country)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT DISTINCT CustomInvoice FROM CustomInvoice WHERE CustCode = @CustCode";
                var parameters = new DynamicParameters();
                parameters.Add("CustCode", custCode);

                if (!string.IsNullOrWhiteSpace(country))
                {
                    sql += " AND Country = @Country";
                    parameters.Add("Country", country);
                }

                sql += " ORDER BY CustomInvoice DESC";

                var list = await db.QueryAsync<string>(sql, parameters);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Invoices for customer {CustCode}", custCode);
                return new List<string>();
            }
        }

        public async Task<List<InvoiceItemLookupDto>> GetInvoiceItemsAsync(string custCode, string? country, string? invoiceNo)
        {
            try
            {
                using var db = CreateConnection();

                if (!string.IsNullOrWhiteSpace(invoiceNo) && invoiceNo != "<All Items>")
                {
                    const string sql = @"
                        SELECT DISTINCT ItemID, CustItemID AS CompItemID
                        FROM VCustomInvoiceItems
                        WHERE CustomInvoice = @InvoiceNo
                        ORDER BY ItemID";

                    var list = await db.QueryAsync<InvoiceItemLookupDto>(sql, new { InvoiceNo = invoiceNo });
                    return list.ToList();
                }
                else
                {
                    const string sql = @"
                        SELECT DISTINCT CompItemID, ItemID
                        FROM VFCustomerCatalog
                        WHERE CustCode = @CustCode AND (Country = @Country OR @Country IS NULL)
                        ORDER BY ItemID";

                    var list = await db.QueryAsync<InvoiceItemLookupDto>(sql, new { CustCode = custCode, Country = country });
                    return list.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching InvoiceItems for invoice {InvoiceNo}", invoiceNo);
                return new List<InvoiceItemLookupDto>();
            }
        }

        public async Task<List<string>> GetLotsAsync(string invoiceNo, string itemId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT LotNo 
                    FROM VCustomPList 
                    WHERE CustomInvoice = @InvoiceNo AND ItemID = @ItemID AND LotNo IS NOT NULL AND LotNo <> ''
                    ORDER BY LotNo";

                var list = await db.QueryAsync<string>(sql, new { InvoiceNo = invoiceNo, ItemID = itemId });
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching lots for Invoice {InvoiceNo}, Item {ItemID}", invoiceNo, itemId);
                return new List<string>();
            }
        }
    }
}
