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
    public class QMSDocumentDataAccess : IQMSDocumentDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<QMSDocumentDataAccess> _logger;

        public QMSDocumentDataAccess(IConfiguration configuration, ILogger<QMSDocumentDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<QMSDocumentListItemModel>> GetDocumentsAsync(QMSDocumentFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();
                var sql = new StringBuilder(@"
                    SELECT 
                        d.EntryID,
                        d.DocumentNumber,
                        d.DocumentName,
                        d.Revision,
                        d.DocumentType,
                        d.IssueDT,
                        d.Internal,
                        d.InitiaterEmpID,
                        e.name AS InitiatorName,
                        e.DeptName,
                        (SELECT COUNT(1) FROM QMS_Documents_DistributionList dl WHERE dl.QMS_RefID = d.EntryID) AS DistributionCount,
                        (SELECT COUNT(1) FROM EntityAttachments ea WHERE ea.EntityType = 'QMSDocument' AND ea.EntityId = CAST(d.EntryID AS VARCHAR(50))) AS AttachmentCount
                    FROM QMS_Documents d
                    LEFT JOIN VEmp e ON d.InitiaterEmpID = e.empid
                    WHERE 1 = 1");

                var parameters = new DynamicParameters();

                if (filter != null)
                {
                    if (filter.DateFrom.HasValue && filter.DateTo.HasValue)
                    {
                        sql.Append(" AND d.IssueDT >= @DateFrom AND d.IssueDT <= @DateTo");
                        parameters.Add("DateFrom", filter.DateFrom.Value.Date);
                        parameters.Add("DateTo", filter.DateTo.Value.Date.AddDays(1).AddTicks(-1));
                    }

                    if (!string.IsNullOrWhiteSpace(filter.DocumentType))
                    {
                        sql.Append(" AND d.DocumentType = @DocumentType");
                        parameters.Add("DocumentType", filter.DocumentType.Trim());
                    }

                    if (filter.InternalFilter.HasValue && filter.InternalFilter.Value >= 0)
                    {
                        sql.Append(" AND d.Internal = @Internal");
                        parameters.Add("Internal", filter.InternalFilter.Value == 1);
                    }

                    if (!string.IsNullOrWhiteSpace(filter.SearchText))
                    {
                        sql.Append(@" AND (
                            d.DocumentNumber LIKE @Search 
                            OR d.DocumentName LIKE @Search 
                            OR d.DocumentType LIKE @Search
                            OR d.Revision LIKE @Search
                            OR e.name LIKE @Search
                        )");
                        parameters.Add("Search", $"%{filter.SearchText.Trim()}%");
                    }
                }

                sql.Append(" ORDER BY d.IssueDT DESC, d.EntryID DESC");

                var results = await db.QueryAsync<QMSDocumentListItemModel>(sql.ToString(), parameters);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS documents list.");
                throw;
            }
        }

        public async Task<QMSDocumentModel?> GetDocumentByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string docSql = @"
                    SELECT 
                        d.EntryID,
                        d.DocumentNumber,
                        d.DocumentName,
                        d.Revision,
                        d.DocumentType,
                        d.IssueDT,
                        d.InitiaterEmpID,
                        e.name AS InitiatorName,
                        e.DeptName AS InitiatorDepartment,
                        e.Designation AS InitiatorDesignation,
                        ISNULL(d.Internal, 1) AS Internal
                    FROM QMS_Documents d
                    LEFT JOIN VEmp e ON d.InitiaterEmpID = e.empid
                    WHERE d.EntryID = @EntryID";

                var doc = await db.QueryFirstOrDefaultAsync<QMSDocumentModel>(docSql, new { EntryID = entryId });
                if (doc == null) return null;

                const string distSql = @"
                    SELECT 
                        dl.EntryID,
                        dl.QMS_RefID,
                        dl.EmpID,
                        dl.name AS EmpName,
                        dl.DeptName,
                        dl.Designation
                    FROM VQMS_Documents_DistributionList dl
                    WHERE dl.QMS_RefID = @EntryID
                    ORDER BY dl.name";

                var distList = await db.QueryAsync<QMSDocumentDistributionModel>(distSql, new { EntryID = entryId });
                doc.DistributionList = distList.ToList();

                return doc;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS document by ID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<int> SaveDocumentAsync(QMSDocumentModel model, List<int> deletedDistributionIds, string currentUserName)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                int docId = model.EntryID;

                if (docId == 0)
                {
                    const string insertSql = @"
                        INSERT INTO QMS_Documents (
                            DocumentName,
                            Revision,
                            DocumentNumber,
                            IssueDT,
                            InitiaterEmpID,
                            DocumentType,
                            Internal
                        ) VALUES (
                            @DocumentName,
                            @Revision,
                            @DocumentNumber,
                            @IssueDT,
                            @InitiaterEmpID,
                            @DocumentType,
                            @Internal
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    docId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.DocumentName,
                        model.Revision,
                        model.DocumentNumber,
                        model.IssueDT,
                        model.InitiaterEmpID,
                        model.DocumentType,
                        model.Internal
                    }, trans);

                    model.EntryID = docId;
                }
                else
                {
                    const string updateSql = @"
                        UPDATE QMS_Documents
                        SET 
                            DocumentName = @DocumentName,
                            Revision = @Revision,
                            DocumentNumber = @DocumentNumber,
                            IssueDT = @IssueDT,
                            InitiaterEmpID = @InitiaterEmpID,
                            DocumentType = @DocumentType,
                            Internal = @Internal
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.EntryID,
                        model.DocumentName,
                        model.Revision,
                        model.DocumentNumber,
                        model.IssueDT,
                        model.InitiaterEmpID,
                        model.DocumentType,
                        model.Internal
                    }, trans);
                }

                // Delete removed distribution list entries
                if (deletedDistributionIds != null && deletedDistributionIds.Any())
                {
                    const string deleteDistSql = "DELETE FROM QMS_Documents_DistributionList WHERE EntryID IN @DeletedIds";
                    await db.ExecuteAsync(deleteDistSql, new { DeletedIds = deletedDistributionIds }, trans);
                }

                // Insert new distribution list rows
                if (model.DistributionList != null)
                {
                    var newEntries = model.DistributionList.Where(d => d.EntryID == 0 && !string.IsNullOrWhiteSpace(d.EmpID)).ToList();
                    if (newEntries.Any())
                    {
                        const string insertDistSql = @"
                            INSERT INTO QMS_Documents_DistributionList (QMS_RefID, EmpID)
                            VALUES (@QMS_RefID, @EmpID);";

                        foreach (var dist in newEntries)
                        {
                            await db.ExecuteAsync(insertDistSql, new
                            {
                                QMS_RefID = docId,
                                dist.EmpID
                            }, trans);
                        }
                    }
                }

                trans.Commit();
                return docId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving QMS document.");
                throw;
            }
        }

        public async Task<bool> DeleteDocumentAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                // Delete distribution records first
                await db.ExecuteAsync("DELETE FROM QMS_Documents_DistributionList WHERE QMS_RefID = @EntryID", new { EntryID = entryId }, trans);

                // Delete associated change requests if any
                await db.ExecuteAsync("DELETE FROM QMS_Documents_ChangeRequests WHERE QMS_RefID = @EntryID", new { EntryID = entryId }, trans);

                // Delete modern attachments
                await db.ExecuteAsync("DELETE FROM EntityAttachments WHERE EntityType = 'QMSDocument' AND EntityId = @EntityId", new { EntityId = entryId.ToString() }, trans);

                // Delete document master record
                int rows = await db.ExecuteAsync("DELETE FROM QMS_Documents WHERE EntryID = @EntryID", new { EntryID = entryId }, trans);

                trans.Commit();
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting QMS document with ID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<List<EmployeeDropdownDto>> GetEmployeesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        empid AS EmpID,
                        name AS Name,
                        fname AS FName,
                        DeptName,
                        Designation
                    FROM VEmp
                    WHERE (Active = 1 OR Active IS NULL)
                      AND empid IS NOT NULL AND RTRIM(LTRIM(empid)) <> ''
                      AND name IS NOT NULL AND RTRIM(LTRIM(name)) <> ''
                    ORDER BY name";

                var results = await db.QueryAsync<EmployeeDropdownDto>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading employees lookup for QMS documents.");
                return new List<EmployeeDropdownDto>();
            }
        }

        public async Task<List<string>> GetDocumentTypesListAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT DocumentType FROM QMS_DocumentTypes ORDER BY DocumentType";
                var results = await db.QueryAsync<string>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading document types.");
                return new List<string>();
            }
        }
    }
}

