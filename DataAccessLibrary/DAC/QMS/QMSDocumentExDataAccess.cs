using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.QMS
{
    public class QMSDocumentExDataAccess : IQMSDocumentExDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<QMSDocumentExDataAccess> _logger;

        public QMSDocumentExDataAccess(IConfiguration configuration, ILogger<QMSDocumentExDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<QMSDocumentExListItemDto>> GetDocumentsAsync(string masterCategory, string? searchText = null, DateTime? fromDate = null, DateTime? toDate = null)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        v.EntryID,
                        v.DocumentNumber,
                        v.DocumentName,
                        v.Revision,
                        v.DocumentType,
                        v.Master_DocumentType,
                        v.IssueDT,
                        v.InitiaterEmpID,
                        v.name AS InitiatorName,
                        v.DeptName,
                        v.LastApprovedEntryDT
                    FROM VQMS_Documents_Ex v
                    WHERE v.Master_DocumentType = @MasterCategory
                      AND (@Search IS NULL OR v.DocumentNumber LIKE '%' + @Search + '%' OR v.DocumentName LIKE '%' + @Search + '%' OR v.DocumentType LIKE '%' + @Search + '%' OR v.name LIKE '%' + @Search + '%')
                      AND (@FromDate IS NULL OR v.IssueDT >= @FromDate)
                      AND (@ToDate IS NULL OR v.IssueDT <= @ToDate)
                    ORDER BY v.IssueDT DESC, v.EntryID DESC;";

                var paramSearch = string.IsNullOrWhiteSpace(searchText) ? null : searchText.Trim();
                var results = await db.QueryAsync<QMSDocumentExListItemDto>(sql, new
                {
                    MasterCategory = masterCategory.Trim(),
                    Search = paramSearch,
                    FromDate = fromDate,
                    ToDate = toDate
                });

                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Documents for category {Category}", masterCategory);
                throw;
            }
        }

        public async Task<QMSDocumentExModel?> GetDocumentByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        v.EntryID,
                        v.DocumentNumber,
                        v.DocumentName,
                        v.Revision,
                        v.DocumentType,
                        v.Master_DocumentType,
                        v.IssueDT,
                        v.InitiaterEmpID,
                        v.name AS InitiatorName,
                        v.DeptName AS InitiatorDept,
                        ISNULL(d.Internal, 1) AS Internal
                    FROM VQMS_Documents_Ex v
                    LEFT JOIN QMS_Documents_Ex d ON v.EntryID = d.EntryID
                    WHERE v.EntryID = @EntryID;";

                return await db.QueryFirstOrDefaultAsync<QMSDocumentExModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS Document ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<List<QMSDocumentExDistributionItemModel>> GetDistributionListAsync(int docEntryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        EntryID,
                        QMS_RefID,
                        EmpID,
                        Name,
                        DeptName,
                        Designation
                    FROM VQMS_Documents_Ex_DistributionList
                    WHERE QMS_RefID = @DocEntryId
                    ORDER BY Name;";

                var results = await db.QueryAsync<QMSDocumentExDistributionItemModel>(sql, new { DocEntryId = docEntryId });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Distribution List for QMS Document ID {DocEntryId}", docEntryId);
                throw;
            }
        }

        public async Task<List<string>> GetDocumentTypesByMasterAsync(string masterCategory)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT DocumentType 
                    FROM QMS_Master_DocumentTypes_DocTypes 
                    WHERE MasterDocumentType = @MasterCategory OR ISNULL(MasterDocumentType, '') = ''
                    ORDER BY DocumentType;";

                var results = await db.QueryAsync<string>(sql, new { MasterCategory = masterCategory.Trim() });
                var list = results.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                if (!list.Any())
                {
                    // Fallback to default types if none registered in lookup
                    list.AddRange(new[] { "Standard", "Policy", "Manual", "Procedure", "Work Instruction", "Record" });
                }
                return list;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Document Types for master category {Category}", masterCategory);
                return new List<string> { "Standard", "Policy", "Manual", "Procedure", "Work Instruction" };
            }
        }

        public async Task<List<EmployeeDropdownDto>> GetEmployeesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        EmpID,
                        Name,
                        FName,
                        DeptName,
                        Designation
                    FROM VEmp
                    WHERE Active = 1
                    ORDER BY Name;";

                var results = await db.QueryAsync<EmployeeDropdownDto>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Employees for dropdown");
                return new List<EmployeeDropdownDto>();
            }
        }

        public async Task<int> SaveDocumentAsync(QMSDocumentExModel model, List<QMSDocumentExDistributionItemModel> distList)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                int entryId = model.EntryID;

                if (entryId == 0)
                {
                    const string insertDocSql = @"
                        INSERT INTO QMS_Documents_Ex (
                            DocumentNumber, DocumentName, Revision, DocumentType, Master_DocumentType,
                            IssueDT, InitiaterEmpID, Internal
                        ) VALUES (
                            @DocumentNumber, @DocumentName, @Revision, @DocumentType, @Master_DocumentType,
                            @IssueDT, @InitiaterEmpID, @Internal
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    entryId = await db.ExecuteScalarAsync<int>(insertDocSql, new
                    {
                        model.DocumentNumber,
                        model.DocumentName,
                        Revision = model.Revision ?? "00",
                        model.DocumentType,
                        model.Master_DocumentType,
                        model.IssueDT,
                        model.InitiaterEmpID,
                        model.Internal
                    }, tx);
                }
                else
                {
                    const string updateDocSql = @"
                        UPDATE QMS_Documents_Ex 
                        SET DocumentNumber = @DocumentNumber,
                            DocumentName = @DocumentName,
                            Revision = @Revision,
                            DocumentType = @DocumentType,
                            Master_DocumentType = @Master_DocumentType,
                            IssueDT = @IssueDT,
                            InitiaterEmpID = @InitiaterEmpID,
                            Internal = @Internal
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateDocSql, new
                    {
                        model.DocumentNumber,
                        model.DocumentName,
                        Revision = model.Revision ?? "00",
                        model.DocumentType,
                        model.Master_DocumentType,
                        model.IssueDT,
                        model.InitiaterEmpID,
                        model.Internal,
                        EntryID = entryId
                    }, tx);
                }

                // Handle Distribution List changes
                if (distList != null)
                {
                    // Delete removed items
                    var deleteIds = distList.Where(d => d.IsDeleted && d.EntryID > 0).Select(d => d.EntryID).ToList();
                    if (deleteIds.Any())
                    {
                        const string deleteDistSql = "DELETE FROM QMS_Documents_Ex_DistributionList WHERE EntryID IN @DeleteIds;";
                        await db.ExecuteAsync(deleteDistSql, new { DeleteIds = deleteIds }, tx);
                    }

                    // Insert new items
                    var newItems = distList.Where(d => !d.IsDeleted && (d.IsNew || d.EntryID == 0)).ToList();
                    if (newItems.Any())
                    {
                        const string insertDistSql = @"
                            INSERT INTO QMS_Documents_Ex_DistributionList (QMS_RefID, EmpID)
                            VALUES (@QMS_RefID, @EmpID);";

                        foreach (var item in newItems)
                        {
                            await db.ExecuteAsync(insertDistSql, new
                            {
                                QMS_RefID = entryId,
                                item.EmpID
                            }, tx);
                        }
                    }
                }

                tx.Commit();
                return entryId;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error saving QMS Document {DocumentNumber}", model.DocumentNumber);
                throw;
            }
        }

        public async Task<bool> DeleteDocumentAsync(int entryId)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var tx = db.BeginTransaction();

            try
            {
                const string deleteDistSql = "DELETE FROM QMS_Documents_Ex_DistributionList WHERE QMS_RefID = @EntryID;";
                await db.ExecuteAsync(deleteDistSql, new { EntryID = entryId }, tx);

                const string deleteDocSql = "DELETE FROM QMS_Documents_Ex WHERE EntryID = @EntryID;";
                int affected = await db.ExecuteAsync(deleteDocSql, new { EntryID = entryId }, tx);

                tx.Commit();
                return affected > 0;
            }
            catch (Exception ex)
            {
                tx.Rollback();
                _logger.LogError(ex, "Error deleting QMS Document ID {EntryID}", entryId);
                throw;
            }
        }

        public async Task<string> GetNextDocumentNumberAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT ISNULL(MAX(EntryID), 0) + 1 FROM QMS_Documents_Ex;";
                int nextId = await db.ExecuteScalarAsync<int>(sql);
                return $"DOC-{nextId:D4}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating next document number");
                return $"DOC-{DateTime.Now:yyyyMMddHHmm}";
            }
        }
    }
}
