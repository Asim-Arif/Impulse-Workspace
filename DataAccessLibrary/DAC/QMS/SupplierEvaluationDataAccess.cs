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
    public class SupplierEvaluationDataAccess : ISupplierEvaluationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<SupplierEvaluationDataAccess> _logger;

        public SupplierEvaluationDataAccess(IConfiguration configuration, ILogger<SupplierEvaluationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<SupplierEvaluationListItemDto>> GetSupplierEvaluationsAsync(SupplierEvaluationFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        se.EntryID,
                        se.AccNo,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle,
                        se.EntryDT,
                        se.TimeApproval,
                        se.MarketingReputation,
                        se.TechnicalCapability,
                        se.Capacity,
                        se.ProductLine,
                        se.Remarks1,
                        se.Remarks
                    FROM dbo.SupplierEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE (@DateFrom IS NULL OR se.EntryDT >= @DateFrom)
                      AND (@DateTo IS NULL OR se.EntryDT <= @DateTo)
                      AND (@AccNo IS NULL OR @AccNo = '' OR se.AccNo = @AccNo)
                      AND (
                            @SearchText IS NULL OR @SearchText = ''
                            OR se.AccNo LIKE '%' + @SearchText + '%'
                            OR vm.AccTitle LIKE '%' + @SearchText + '%'
                            OR a.AccTitle LIKE '%' + @SearchText + '%'
                            OR se.Remarks1 LIKE '%' + @SearchText + '%'
                            OR se.Remarks LIKE '%' + @SearchText + '%'
                            OR se.ProductLine LIKE '%' + @SearchText + '%'
                      )
                    ORDER BY se.EntryDT DESC, se.EntryID DESC";

                DateTime? dtToInclusive = filter.DateTo?.Date.AddDays(1).AddTicks(-1);

                var items = await db.QueryAsync<SupplierEvaluationListItemDto>(sql, new
                {
                    DateFrom = filter.DateFrom?.Date,
                    DateTo = dtToInclusive,
                    AccNo = string.IsNullOrWhiteSpace(filter.AccNo) ? null : filter.AccNo.Trim(),
                    SearchText = string.IsNullOrWhiteSpace(filter.SearchText) ? null : filter.SearchText.Trim()
                });

                return items.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching supplier evaluations list.");
                throw;
            }
        }

        public async Task<SupplierEvaluationModel?> GetEvaluationByAccNoAsync(string accNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        se.EntryID,
                        se.AccNo,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle,
                        se.EntryDT,
                        se.EmpNo,
                        se.MachineNo,
                        se.ToolNo,
                        se.ProductLine,
                        se.CoreCompetencies,
                        se.Remarks1,
                        se.Remarks2,
                        se.Capacity,
                        se.MajorCustomers,
                        se.Generator,
                        se.SamplingHandling,
                        se.UPS,
                        se.Cleaniness,
                        se.InspRecMaintain,
                        se.TracebilityControl,
                        se.CalibrationSystem,
                        se.DrawingComprehension,
                        se.MarketingReputation,
                        se.TechnicalCapability,
                        se.TimeApproval,
                        se.Remarks,
                        se.ISO9000,
                        se.CustComplaintHanding,
                        se.StaffTrainingSystem,
                        se.PreparedBy,
                        se.EvaluatedBy,
                        se.ApprovedBy,
                        se.PDFFileName
                    FROM dbo.SupplierEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE se.AccNo = @AccNo";

                return await db.QueryFirstOrDefaultAsync<SupplierEvaluationModel>(sql, new { AccNo = accNo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching supplier evaluation for AccNo {AccNo}.", accNo);
                throw;
            }
        }

        public async Task<SupplierEvaluationModel?> GetEvaluationByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        se.EntryID,
                        se.AccNo,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle,
                        se.EntryDT,
                        se.EmpNo,
                        se.MachineNo,
                        se.ToolNo,
                        se.ProductLine,
                        se.CoreCompetencies,
                        se.Remarks1,
                        se.Remarks2,
                        se.Capacity,
                        se.MajorCustomers,
                        se.Generator,
                        se.SamplingHandling,
                        se.UPS,
                        se.Cleaniness,
                        se.InspRecMaintain,
                        se.TracebilityControl,
                        se.CalibrationSystem,
                        se.DrawingComprehension,
                        se.MarketingReputation,
                        se.TechnicalCapability,
                        se.TimeApproval,
                        se.Remarks,
                        se.ISO9000,
                        se.CustComplaintHanding,
                        se.StaffTrainingSystem,
                        se.PreparedBy,
                        se.EvaluatedBy,
                        se.ApprovedBy,
                        se.PDFFileName
                    FROM dbo.SupplierEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE se.EntryID = @EntryID";

                return await db.QueryFirstOrDefaultAsync<SupplierEvaluationModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching supplier evaluation for EntryID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<List<SupplierDropdownDto>> GetSuppliersListAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT
                        AccNo,
                        AccTitle,
                        Vender
                    FROM dbo.VVendersAndMakers
                    WHERE Active = 1 AND AccNo IS NOT NULL AND RTRIM(LTRIM(AccNo)) <> ''
                    ORDER BY AccTitle";

                var results = await db.QueryAsync<SupplierDropdownDto>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching suppliers and makers list.");
                throw;
            }
        }

        public async Task<bool> SaveEvaluationAsync(SupplierEvaluationModel model)
        {
            try
            {
                using var db = CreateConnection();
                const string checkSql = "SELECT COUNT(*) FROM dbo.SupplierEvaluation WHERE AccNo = @AccNo";
                int existing = await db.ExecuteScalarAsync<int>(checkSql, new { model.AccNo });

                if (existing > 0)
                {
                    const string updateSql = @"
                        UPDATE dbo.SupplierEvaluation SET
                            EmpNo = @EmpNo,
                            MachineNo = @MachineNo,
                            ToolNo = @ToolNo,
                            ProductLine = @ProductLine,
                            CoreCompetencies = @CoreCompetencies,
                            Remarks1 = @Remarks1,
                            Remarks2 = @Remarks2,
                            Capacity = @Capacity,
                            MajorCustomers = @MajorCustomers,
                            Generator = @Generator,
                            SamplingHandling = @SamplingHandling,
                            UPS = @UPS,
                            Cleaniness = @Cleaniness,
                            InspRecMaintain = @InspRecMaintain,
                            TracebilityControl = @TracebilityControl,
                            CalibrationSystem = @CalibrationSystem,
                            DrawingComprehension = @DrawingComprehension,
                            MarketingReputation = @MarketingReputation,
                            TechnicalCapability = @TechnicalCapability,
                            TimeApproval = @TimeApproval,
                            Remarks = @Remarks,
                            ISO9000 = @ISO9000,
                            CustComplaintHanding = @CustComplaintHanding,
                            StaffTrainingSystem = @StaffTrainingSystem,
                            PreparedBy = @PreparedBy,
                            EvaluatedBy = @EvaluatedBy,
                            ApprovedBy = @ApprovedBy,
                            EntryDT = @EntryDT
                        WHERE AccNo = @AccNo";

                    int rows = await db.ExecuteAsync(updateSql, model);
                    return rows > 0;
                }
                else
                {
                    const string insertSql = @"
                        INSERT INTO dbo.SupplierEvaluation (
                            AccNo, EmpNo, MachineNo, ToolNo, ProductLine, CoreCompetencies,
                            Remarks1, Remarks2, Capacity, MajorCustomers, Generator,
                            SamplingHandling, UPS, Cleaniness, InspRecMaintain,
                            TracebilityControl, CalibrationSystem, DrawingComprehension,
                            MarketingReputation, TechnicalCapability, TimeApproval,
                            Remarks, ISO9000, CustComplaintHanding, StaffTrainingSystem,
                            PreparedBy, EvaluatedBy, ApprovedBy, EntryDT
                        ) VALUES (
                            @AccNo, @EmpNo, @MachineNo, @ToolNo, @ProductLine, @CoreCompetencies,
                            @Remarks1, @Remarks2, @Capacity, @MajorCustomers, @Generator,
                            @SamplingHandling, @UPS, @Cleaniness, @InspRecMaintain,
                            @TracebilityControl, @CalibrationSystem, @DrawingComprehension,
                            @MarketingReputation, @TechnicalCapability, @TimeApproval,
                            @Remarks, @ISO9000, @CustComplaintHanding, @StaffTrainingSystem,
                            @PreparedBy, @EvaluatedBy, @ApprovedBy, @EntryDT
                        )";

                    int rows = await db.ExecuteAsync(insertSql, model);
                    return rows > 0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving supplier evaluation for AccNo {AccNo}.", model.AccNo);
                throw;
            }
        }

        public async Task<bool> DeleteEvaluationAsync(string accNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM dbo.SupplierEvaluation WHERE AccNo = @AccNo";
                int rows = await db.ExecuteAsync(sql, new { AccNo = accNo });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting supplier evaluation for AccNo {AccNo}.", accNo);
                throw;
            }
        }
    }
}

