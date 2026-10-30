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
    public class SupplierReEvaluationDataAccess : ISupplierReEvaluationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<SupplierReEvaluationDataAccess> _logger;

        public SupplierReEvaluationDataAccess(IConfiguration configuration, ILogger<SupplierReEvaluationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<SupplierReEvaluationListItemDto>> GetSupplierReEvaluationsAsync(SupplierReEvaluationFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        se.EntryID,
                        se.AccNo,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle,
                        se.ReviewDT,
                        se.EntryDT,
                        se.ReviewBy,
                        se.PreparedBy,
                        se.ApprovedBy,
                        ISNULL(se.QualityOfProduction, 0) AS QualityOfProduction,
                        ISNULL(se.Price, 0) AS Price,
                        ISNULL(se.Packing, 0) AS Packing,
                        ISNULL(se.RejectionRate, 0) AS RejectionRate,
                        ISNULL(se.DeliveryInTime, 0) AS DeliveryInTime,
                        ISNULL(se.ReWorkRate, 0) AS ReWorkRate,
                        ISNULL(se.NoOfOrdersCompleted, 0) AS NoOfOrdersCompleted,
                        ISNULL(se.CancelledOrders, 0) AS CancelledOrders,
                        ISNULL(se.ResponseTime, 0) AS ResponseTime,
                        ISNULL(se.TechnicalAbility, 0) AS TechnicalAbility,
                        ISNULL(se.ProductionCapacity, 0) AS ProductionCapacity,
                        ISNULL(se.EnvironmentRequirements, 0) AS EnvironmentRequirements,
                        ISNULL(se.LightingEtc, 0) AS LightingEtc
                    FROM dbo.SupplierReEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE (@DateFrom IS NULL OR se.ReviewDT >= @DateFrom)
                      AND (@DateTo IS NULL OR se.ReviewDT <= @DateTo)
                      AND (@AccNo IS NULL OR @AccNo = '' OR se.AccNo = @AccNo)
                      AND (
                            @SearchText IS NULL OR @SearchText = ''
                            OR se.AccNo LIKE '%' + @SearchText + '%'
                            OR vm.AccTitle LIKE '%' + @SearchText + '%'
                            OR a.AccTitle LIKE '%' + @SearchText + '%'
                            OR se.ReviewBy LIKE '%' + @SearchText + '%'
                            OR se.PreparedBy LIKE '%' + @SearchText + '%'
                            OR se.ApprovedBy LIKE '%' + @SearchText + '%'
                      )
                    ORDER BY se.ReviewDT DESC, se.EntryID DESC";

                DateTime? dtToInclusive = filter.DateTo?.Date.AddDays(1).AddTicks(-1);

                var items = await db.QueryAsync<SupplierReEvaluationListItemDto>(sql, new
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
                _logger.LogError(ex, "Error fetching supplier re-evaluations list.");
                throw;
            }
        }

        public async Task<SupplierReEvaluationModel?> GetReEvaluationByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        se.*,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle
                    FROM dbo.SupplierReEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE se.EntryID = @EntryID";

                return await db.QueryFirstOrDefaultAsync<SupplierReEvaluationModel>(sql, new { EntryID = entryId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching supplier re-evaluation with EntryID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<SupplierReEvaluationModel?> GetLatestReEvaluationByAccNoAsync(string accNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 
                        se.*,
                        COALESCE(vm.AccTitle, a.AccTitle, se.AccNo) AS AccTitle
                    FROM dbo.SupplierReEvaluation se
                    LEFT JOIN dbo.Accounts a ON se.AccNo = a.AccNo
                    LEFT JOIN dbo.VVendersAndMakers vm ON se.AccNo = vm.AccNo
                    WHERE se.AccNo = @AccNo
                    ORDER BY se.ReviewDT DESC, se.EntryID DESC";

                return await db.QueryFirstOrDefaultAsync<SupplierReEvaluationModel>(sql, new { AccNo = accNo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching latest supplier re-evaluation for AccNo {AccNo}.", accNo);
                throw;
            }
        }

        public async Task<int> SaveReEvaluationAsync(SupplierReEvaluationModel model)
        {
            try
            {
                using var db = CreateConnection();

                if (model.EntryID == 0)
                {
                    const string insertSql = @"
                        INSERT INTO dbo.SupplierReEvaluation (
                            AccNo, ReviewDT, ReviewBy, PreparedBy, ApprovedBy, EntryDT, UserName, MachineName,
                            QualityOfProduction, QualityOfProductionComments,
                            Price, PriceComments,
                            Packing, PackingComments,
                            RejectionRate, RejectionRateComments,
                            DeliveryInTime, DeliveryInTimeComments,
                            ReWorkRate, ReWorkRateComments,
                            NoOfOrdersCompleted, NoOfOrdersCompletedComments,
                            CancelledOrders, CancelledOrdersComments,
                            ResponseTime, ResponseTimeComments,
                            TechnicalAbility, TechnicalAbilityComments,
                            ProductionCapacity, ProductionCapacityComments,
                            EnvironmentRequirements, EnvironmentRequirementsComments,
                            LightingEtc, LightingEtcComments
                        ) VALUES (
                            @AccNo, @ReviewDT, @ReviewBy, @PreparedBy, @ApprovedBy, @EntryDT, @UserName, @MachineName,
                            @QualityOfProduction, @QualityOfProductionComments,
                            @Price, @PriceComments,
                            @Packing, @PackingComments,
                            @RejectionRate, @RejectionRateComments,
                            @DeliveryInTime, @DeliveryInTimeComments,
                            @ReWorkRate, @ReWorkRateComments,
                            @NoOfOrdersCompleted, @NoOfOrdersCompletedComments,
                            @CancelledOrders, @CancelledOrdersComments,
                            @ResponseTime, @ResponseTimeComments,
                            @TechnicalAbility, @TechnicalAbilityComments,
                            @ProductionCapacity, @ProductionCapacityComments,
                            @EnvironmentRequirements, @EnvironmentRequirementsComments,
                            @LightingEtc, @LightingEtcComments
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    int newId = await db.ExecuteScalarAsync<int>(insertSql, model);
                    model.EntryID = newId;
                    return newId;
                }
                else
                {
                    const string updateSql = @"
                        UPDATE dbo.SupplierReEvaluation SET
                            AccNo = @AccNo,
                            ReviewDT = @ReviewDT,
                            ReviewBy = @ReviewBy,
                            PreparedBy = @PreparedBy,
                            ApprovedBy = @ApprovedBy,
                            QualityOfProduction = @QualityOfProduction,
                            QualityOfProductionComments = @QualityOfProductionComments,
                            Price = @Price,
                            PriceComments = @PriceComments,
                            Packing = @Packing,
                            PackingComments = @PackingComments,
                            RejectionRate = @RejectionRate,
                            RejectionRateComments = @RejectionRateComments,
                            DeliveryInTime = @DeliveryInTime,
                            DeliveryInTimeComments = @DeliveryInTimeComments,
                            ReWorkRate = @ReWorkRate,
                            ReWorkRateComments = @ReWorkRateComments,
                            NoOfOrdersCompleted = @NoOfOrdersCompleted,
                            NoOfOrdersCompletedComments = @NoOfOrdersCompletedComments,
                            CancelledOrders = @CancelledOrders,
                            CancelledOrdersComments = @CancelledOrdersComments,
                            ResponseTime = @ResponseTime,
                            ResponseTimeComments = @ResponseTimeComments,
                            TechnicalAbility = @TechnicalAbility,
                            TechnicalAbilityComments = @TechnicalAbilityComments,
                            ProductionCapacity = @ProductionCapacity,
                            ProductionCapacityComments = @ProductionCapacityComments,
                            EnvironmentRequirements = @EnvironmentRequirements,
                            EnvironmentRequirementsComments = @EnvironmentRequirementsComments,
                            LightingEtc = @LightingEtc,
                            LightingEtcComments = @LightingEtcComments
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, model);
                    return model.EntryID;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving supplier re-evaluation with EntryID {EntryID}.", model.EntryID);
                throw;
            }
        }

        public async Task<bool> DeleteReEvaluationAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM dbo.SupplierReEvaluation WHERE EntryID = @EntryID";
                int rows = await db.ExecuteAsync(sql, new { EntryID = entryId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting supplier re-evaluation with EntryID {EntryID}.", entryId);
                throw;
            }
        }
    }
}

