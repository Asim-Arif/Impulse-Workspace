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
    public class TrackBackDataAccess : ITrackBackDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<TrackBackDataAccess> _logger;

        public TrackBackDataAccess(IConfiguration configuration, ILogger<TrackBackDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<string>> GetRecentCustomInvoicesAsync(string? searchTerm = null)
        {
            try
            {
                using var db = CreateConnection();
                string sql;
                object? param = null;

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    sql = @"
                        SELECT TOP 100 CustomInvoice 
                        FROM CustomInvoice 
                        WHERE CustomInvoice LIKE @Search 
                        ORDER BY CustomInvoice DESC";
                    param = new { Search = $"%{searchTerm.Trim()}%" };
                }
                else
                {
                    sql = @"
                        SELECT TOP 200 CustomInvoice 
                        FROM CustomInvoice 
                        ORDER BY CustomInvoice DESC";
                }

                var results = await db.QueryAsync<string>(sql, param);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching CustomInvoice list.");
                return new List<string>();
            }
        }

        public async Task<List<TrackBackInvoiceItemDto>> GetInvoiceItemsAsync(string invoiceNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        CompItemCode,
                        ItemName,
                        SUM(Qty) AS Qty
                    FROM VCustomInvoice
                    WHERE CustomInvoice = @InvoiceNo
                    GROUP BY CompItemCode, ItemName
                    ORDER BY CompItemCode";

                var results = (await db.QueryAsync<TrackBackInvoiceItemDto>(sql, new { InvoiceNo = invoiceNo.Trim() })).ToList();
                for (int i = 0; i < results.Count; i++)
                {
                    results[i].RowIndex = i + 1;
                }
                return results;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching items for invoice {InvoiceNo}.", invoiceNo);
                return new List<TrackBackInvoiceItemDto>();
            }
        }

        public async Task<List<TrackBackLotDto>> GetLotsByItemAsync(string invoiceNo, string compItemCode)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        d.LotNo,
                        SUM(d.DLD_VRD_Qty) AS Qty,
                        (SELECT COUNT(1) FROM EntityAttachments ea WHERE (ea.EntityType = 'Lot' OR ea.EntityType = 'PTC') AND ea.EntityId = d.LotNo) AS AttachmentCount
                    FROM VCustomPList_DispatchListDetail d
                    WHERE d.CustomInvoice = @InvoiceNo 
                      AND d.CompItemCode = @CompItemCode
                      AND d.LotNo IS NOT NULL
                    GROUP BY d.LotNo
                    ORDER BY d.LotNo";

                var results = await db.QueryAsync<TrackBackLotDto>(sql, new
                {
                    InvoiceNo = invoiceNo.Trim(),
                    CompItemCode = compItemCode.Trim()
                });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching lots for invoice {InvoiceNo} and item {CompItemCode}.", invoiceNo, compItemCode);
                return new List<TrackBackLotDto>();
            }
        }

        public async Task<List<TrackBackMasterPODto>> GetMasterPOsByItemAsync(string invoiceNo, string compItemCode)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT 
                        l.MasterPONo,
                        l.VenderName
                    FROM VLotsWithMasterPONo l
                    INNER JOIN VCustomPList_DispatchListDetail d ON d.LotNo = l.LotNo
                    WHERE d.CustomInvoice = @InvoiceNo
                      AND d.CompItemCode = @CompItemCode
                      AND l.MasterPONo IS NOT NULL
                    ORDER BY l.MasterPONo";

                var results = await db.QueryAsync<TrackBackMasterPODto>(sql, new
                {
                    InvoiceNo = invoiceNo.Trim(),
                    CompItemCode = compItemCode.Trim()
                });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Master POs for invoice {InvoiceNo} and item {CompItemCode}.", invoiceNo, compItemCode);
                return new List<TrackBackMasterPODto>();
            }
        }

        public async Task<List<TrackBackOrderDto>> GetOrdersByItemAsync(string invoiceNo, string compItemCode)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        OrderNo,
                        SUM(Qty) AS Qty
                    FROM VCustomInvoiceItems
                    WHERE CustomInvoice = @InvoiceNo
                      AND CompItemCode = @CompItemCode
                      AND OrderNo IS NOT NULL
                    GROUP BY OrderNo
                    ORDER BY OrderNo";

                var results = await db.QueryAsync<TrackBackOrderDto>(sql, new
                {
                    InvoiceNo = invoiceNo.Trim(),
                    CompItemCode = compItemCode.Trim()
                });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching orders for invoice {InvoiceNo} and item {CompItemCode}.", invoiceNo, compItemCode);
                return new List<TrackBackOrderDto>();
            }
        }

        public async Task<List<string>> GetInvoiceMasterPONosAsync(string invoiceNo)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT DISTINCT l.MasterPONo
                    FROM VLotsWithMasterPONo l
                    INNER JOIN VCustomPList_DispatchListDetail d ON d.LotNo = l.LotNo
                    WHERE d.CustomInvoice = @InvoiceNo
                      AND l.MasterPONo IS NOT NULL";

                var results = await db.QueryAsync<string>(sql, new { InvoiceNo = invoiceNo.Trim() });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching all Master POs for invoice {InvoiceNo}.", invoiceNo);
                return new List<string>();
            }
        }
    }
}

