using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class CustomerQuotationDataAccess : ICustomerQuotationDataAccess
    {
        private readonly string _connectionString;

        public CustomerQuotationDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerQuotationListItemModel>> GetQuotationListAsync(
            DateTime dtFrom,
            DateTime dtTo,
            string custCode,
            string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT 
                    v.*, 
                    fc.Name AS CustomerName 
                FROM VCustomerQuotationsList v
                LEFT JOIN ForeignCustomers fc 
                    ON v.CustCode = fc.CustCode AND v.Country = fc.Country
                WHERE v.DT >= @DtFrom AND v.DT <= @DtTo";

            if (custCode != "0" && !string.IsNullOrEmpty(custCode))
            {
                sql += " AND v.CustCode = @CustCode";
            }
            if (country != "<All Countries>" && !string.IsNullOrEmpty(country))
            {
                sql += " AND v.Country = @Country";
            }

            sql += " ORDER BY v.DT DESC";

            var results = await db.QueryAsync<CustomerQuotationListItemModel>(
                sql,
                new { DtFrom = dtFrom.Date, DtTo = dtTo.Date.AddDays(1).AddTicks(-1), CustCode = custCode, Country = country });

            return results.ToList();
        }

        public async Task<CustomerQuotationViewModel?> GetQuotationAsync(int quotationNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sqlMaster = "SELECT * FROM CustomerQuotations WHERE QuotationNo = @QuotationNo";
            var quotation = await db.QuerySingleOrDefaultAsync<CustomerQuotationViewModel>(sqlMaster, new { QuotationNo = quotationNo });

            if (quotation != null)
            {
                // Load Items
                string sqlItems = @"
                    SELECT 
                        cqi.*,
                        i.ItemName AS ItemName 
                    FROM CustomerQuotationsItems cqi
                    LEFT JOIN Items i ON cqi.ItemCode = i.ItemID
                    WHERE cqi.QuotationNo = @QuotationNo
                    ORDER BY cqi.EntryID";
                var items = await db.QueryAsync<CustomerQuotationItemViewModel>(sqlItems, new { QuotationNo = quotationNo });
                quotation.QuotationItems = items.ToList();

                // Load Other Charges
                string sqlCharges = "SELECT * FROM CustomerQuotations_OtherCharges WHERE QuotationNo = @QuotationNo ORDER BY SNo";
                var charges = await db.QueryAsync<CustomerQuotationOtherChargeViewModel>(sqlCharges, new { QuotationNo = quotationNo });
                quotation.OtherCharges = charges.ToList();
            }

            return quotation;
        }

        public async Task<int> GetNextQuotationNoAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT ISNULL(MAX(QuotationNo), 0) + 1 FROM CustomerQuotations";
            return await db.ExecuteScalarAsync<int>(sql);
        }

        public async Task<bool> SaveQuotationAsync(CustomerQuotationViewModel quotation, List<int> deletedItemIds, List<int> deletedChargeIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var tran = db.BeginTransaction();
            try
            {
                var existing = await db.ExecuteScalarAsync<int>("SELECT COUNT(1) FROM CustomerQuotations WHERE QuotationNo = @QuotationNo", new { quotation.QuotationNo }, transaction: tran);
                if (existing > 0)
                {
                    // Update
                    string sqlUpdate = @"UPDATE CustomerQuotations SET
                                            DT = @DT, CustCode = @CustCode, Country = @Country, Validity = @Validity,
                                            PaymentTerms = @PaymentTerms, DeliveryDT = @DeliveryDT, ShipVia = @ShipVia,
                                            Packing = @Packing, DeliveryTime = @DeliveryTime, CompanyRefID = @CompanyRefID,
                                            TradeTerms = @TradeTerms, BankID = @BankID, UserName = @UserName, MachineName = @MachineName,
                                            DTEntry = @DTEntry, Remarks = @Remarks, CustomerRef = @CustomerRef, Department = @Department,
                                            AttentionPerson = @AttentionPerson, SwiftCode = @SwiftCode, GrossWeight = @GrossWeight,
                                            NetWeight = @NetWeight, NoOfPackages = @NoOfPackages, RevisionHistory = @RevisionHistory,
                                            Destination = @Destination
                                         WHERE QuotationNo = @QuotationNo";
                    await db.ExecuteAsync(sqlUpdate, quotation, transaction: tran);
                }
                else
                {
                    // Insert
                    string sqlInsert = @"INSERT INTO CustomerQuotations 
                                            (QuotationNo, DT, CustCode, Country, Validity, PaymentTerms, DeliveryDT, ShipVia, Packing, DeliveryTime,
                                             CompanyRefID, TradeTerms, BankID, UserName, MachineName, DTEntry, Remarks, CustomerRef, Department,
                                             AttentionPerson, SwiftCode, GrossWeight, NetWeight, NoOfPackages, RevisionHistory, Destination)
                                         VALUES 
                                            (@QuotationNo, @DT, @CustCode, @Country, @Validity, @PaymentTerms, @DeliveryDT, @ShipVia, @Packing, @DeliveryTime,
                                             @CompanyRefID, @TradeTerms, @BankID, @UserName, @MachineName, @DTEntry, @Remarks, @CustomerRef, @Department,
                                             @AttentionPerson, @SwiftCode, @GrossWeight, @NetWeight, @NoOfPackages, @RevisionHistory, @Destination)";
                    await db.ExecuteAsync(sqlInsert, quotation, transaction: tran);
                }

                // Delete Items
                if (deletedItemIds != null && deletedItemIds.Any())
                {
                    await db.ExecuteAsync("DELETE FROM CustomerQuotationsItems WHERE EntryID IN @Ids", new { Ids = deletedItemIds }, transaction: tran);
                }
                // Save Items
                foreach (var item in quotation.QuotationItems)
                {
                    if (item.EntryID > 0)
                    {
                        string sqlUpdateItem = @"UPDATE CustomerQuotationsItems SET ItemCode = @ItemCode, Qty = @Qty, Price = @Price, DeliveryDT = @DeliveryDT, QualityType = @QualityType, ItemRemarks = @ItemRemarks WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(sqlUpdateItem, item, transaction: tran);
                    }
                    else
                    {
                        string sqlInsertItem = @"INSERT INTO CustomerQuotationsItems (QuotationNo, ItemCode, Qty, Price, DeliveryDT, QualityType, ItemRemarks)
                                                 VALUES (@QuotationNo, @ItemCode, @Qty, @Price, @DeliveryDT, @QualityType, @ItemRemarks)";
                        item.QuotationNo = quotation.QuotationNo;
                        await db.ExecuteAsync(sqlInsertItem, item, transaction: tran);
                    }
                }

                // Delete Other Charges
                if (deletedChargeIds != null && deletedChargeIds.Any())
                {
                    await db.ExecuteAsync("DELETE FROM CustomerQuotations_OtherCharges WHERE EntryID IN @Ids", new { Ids = deletedChargeIds }, transaction: tran);
                }
                // Save Other Charges
                foreach (var charge in quotation.OtherCharges)
                {
                    if (charge.EntryID > 0)
                    {
                        string sqlUpdateCharge = @"UPDATE CustomerQuotations_OtherCharges SET ChargesDetail = @ChargesDetail, ChargesAmt = @ChargesAmt, SNo = @SNo WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(sqlUpdateCharge, charge, transaction: tran);
                    }
                    else
                    {
                        string sqlInsertCharge = @"INSERT INTO CustomerQuotations_OtherCharges (QuotationNo, ChargesDetail, ChargesAmt, SNo)
                                                   VALUES (@QuotationNo, @ChargesDetail, @ChargesAmt, @SNo)";
                        charge.QuotationNo = quotation.QuotationNo;
                        await db.ExecuteAsync(sqlInsertCharge, charge, transaction: tran);
                    }
                }

                tran.Commit();
                return true;
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }

        public async Task<bool> DeleteQuotationAsync(int quotationNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var tran = db.BeginTransaction();
            try
            {
                await db.ExecuteAsync("DELETE FROM CustomerQuotations_OtherCharges WHERE QuotationNo = @QuotationNo", new { QuotationNo = quotationNo }, transaction: tran);
                await db.ExecuteAsync("DELETE FROM CustomerQuotationsItems WHERE QuotationNo = @QuotationNo", new { QuotationNo = quotationNo }, transaction: tran);
                int rows = await db.ExecuteAsync("DELETE FROM CustomerQuotations WHERE QuotationNo = @QuotationNo", new { QuotationNo = quotationNo }, transaction: tran);
                
                tran.Commit();
                return rows > 0;
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }
    }
}
