using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class ProformaInvoiceDataAccess : IProformaInvoiceDataAccess
    {
        private readonly string _connectionString;

        public ProformaInvoiceDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        // ──────────────────────────────────────────────────────────────────────
        // LOOKUPS
        // ──────────────────────────────────────────────────────────────────────

        public async Task<List<GenericDropDownModel>> GetCompaniesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID AS DropDownValue_ID, CompanyName AS DropDownValue_Description FROM Companies ORDER BY CompanyName";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, CustCode AS Name FROM ForeignCustomers WHERE Active = 1 GROUP BY CustCode ORDER BY CustCode";
            return (await db.QueryAsync<CustomerLookupModel>(sql)).ToList();
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";
            return (await db.QueryAsync<string>(sql, new { CustCode = custCode })).ToList();
        }

        public async Task<string> GetCurrencyForCustomerAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT TOP 1 Curr FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
            return await db.QueryFirstOrDefaultAsync<string>(sql, new { CustCode = custCode, Country = country }) ?? string.Empty;
        }

        public async Task<List<GenericDropDownModel>> GetTradeTermsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID AS DropDownValue_ID, Description AS DropDownValue_Description FROM TradeTerms ORDER BY ID";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetPaymentTermsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID AS DropDownValue_ID, Description AS DropDownValue_Description FROM PaymentTerms ORDER BY Description";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetShippingMethodsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID AS DropDownValue_ID, Description AS DropDownValue_Description FROM ShippingMethods ORDER BY EntryID";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<AgentLookupModel>> GetShippingAgentsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID, Name FROM ShippingAgents WHERE Active = 1 ORDER BY Name";
            return (await db.QueryAsync<AgentLookupModel>(sql)).ToList();
        }

        public async Task<List<AgentLookupModel>> GetInsuranceAgentsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID, Name FROM InsuranceAgents WHERE Active = 1 ORDER BY Name";
            return (await db.QueryAsync<AgentLookupModel>(sql)).ToList();
        }

        public async Task<List<CustomerBankModel>> GetCustomerBanksAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT BankID, BankName, ISNULL(Address1,'') AS Address1
                FROM FCustBanks
                WHERE CustCode = @CustCode AND Country = @Country
                ORDER BY BankName";
            return (await db.QueryAsync<CustomerBankModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<CompanyBankAccountModel>> GetCompanyBankAccountsAsync(int companyId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT BankID, ISNULL(Bank,'') AS BankName, AccNo + ' (' + ISNULL(Branch,'') + ')' AS DisplayText
                FROM BankList
                WHERE CompanyRefID = @CompanyID
                ORDER BY Bank, AccNo";
            return (await db.QueryAsync<CompanyBankAccountModel>(sql, new { CompanyID = companyId })).ToList();
        }

        public async Task<List<ShippingAddressPresetModel>> GetShippingAddressPresetsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT EntryID, ShippingAddress
                FROM ForeignCustomersShippingAddresses
                WHERE CustCode = @CustCode AND Country = @Country
                ORDER BY EntryID";
            return (await db.QueryAsync<ShippingAddressPresetModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<InvoiceToPresetModel>> GetInvoiceToPresetsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT EntryID, InvoiceTo
                FROM ForeignCustomers_InvoiceTo
                WHERE CustCode = @CustCode AND Country = @Country
                ORDER BY EntryID";
            return (await db.QueryAsync<InvoiceToPresetModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<string>> GetOrdersForProformaAsync(int companyId, string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            // Uses the same logic as SP_OrdersForProforma
            const string sql = @"
                SELECT OrderNo FROM VrptOrders
                WHERE CustCode = @CustCode AND Country = @Country AND CompanyRefID = @CompanyID
                  AND ISNULL(Cancelled,0) = 0 AND InvQty > 0
                GROUP BY OrderNo
                HAVING SUM(InvQty) > 0
                ORDER BY OrderNo";
            return (await db.QueryAsync<string>(sql, new { CompanyID = companyId, CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<OrderItemForProformaModel>> GetOrderItemsForProformaAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT
                    OrderItemID,
                    OrderNo,
                    ItemCode,
                    CompItemCode,
                    Description,
                    Unit,
                    Price,
                    ISNULL(CustomPrice,0) AS CustomPrice,
                    InvQty,
                    TradeTerms,
                    PartialShipment,
                    TransShipment,
                    PaymentTerms,
                    ISNULL(Packaging,'') AS Packaging,
                    DeliveryDT,
                    DT,
                    DeliveryDTItem,
                    ISNULL(QualityItem,'') AS QualityItem,
                    ISNULL(StampsItem,'') AS StampsItem
                FROM VFCustOrders
                WHERE OrderNo = @OrderNo AND InvQty > 0
                ORDER BY CompItemCode";
            var items = (await db.QueryAsync<OrderItemForProformaModel>(sql, new { OrderNo = orderNo })).ToList();
            // Default QtyToAdd = full available qty
            foreach (var i in items) i.QtyToAdd = i.InvQty;
            return items;
        }

        public async Task<List<CityModel>> GetCitiesAsync(string countryName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CityID, City AS CityName FROM Cities
                WHERE CountryName = @CountryName
                ORDER BY City";
            return (await db.QueryAsync<CityModel>(sql, new { CountryName = countryName })).ToList();
        }

        public async Task<List<PortModel>> GetPortsAsync(int cityId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT PortID, Port, CityID FROM Ports
                WHERE CityID = @CityID
                ORDER BY Port";
            return (await db.QueryAsync<PortModel>(sql, new { CityID = cityId })).ToList();
        }

        public async Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT p.CityID, c.CountryName 
                FROM Ports p 
                INNER JOIN Cities c ON p.CityID = c.CityID
                WHERE p.PortID = @PortID";
            var row = await db.QueryFirstOrDefaultAsync(sql, new { PortID = portId });
            if (row == null) return (0, string.Empty);
            return (row.CityID, row.CountryName ?? string.Empty);
        }

        public async Task<double> GetCustomerDefaultPaymentTermsAsync(string custCode, string country)
            => 0; // Not used; defaults come from GetCustomerCountryDefaultsAsync

        public async Task<string> GetDefaultShipMethodAsync(string custCode, string country)
            => string.Empty; // Defaults come from GetCustomerCountryDefaultsAsync

        public async Task<(bool AcceptsExtraQty, string DefaultShipMethod, string PaymentTerms)> GetCustomerCountryDefaultsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT TOP 1
                    ISNULL(AcceptsExtraQty,0) AS AcceptsExtraQty,
                    ISNULL(DefaultShipMethod,'') AS DefaultShipMethod,
                    ISNULL(PaymentTerms,'') AS PaymentTerms
                FROM ForeignCustomers
                WHERE CustCode = @CustCode AND Country = @Country";
            var row = await db.QueryFirstOrDefaultAsync(sql, new { CustCode = custCode, Country = country });
            if (row == null) return (false, string.Empty, string.Empty);
            return ((bool)row.AcceptsExtraQty, (string)row.DefaultShipMethod, (string)row.PaymentTerms);
        }

        // ──────────────────────────────────────────────────────────────────────
        // PROFORMA # GENERATION
        // ──────────────────────────────────────────────────────────────────────

        public async Task<string> GetNextProformaNumberAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT ISNULL(MAX(CAST(SUBSTRING(PInvoice, CHARINDEX('-', PInvoice)+1, 50) AS INT)), 0) + 1
                FROM FPInvoice
                WHERE ISNUMERIC(SUBSTRING(PInvoice, CHARINDEX('-', PInvoice)+1, 50)) = 1
                  AND CHARINDEX('-', PInvoice) > 0";
            long nextNo = await db.QueryFirstOrDefaultAsync<long>(sql);
            return $"PI-{nextNo}";
        }

        // ──────────────────────────────────────────────────────────────────────
        // LOAD PROFORMA
        // ──────────────────────────────────────────────────────────────────────

        public async Task<ProformaInvoiceViewModel?> GetProformaInvoiceAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Header
            const string headerSql = @"
                SELECT
                    PInvoiceID, PInvoice, DT, CustCode, Country, SMethod,
                    LoadPort, DischargePort, TradeTerms, PaymentTerms, LCNo, BankID,
                    PartialShipment, TransShipment, packaging AS Packaging, DeliveryDT,
                    ISNULL(GrossW,'') AS GrossWeight, ISNULL(NetW,'') AS NetWeight,
                    ISNULL(Packages,'') AS Packages,
                    ShippingAgent AS ShippingAgentID, InsuranceAgent AS InsuranceAgentID,
                    ISNULL(BatchNo,0) AS BatchNo, ISNULL(Declarations,'') AS Declarations,
                    OurBankID, CompanyRefID,
                    ISNULL(InvoiceOf,'Surgical Instruments') AS InvoiceOf,
                    ISNULL(DeliveryDetail,'') AS DeliveryDetail,
                    ISNULL(ShippingAddress,'') AS ShippingAddress,
                    ISNULL(InvoiceTo,'') AS InvoiceTo,
                    ISNULL(Notes1,'') AS Notes1, ISNULL(Notes2,'') AS Notes2,
                    ISNULL(DiscountPer,0) AS DiscountPer,
                    ISNULL(DiscountAmount,0) AS DiscountAmount,
                    ISNULL(Insurance_Freight_Amount,0) AS InsuranceFreightAmount,
                    ISNULL(Insurance_Freight_Amount_II,0) AS InsuranceFreightAmountII,
                    ISNULL(Discount_Description,'DISCOUNT') AS DiscountDescription,
                    ISNULL(Insurance_Description,'FREIGHT') AS InsuranceDescription,
                    ISNULL(Insurance_Description_II,'') AS InsuranceDescriptionII,
                    ISNULL(CIF_Description,'GRAND TOTAL') AS CIFDescription,
                    ISNULL(Total_Description,'TOTAL') AS TotalDescription
                FROM FPInvoice
                WHERE PInvoice = @PInvoice";

            var vm = await db.QueryFirstOrDefaultAsync<ProformaInvoiceViewModel>(headerSql, new { PInvoice = pInvoice });
            if (vm == null) return null;

            // Steel weight
            const string steelSql = "SELECT ISNULL(SteelWeight,0) FROM FPInvoiceSteelWeight WHERE PInvoice = @PInvoice";
            vm.SteelWeight = await db.QueryFirstOrDefaultAsync<double>(steelSql, new { PInvoice = pInvoice });

            // Line items
            const string linesSql = @"
                SELECT
                    po.EntryID,
                    po.OrderEntryID AS OrderItemID,
                    ISNULL(v.OrderNo,'') AS OrderNo,
                    ISNULL(v.ItemID, '') AS ItemCode,
                    ISNULL(v.CompItemID,'') AS CompItemCode,
                    ISNULL(po.ItemDescription, v.Description, '') AS Description,
                    ISNULL(v.Unit,'') AS Unit,
                    ISNULL(po.Price,0) AS Price,
                    ISNULL(po.CustomPrice,0) AS CustomPrice,
                    ISNULL(po.Qty,0) AS Qty,
                    ISNULL(po.InvQty,0) AS InvQty,
                    ISNULL(po.Qty,0) AS InvoicedQty,
                    ISNULL(po.Item_Highlight,0) AS IsHighlighted,
                    v.DeliveryDTItem,
                    ISNULL(v.QualityItem,'') AS QualityItem,
                    ISNULL(v.StampsItem,'') AS StampsItem
                FROM FProformaOrders po
                LEFT JOIN VProformaDetail v ON v.EntryID = po.EntryID
                WHERE po.PInvoice = @PInvoice
                ORDER BY po.EntryID";

            // VProformaDetail has column issue: use a simpler join
            const string linesSqlSimple = @"
                SELECT
                    po.EntryID,
                    po.OrderEntryID AS OrderItemID,
                    ISNULL(fi.OrderNo,'') AS OrderNo,
                    ISNULL(fi.ItemID,'') AS ItemCode,
                    ISNULL(fi.CompItemID,'') AS CompItemCode,
                    ISNULL(po.ItemDescription, ISNULL(fi.ItemDescription,'')) AS Description,
                    ISNULL(fi.Unit,'') AS Unit,
                    ISNULL(po.Price,0) AS Price,
                    ISNULL(po.CustomPrice,0) AS CustomPrice,
                    ISNULL(po.Qty,0) AS Qty,
                    ISNULL(po.InvQty,0) AS InvQty,
                    ISNULL(po.Qty,0) - ISNULL(po.InvQty,0) AS InvoicedQty,
                    ISNULL(po.Item_Highlight,0) AS IsHighlighted,
                    fi.DeliveryDTItem,
                    ISNULL(fi.QualityItem,'') AS QualityItem,
                    ISNULL(fi.StampsItem,'') AS StampsItem
                FROM FProformaOrders po
                LEFT JOIN VProformaDetail fi ON fi.EntryID = po.EntryID AND fi.PInvoice = po.PInvoice
                WHERE po.PInvoice = @PInvoice
                ORDER BY po.EntryID";

            vm.LineItems = (await db.QueryAsync<ProformaLineItemModel>(linesSqlSimple, new { PInvoice = pInvoice })).ToList();

            // Other charges
            const string chargesSql = @"
                SELECT SNo, ISNULL(ChargesDetail,'') AS ChargesDetail, ISNULL(ChargesAmt,0) AS ChargesAmt
                FROM FPInvoiceOtherCharges
                WHERE PInvoice = @PInvoice
                ORDER BY SNo";
            vm.OtherCharges = (await db.QueryAsync<ProformaOtherChargeModel>(chargesSql, new { PInvoice = pInvoice })).ToList();

            return vm;
        }

        // ──────────────────────────────────────────────────────────────────────
        // SAVE PROFORMA (INSERT / UPDATE in a transaction)
        // ──────────────────────────────────────────────────────────────────────

        public async Task<bool> SaveProformaInvoiceAsync(ProformaInvoiceViewModel dto)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                bool isEdit = dto.PInvoiceID > 0;

                if (isEdit)
                {
                    // Restore InvQty on FOrderItems for all current lines (will be re-deducted below)
                    await db.ExecuteAsync(@"
                        UPDATE FOrderItems
                        SET InvQty = InvQty + ISNULL(
                            (SELECT Qty FROM FProformaOrders WHERE PInvoice = @PInvoice AND OrderEntryID = ID),0)
                        WHERE ID IN (SELECT OrderEntryID FROM FProformaOrders WHERE PInvoice = @PInvoice)",
                        new { PInvoice = dto.PInvoice }, trans);

                    // Update header
                    await db.ExecuteAsync(@"
                        UPDATE FPInvoice SET
                            PInvoice=@PInvoice, DT=@DT, SMethod=@SMethod,
                            LoadPort=@LoadPort, DischargePort=@DischargePort,
                            TradeTerms=@TradeTerms, PaymentTerms=@PaymentTerms, LCNo=@LCNo,
                            BankID=@BankID, packaging=@Packaging, PartialShipment=@PartialShipment,
                            TransShipment=@TransShipment, DeliveryDT=@DeliveryDT,
                            ShippingAgent=@ShippingAgentID, InsuranceAgent=@InsuranceAgentID,
                            BatchNo=@BatchNo, Declarations=@Declarations, OurBankID=@OurBankID,
                            InvoiceOf=@InvoiceOf, DeliveryDetail=@DeliveryDetail,
                            GrossW=@GrossWeight, NetW=@NetWeight, Packages=@Packages,
                            ShippingAddress=@ShippingAddress, InvoiceTo=@InvoiceTo,
                            Notes1=@Notes1, Notes2=@Notes2,
                            DiscountPer=@DiscountPer, DiscountAmount=@DiscountAmount,
                            Insurance_Freight_Amount=@InsuranceFreightAmount,
                            Insurance_Freight_Amount_II=@InsuranceFreightAmountII,
                            Discount_Description=@DiscountDescription,
                            Insurance_Description=@InsuranceDescription,
                            Insurance_Description_II=@InsuranceDescriptionII,
                            CIF_Description=@CIFDescription,
                            Total_Description=@TotalDescription
                        WHERE PInvoiceID=@PInvoiceID",
                        dto, trans);

                    // Update steel weight
                    await db.ExecuteAsync(@"
                        UPDATE FPInvoiceSteelWeight SET SteelWeight=@SteelWeight WHERE PInvoice=@PInvoice",
                        new { dto.SteelWeight, dto.PInvoice }, trans);

                    // Delete removed line items
                    if (dto.DeletedLineItemIds?.Count > 0)
                    {
                        await db.ExecuteAsync(@"
                            DELETE FROM FProformaOrders WHERE EntryID IN @Ids",
                            new { Ids = dto.DeletedLineItemIds }, trans);
                    }

                    // Delete all other charges (will re-insert below)
                    await db.ExecuteAsync(@"
                        DELETE FROM FPInvoiceOtherCharges WHERE PInvoice=@PInvoice",
                        new { dto.PInvoice }, trans);
                }
                else
                {
                    // Insert header
                    await db.ExecuteAsync(@"
                        INSERT INTO FPInvoice (
                            PInvoice, DT, CustCode, Country, SMethod,
                            LoadPort, DischargePort, TradeTerms, PaymentTerms, LCNo, BankID,
                            packaging, PartialShipment, TransShipment, DeliveryDT,
                            ShippingAgent, InsuranceAgent, BatchNo, Declarations, OurBankID,
                            CompanyRefID, InvoiceOf, DeliveryDetail, GrossW, NetW, Packages,
                            ShippingAddress, InvoiceTo, Notes1, Notes2,
                            DiscountPer, DiscountAmount, Insurance_Freight_Amount, Insurance_Freight_Amount_II,
                            Discount_Description, Insurance_Description, Insurance_Description_II,
                            CIF_Description, Total_Description
                        ) VALUES (
                            @PInvoice, @DT, @CustCode, @Country, @SMethod,
                            @LoadPort, @DischargePort, @TradeTerms, @PaymentTerms, @LCNo, @BankID,
                            @Packaging, @PartialShipment, @TransShipment, @DeliveryDT,
                            @ShippingAgentID, @InsuranceAgentID, @BatchNo, @Declarations, @OurBankID,
                            @CompanyRefID, @InvoiceOf, @DeliveryDetail, @GrossWeight, @NetWeight, @Packages,
                            @ShippingAddress, @InvoiceTo, @Notes1, @Notes2,
                            @DiscountPer, @DiscountAmount, @InsuranceFreightAmount, @InsuranceFreightAmountII,
                            @DiscountDescription, @InsuranceDescription, @InsuranceDescriptionII,
                            @CIFDescription, @TotalDescription
                        )",
                        dto, trans);

                    // Insert steel weight
                    await db.ExecuteAsync(@"
                        INSERT INTO FPInvoiceSteelWeight (PInvoice, SteelWeight) VALUES (@PInvoice, @SteelWeight)",
                        new { dto.PInvoice, dto.SteelWeight }, trans);
                }

                // Upsert line items
                foreach (var item in dto.LineItems)
                {
                    if (item.EntryID > 0)
                    {
                        // Update existing
                        await db.ExecuteAsync(@"
                            UPDATE FProformaOrders SET
                                Qty=@Qty, InvQty=@Qty, Price=@Price, CustomPrice=@CustomPrice,
                                ItemDescription=@Description, Item_Highlight=@IsHighlighted
                            WHERE EntryID=@EntryID",
                            item, trans);
                    }
                    else
                    {
                        // Insert new
                        await db.ExecuteAsync(@"
                            INSERT INTO FProformaOrders (PInvoice, OrderEntryID, Qty, InvQty, Price, CustomPrice, ItemDescription, Item_Highlight)
                            VALUES (@PInvoice, @OrderItemID, @Qty, @Qty, @Price, @CustomPrice, @Description, @IsHighlighted)",
                            new
                            {
                                dto.PInvoice,
                                item.OrderItemID,
                                item.Qty,
                                item.Price,
                                item.CustomPrice,
                                item.Description,
                                item.IsHighlighted
                            }, trans);
                    }

                    // Deduct from FOrderItems
                    await db.ExecuteAsync(@"
                        UPDATE FOrderItems SET InvQty = InvQty - @Qty WHERE ID = @OrderItemID",
                        new { item.Qty, item.OrderItemID }, trans);
                }

                // Insert other charges
                int sno = 1;
                foreach (var charge in dto.OtherCharges)
                {
                    await db.ExecuteAsync(@"
                        INSERT INTO FPInvoiceOtherCharges (PInvoice, ChargesDetail, ChargesAmt, SNo)
                        VALUES (@PInvoice, @ChargesDetail, @ChargesAmt, @SNo)",
                        new { dto.PInvoice, charge.ChargesDetail, charge.ChargesAmt, SNo = sno++ }, trans);
                }

                trans.Commit();
                return true;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> DeleteProformaInvoiceAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                // Restore InvQty
                await db.ExecuteAsync(@"
                    UPDATE FOrderItems SET InvQty = InvQty +
                        ISNULL((SELECT Qty FROM FProformaOrders WHERE PInvoice=@PInvoice AND OrderEntryID=ID),0)",
                    new { PInvoice = pInvoice }, trans);

                await db.ExecuteAsync("DELETE FROM FProformaOrders WHERE PInvoice=@PInvoice", new { PInvoice = pInvoice }, trans);
                await db.ExecuteAsync("DELETE FROM FPInvoiceOtherCharges WHERE PInvoice=@PInvoice", new { PInvoice = pInvoice }, trans);
                await db.ExecuteAsync("DELETE FROM FPInvoiceSteelWeight WHERE PInvoice=@PInvoice", new { PInvoice = pInvoice }, trans);
                await db.ExecuteAsync("DELETE FROM FPInvoice WHERE PInvoice=@PInvoice", new { PInvoice = pInvoice }, trans);

                trans.Commit();
                return true;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<double> CalculateSteelWeightAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT ISNULL(SUM(Qty * ISNULL(FinishedWeight,0)),0) / 1000.0
                FROM VrptProformas
                WHERE PInvoice = @PInvoice";
            return Math.Round(await db.QueryFirstOrDefaultAsync<double>(sql, new { PInvoice = pInvoice }), 2);
        }
    }
}
