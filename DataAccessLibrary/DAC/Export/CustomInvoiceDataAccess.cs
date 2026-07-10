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
    public class CustomInvoiceDataAccess : ICustomInvoiceDataAccess
    {
        private readonly string _connectionString;

        public CustomInvoiceDataAccess(IConfiguration configuration)
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
                SELECT *
                FROM FCustBanks
                WHERE CustCode = @CustCode AND Country = @Country";
            return (await db.QueryAsync<CustomerBankModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<CompanyBankAccountModel>> GetCompanyBankAccountsAsync(int companyId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT BankID, ISNULL(Bank,'') AS BankName, AccNo + ' (' + ISNULL(Branch,'') + ')' AS DisplayText
                FROM BankList
                WHERE CompanyRefID = @CompanyId
                ORDER BY Bank, AccNo";
            return (await db.QueryAsync<CompanyBankAccountModel>(sql, new { CompanyId = companyId })).ToList();
        }

        public async Task<List<ShippingAddressPresetModel>> GetShippingAddressPresetsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID, ShippingAddress FROM ForeignCustomersShippingAddresses WHERE CustCode=@CustCode AND Country=@Country ORDER BY EntryID";
            return (await db.QueryAsync<ShippingAddressPresetModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<InvoiceToPresetModel>> GetInvoiceToPresetsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID, InvoiceTo FROM ForeignCustomers_InvoiceTo WHERE CustCode=@CustCode AND Country=@Country ORDER BY EntryID";
            return (await db.QueryAsync<InvoiceToPresetModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetCustomerItemsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CompItemID as DropDownValue_ID, '{' + CompItemID + '} {' + ItemID + '} ' + ISNULL(CompDesc,'') as DropDownValue_Description
                FROM VFCustomerCatalog 
                WHERE CustCode=@CustCode AND Country=@Country
                ORDER BY ItemID";
            return (await db.QueryAsync<GenericDropDownModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetProformasByItemAsync(string itemCode, string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CAST(EntryID AS VARCHAR) as DropDownValue_ID, 
                       PInvoice + ' {' + OrderNo + '}-{' + CAST(InternalRefNo AS VARCHAR) + '} {' + CAST(InvQty AS VARCHAR) + '}' as DropDownValue_Description
                FROM VFProformas 
                WHERE ItemID=@ItemCode AND CustCode=@CustCode AND Country=@Country AND InvQty > 0
                ORDER BY PInvoice";
            return (await db.QueryAsync<GenericDropDownModel>(sql, new { ItemCode = itemCode, CustCode = custCode, Country = country })).ToList();
        }



        public async Task<List<CityModel>> GetCitiesAsync(string countryName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CityID, City AS CityName FROM Cities WHERE CountryName = @CountryName ORDER BY City";
            return (await db.QueryAsync<CityModel>(sql, new { CountryName = countryName })).ToList();
        }

        public async Task<List<PortModel>> GetPortsAsync(int cityId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT PortID, Port, CityID FROM Ports WHERE CityID = @CityID ORDER BY Port";
            return (await db.QueryAsync<PortModel>(sql, new { CityID = cityId })).ToList();
        }

        public async Task<(int CityID, string CountryName)> GetPortCityAndCountryAsync(int portId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT p.CityID, c.CountryName FROM Ports p INNER JOIN Cities c ON p.CityID = c.CityID WHERE p.PortID = @PortID";
            var row = await db.QueryFirstOrDefaultAsync(sql, new { PortID = portId });
            if (row == null) return (0, string.Empty);
            return (row.CityID, row.CountryName ?? string.Empty);
        }
        
        public async Task<List<AdvancePaymentModel>> GetAdvancePaymentsAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT EntryID, AdviceNo, Amount, AdjustedAmount
                FROM VFCustAdvancePayments
                WHERE CustCode=@CustCode AND Amount > AdjustedAmount
                ORDER BY EntryID";
            return (await db.QueryAsync<AdvancePaymentModel>(sql, new { CustCode = custCode })).ToList();
        }

        public async Task<List<string>> GetProformasForInvoiceAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT DISTINCT PInvoice 
                FROM VFProformas 
                WHERE CustCode = @CustCode AND Country = @Country AND InvQty > 0
                ORDER BY PInvoice";
            return (await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<CustomPListModel>> GetItemsForProformaAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    EntryID AS OrderItemID,
                    PInvoice AS ProformaInvoice,
                    OrderNo,
                    ItemID AS ItemCode,
                    ItemName AS Description,
                    Price,
                    InvQty AS Qty
                FROM VFProformas
                WHERE PInvoice = @PInvoice AND InvQty > 0
                ORDER BY CompItemCode";
            // Map the result to CustomPListModel for packing list creation
            var items = await db.QueryAsync<CustomPListModel>(sql, new { PInvoice = pInvoice });
            return items.ToList();
        }

        public async Task<List<DispatchListLookupModel>> GetAvailableDispatchListsAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT EntryID, DispatchListNo + '{' + CAST(EntryID AS VARCHAR(10)) + '}' AS Description
                FROM DispatchList 
                WHERE CustCode = @CustCode AND Finalyzed = 1 
                  AND EntryID NOT IN(SELECT DP_RefID FROM CustomPList_DispatchListDetail)";
            var result = await db.QueryAsync<DispatchListLookupModel>(sql, new { CustCode = custCode });
            return result.ToList();
        }

        public async Task<List<DispatchCartonModel>> GetDispatchListCartonsAsync(long dispatchListEntryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    DispatchListEntryID,
                    CartonNo,
                    OrderNo,
                    ItemCode,
                    LotNo,
                    SUM(Qty) AS CartonQty 
                FROM VDispatchListDetail_Ex_OrderWise 
                WHERE DispatchListEntryID = @EntryId 
                GROUP BY DispatchListEntryID, CartonNo, OrderNo, ItemCode, LotNo 
                ORDER BY CartonNo";
            var result = await db.QueryAsync<DispatchCartonModel>(sql, new { EntryId = dispatchListEntryId });
            return result.ToList();
        }

        public async Task<List<CustomPListModel>> GetProformaItemsForDispatchMappingAsync(string orderNo, string itemCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT EntryID AS OrderItemID, OrderNo, PInvoice AS ProformaInvoice, ItemID AS ItemCode, 
                       ItemName AS Description, Qty, InvQty AS Qty, Price, CustomPrice 
                FROM VFProformas 
                WHERE OrderNo = @OrderNo AND ItemID = @ItemCode AND InvQty > 0";
            var result = await db.QueryAsync<CustomPListModel>(sql, new { OrderNo = orderNo, ItemCode = itemCode });
            return result.ToList();
        }

        public async Task<CustomInvoiceViewModel?> GetProformaHeaderAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT TOP 1 TradeTerms, PartialShipment, TransShipment, PaymentTerms, packaging AS Packaging, 
                             SMethod as ShippingMode, DeliveryDT as CustomDT, InvoiceOf, ShippingAgent as ShippingAgentID, 
                             OurBankID, LoadPort as LoadPortID, DischargePort as DisPortID
                FROM FPInvoice
                WHERE PInvoice = @pInvoice";
            return await db.QueryFirstOrDefaultAsync<CustomInvoiceViewModel>(sql, new { pInvoice });
        }

        public async Task<List<CustomInvoiceOtherChargeModel>> GetProformaOtherChargesAsync(string pInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT SNo, ChargesDetail, ChargesAmt
                FROM FPInvoiceOtherCharges
                WHERE PInvoice = @pInvoice";
            var result = await db.QueryAsync<CustomInvoiceOtherChargeModel>(sql, new { pInvoice });
            return result.ToList();
        }

        // ──────────────────────────────────────────────────────────────────────
        // INVOICE # GENERATION
        // ──────────────────────────────────────────────────────────────────────

        public async Task<string> GetNextCustomInvoiceNumberAsync(string strCompany)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string prefix = "E-";
            if (strCompany == "Tecno")
                prefix = "";

            const string sql = @"
                SELECT ISNULL(MAX(CAST(SUBSTRING(CustomInvoice,5,LEN(CustomInvoice)-4) AS INT)), 0) + 1 
                FROM CustomInvoice 
                WHERE LEFT(CustomInvoice,2)='E-' AND LEN(CustomInvoice)>=7 AND LEN(REPLACE(CustomInvoice,'-',''))=(LEN(CustomInvoice)-1)";
            
            long nextNo = await db.QueryFirstOrDefaultAsync<long>(sql);
            return $"{prefix}{nextNo}"; // Wait, if prefix is E-, should it be E-123 ? legacy said: E-123? We just return prefix + nextNo. Wait, the legacy format was E-123.
        }

        // ──────────────────────────────────────────────────────────────────────
        // LOAD INVOICE
        // ──────────────────────────────────────────────────────────────────────
        public async Task<CustomInvoiceViewModel?> GetCustomInvoiceAsync(string customInvoiceNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            const string sqlHeader = "SELECT * FROM CustomInvoice WHERE CustomInvoice = @CustomInvoiceNo";
            var headerRow = await db.QueryFirstOrDefaultAsync(sqlHeader, new { CustomInvoiceNo = customInvoiceNo });
            
            if (headerRow == null) return null;

            var model = new CustomInvoiceViewModel
            {
                CustomInvoiceNo = headerRow.CustomInvoice,
                DT = headerRow.DT ?? DateTime.Today,
                CustCode = headerRow.CustCode ?? "",
                Country = headerRow.Country ?? "",
                Currency = headerRow.CustCurrency ?? "",
                ShippingMode = headerRow.ShippingMode ?? "",
                TradeTerms = headerRow.TradeTerms ?? "",
                PaymentTerms = headerRow.PaymentTerms ?? "",
                LCNo = headerRow.LCNo ?? "",
                LCDate = headerRow.LCDate,
                PartialShipment = headerRow.PartialShipment ?? "",
                TransShipment = headerRow.TransShipment ?? "",
                Packaging = headerRow.Packaging ?? "",
                NoSRO = headerRow.NoSRO ?? false,
                CIFTradeTerms = headerRow.CIFTradeTerms ?? false,
                SampleInvoice = headerRow.SampleInvoice ?? false,
                FreeSamples = headerRow.FreeSamples ?? false,
                GDNo = headerRow.GDNo ?? "",
                CustomDT = headerRow.CustomDT,
                LoadPortID = headerRow.LoadPortID,
                DisPortID = headerRow.DisPortID,
                OurBankID = headerRow.OurBankID,
                ShippingAgentID = headerRow.ShippingAgentID,
                InsuranceAgentID = headerRow.InsuranceAgentID,
                CompanyRefID = headerRow.CompanyRefID ?? 0,
                AWBNo = headerRow.AWBNo ?? "",
                AWBNoDT = headerRow.AWBNoDT,
                FormE = headerRow.FormE ?? "",
                FormEDT = headerRow.FormEDT,
                Vessel = headerRow.Vessel ?? "",
                GrossWeight = headerRow.GrossWeight ?? 0,
                DeclarationsBalls = headerRow.DeclarationsBalls ?? "",
                DeclarationsGarments = headerRow.DeclarationsGarments ?? "",
                DeclarationsGloves = headerRow.DeclarationsGloves ?? "",
                Consignee = headerRow.Consignee ?? "",
                ComDeclarationsBalls = headerRow.ComDeclarationsBalls ?? "",
                ComDeclarationsGarments = headerRow.ComDeclarationsGarments ?? "",
                ComDeclarationsGloves = headerRow.ComDeclarationsGloves ?? "",
                ComNotify = headerRow.ComNotify ?? "",
                ComConsignee = headerRow.ComConsignee ?? "",
                ComPayTerms = headerRow.ComPayTerms ?? "",
                ComAgainst = headerRow.ComAgainst ?? "",
                ComDrawnUnder = headerRow.ComDrawnUnder ?? "",
                ComSpecial = headerRow.ComSpecial ?? "",
                ManDischarge = headerRow.ManDischarge ?? "",
                ManDeclaration = headerRow.ManDeclaration ?? "",
                Warehouse = headerRow.Warehouse ?? "",
                Gate = headerRow.Gate ?? "",
                ContainerNo = headerRow.ContainerNo ?? "",
                SealNo = headerRow.SealNo ?? "",
                DTETA = headerRow.DTETA,
                SRONo = headerRow.SRONo ?? "",
                InvoiceOf = headerRow.InvoiceOf ?? "",
                GatePassNo = headerRow.GatePassNo ?? "",
                GatePassDT = headerRow.GatePassDT,
                DTREDescription = headerRow.DTREDescription ?? "",
                InvoiceDescription = headerRow.InvoiceDescription ?? "",
                InvoiceTo = headerRow.InvoiceTo ?? "",
                TotalCartons = headerRow.TotalCartons ?? 0
            };

            const string sqlItems = @"
                SELECT 
                    i.EntryID,
                    i.RefID,
                    o.PInvoice AS ProformaInvoice,
                    o.OrderNo,
                    o.ItemID AS ItemCode,
                    o.ItemName AS Description,
                    o.Unit,
                    i.Qty,
                    i.Price,
                    i.CustomPrice,
                    i.RebateRate,
                    i.DTRENo,
                    i.ItemForLabTest,
                    i.BatchNo,
                    i.CII_LotNo AS CIILotNo
                FROM CustomInvoiceItems i
                INNER JOIN VFProformas o ON i.RefID = o.EntryID
                WHERE i.CustomInvoice = @CustomInvoiceNo";
            
            model.LineItems = (await db.QueryAsync<CustomInvoiceItemModel>(sqlItems, new { CustomInvoiceNo = customInvoiceNo })).ToList();

            const string sqlPList = @"
                SELECT 
                    p.ID,
                    p.OrderItemID,
                    p.CartonFrom,
                    p.CartonTo,
                    p.Qty,
                    p.LotNo,
                    p.BatchNo,
                    o.OrderNo,
                    o.PInvoice AS ProformaInvoice,
                    o.ItemID AS ItemCode,
                    o.ItemName AS Description
                FROM CustomPList p
                INNER JOIN CustomInvoiceItems i ON p.OrderItemID = i.EntryID
                INNER JOIN VFProformas o ON i.RefID = o.EntryID
                WHERE p.CustomInvoice = @CustomInvoiceNo";
            
            model.PackingListItems = (await db.QueryAsync<CustomPListModel>(sqlPList, new { CustomInvoiceNo = customInvoiceNo })).ToList();

            const string sqlCharges = @"
                SELECT EntryID, SNo, ChargesDetail, ChargesAmt, Custom 
                FROM CustomInvoiceOtherCharges 
                WHERE CustomInvoice = @CustomInvoiceNo
                ORDER BY SNo";
            model.OtherCharges = (await db.QueryAsync<CustomInvoiceOtherChargeModel>(sqlCharges, new { CustomInvoiceNo = customInvoiceNo })).ToList();

            const string sqlAdvances = @"
                SELECT a.EntryID, v.AdviceNo, v.Amount, v.AdjustedAmount, a.Amount AS AdjustNow
                FROM FCustAdvancePaymentsAdjusted a
                INNER JOIN VFCustAdvancePayments v ON a.RefID = v.EntryID
                WHERE a.CustomInvoice = @CustomInvoiceNo";
            model.AdvancePayments = (await db.QueryAsync<AdvancePaymentModel>(sqlAdvances, new { CustomInvoiceNo = customInvoiceNo })).ToList();

            return model;
        }

        // ──────────────────────────────────────────────────────────────────────
        // SAVE INVOICE
        // ──────────────────────────────────────────────────────────────────────
        public async Task<bool> SaveCustomInvoiceAsync(CustomInvoiceViewModel dto)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var tr = db.BeginTransaction();

            try
            {
                var existing = await db.QueryFirstOrDefaultAsync<string>(
                    "SELECT CustomInvoice FROM CustomInvoice WHERE CustomInvoice = @No", 
                    new { No = dto.CustomInvoiceNo }, tr);

                if (existing == null)
                {
                    const string sqlInsert = @"
                        INSERT INTO CustomInvoice (
                            CustomInvoice, DT, CustCode, Country, ShippingMode,
                            LoadPortID, DisPortID, TradeTerms, PaymentTerms, LCNo, PartialShipment,
                            TransShipment, ShippingAgentID, InsuranceAgentID, OurBankID,
                            AWBNo, AWBNoDT, FormE, FormEDT, Vessel, GrossWeight, DeclarationsBalls, DeclarationsGarments,
                            DeclarationsGloves, Consignee, ComDeclarationsBalls, ComDeclarationsGarments,
                            ComDeclarationsGloves, ComNotify, ComConsignee, ComPayTerms, ComAgainst, ComDrawnUnder,
                            ComSpecial, ManDischarge, ManDeclaration, Warehouse, Gate, ContainerNo, DTETA, SRONo, CustCurrency, CompanyRefID, SealNo, InvoiceOf, LCDate, GatePassNo, GatePassDT, DTREDescription, NoSRO, CIFTradeTerms, SampleInvoice, FreeSamples, InvoiceDescription, CustomDT, GDNo, InvoiceTo
                        ) VALUES (
                            @CustomInvoiceNo, @DT, @CustCode, @Country, @ShippingMode,
                            @LoadPortID, @DisPortID, @TradeTerms, @PaymentTerms, @LCNo, @PartialShipment,
                            @TransShipment, @ShippingAgentID, @InsuranceAgentID, @OurBankID,
                            @AWBNo, @AWBNoDT, @FormE, @FormEDT, @Vessel, @GrossWeight, @DeclarationsBalls, @DeclarationsGarments,
                            @DeclarationsGloves, @Consignee, @ComDeclarationsBalls, @ComDeclarationsGarments,
                            @ComDeclarationsGloves, @ComNotify, @ComConsignee, @ComPayTerms, @ComAgainst, @ComDrawnUnder,
                            @ComSpecial, @ManDischarge, @ManDeclaration, @Warehouse, @Gate, @ContainerNo, @DTETA, @SRONo, @Currency, @CompanyRefID, @SealNo, @InvoiceOf, @LCDate, @GatePassNo, @GatePassDT, @DTREDescription, @NoSRO, @CIFTradeTerms, @SampleInvoice, @FreeSamples, @InvoiceDescription, @CustomDT, @GDNo, @InvoiceTo
                        )";
                    await db.ExecuteAsync(sqlInsert, dto, tr);
                }
                else
                {
                    const string sqlUpdate = @"
                        UPDATE CustomInvoice SET
                            DT = @DT, CustCode = @CustCode, Country = @Country, ShippingMode = @ShippingMode,
                            LoadPortID = @LoadPortID, DisPortID = @DisPortID, TradeTerms = @TradeTerms, PaymentTerms = @PaymentTerms, LCNo = @LCNo, PartialShipment = @PartialShipment,
                            TransShipment = @TransShipment, ShippingAgentID = @ShippingAgentID, InsuranceAgentID = @InsuranceAgentID, OurBankID = @OurBankID,
                            AWBNo = @AWBNo, AWBNoDT = @AWBNoDT, FormE = @FormE, FormEDT = @FormEDT, Vessel = @Vessel, GrossWeight = @GrossWeight, DeclarationsBalls = @DeclarationsBalls, DeclarationsGarments = @DeclarationsGarments,
                            DeclarationsGloves = @DeclarationsGloves, Consignee = @Consignee, ComDeclarationsBalls = @ComDeclarationsBalls, ComDeclarationsGarments = @ComDeclarationsGarments,
                            ComDeclarationsGloves = @ComDeclarationsGloves, ComNotify = @ComNotify, ComConsignee = @ComConsignee, ComPayTerms = @ComPayTerms, ComAgainst = @ComAgainst, ComDrawnUnder = @ComDrawnUnder,
                            ComSpecial = @ComSpecial, ManDischarge = @ManDischarge, ManDeclaration = @ManDeclaration, Warehouse = @Warehouse, Gate = @Gate, ContainerNo = @ContainerNo, DTETA = @DTETA, SRONo = @SRONo, CustCurrency = @Currency, CompanyRefID = @CompanyRefID, SealNo = @SealNo, InvoiceOf = @InvoiceOf, LCDate = @LCDate, GatePassNo = @GatePassNo, GatePassDT = @GatePassDT, DTREDescription = @DTREDescription, NoSRO = @NoSRO, CIFTradeTerms = @CIFTradeTerms, SampleInvoice = @SampleInvoice, FreeSamples = @FreeSamples, InvoiceDescription = @InvoiceDescription, CustomDT = @CustomDT, GDNo = @GDNo, InvoiceTo = @InvoiceTo
                        WHERE CustomInvoice = @CustomInvoiceNo";
                    await db.ExecuteAsync(sqlUpdate, dto, tr);
                }

                // Delete removed items
                if (dto.DeletedLineItemIds.Any())
                {
                    await db.ExecuteAsync("DELETE FROM CustomInvoiceItems WHERE EntryID IN @Ids", new { Ids = dto.DeletedLineItemIds }, tr);
                }
                if (dto.DeletedPackingListIds.Any())
                {
                    await db.ExecuteAsync("DELETE FROM CustomPList WHERE ID IN @Ids", new { Ids = dto.DeletedPackingListIds }, tr);
                }
                if (dto.DeletedOtherChargeIds.Any())
                {
                    await db.ExecuteAsync("DELETE FROM CustomInvoiceOtherCharges WHERE EntryID IN @Ids", new { Ids = dto.DeletedOtherChargeIds }, tr);
                }
                // Also delete advance payments if needed
                await db.ExecuteAsync("DELETE FROM FCustAdvancePaymentsAdjusted WHERE CustomInvoice = @CustomInvoiceNo", new { dto.CustomInvoiceNo }, tr);

                // Upsert Line Items
                foreach (var item in dto.LineItems)
                {
                    if (item.EntryID == 0)
                    {
                        var sql = @"INSERT INTO CustomInvoiceItems (CustomInvoice, RefID, Qty, Price, CustomPrice, RebateRate, DTRENo, ItemForLabTest, BatchNo, CII_LotNo)
                                    VALUES (@CustomInvoice, @RefID, @Qty, @Price, @CustomPrice, @RebateRate, @DTRENo, @ItemForLabTest, @BatchNo, @CIILotNo);
                                    SELECT CAST(SCOPE_IDENTITY() as int);";
                        item.EntryID = await db.QuerySingleAsync<int>(sql, new { 
                            CustomInvoice = dto.CustomInvoiceNo,
                            item.RefID, item.Qty, item.Price, item.CustomPrice, item.RebateRate, item.DTRENo, item.ItemForLabTest, item.BatchNo, CIILotNo = item.CIILotNo
                        }, tr);
                    }
                    else
                    {
                        var sql = @"UPDATE CustomInvoiceItems SET 
                                        Qty = @Qty, Price = @Price, CustomPrice = @CustomPrice, 
                                        RebateRate = @RebateRate, DTRENo = @DTRENo, 
                                        ItemForLabTest = @ItemForLabTest, BatchNo = @BatchNo, CII_LotNo = @CIILotNo
                                    WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(sql, new { 
                            item.Qty, item.Price, item.CustomPrice, item.RebateRate, item.DTRENo, 
                            item.ItemForLabTest, item.BatchNo, CIILotNo = item.CIILotNo, item.EntryID
                        }, tr);
                    }
                }

                // Upsert Packing List
                foreach (var pl in dto.PackingListItems)
                {
                    if (pl.ID == 0)
                    {
                        var sql = @"INSERT INTO CustomPList (CustomInvoice, OrderItemID, CartonFrom, CartonTo, Qty, LotNo, BatchNo)
                                    VALUES (@CustomInvoice, @OrderItemID, @CartonFrom, @CartonTo, @Qty, @LotNo, @BatchNo);
                                    SELECT CAST(SCOPE_IDENTITY() as int);";
                        pl.ID = await db.QuerySingleAsync<int>(sql, new { 
                            CustomInvoice = dto.CustomInvoiceNo,
                            pl.OrderItemID, pl.CartonFrom, pl.CartonTo, pl.Qty, pl.LotNo, pl.BatchNo
                        }, tr);

                        if (pl.DP_RefID > 0)
                        {
                            var dispatchSql = @"INSERT INTO CustomPList_DispatchListDetail (CustomPlist_RefID, ItemCode, CartonNo, DP_RefID, FromOrderNo) 
                                                VALUES (@ID, @ItemCode, @CartonNo, @DP_RefID, @OrderNo)";
                            await db.ExecuteAsync(dispatchSql, new {
                                pl.ID, pl.ItemCode, CartonNo = pl.CartonFrom, pl.DP_RefID, pl.OrderNo
                            }, tr);
                        }
                    }
                    else
                    {
                        var sql = @"UPDATE CustomPList SET 
                                        CartonFrom = @CartonFrom, CartonTo = @CartonTo, 
                                        Qty = @Qty, LotNo = @LotNo, BatchNo = @BatchNo
                                    WHERE ID = @ID";
                        await db.ExecuteAsync(sql, pl, tr);
                    }
                }

                // Upsert Other Charges
                foreach (var oc in dto.OtherCharges)
                {
                    if (oc.EntryID == 0)
                    {
                        var sql = @"INSERT INTO CustomInvoiceOtherCharges (CustomInvoice, ChargesDetail, ChargesAmt, SNo, Custom)
                                    VALUES (@CustomInvoice, @ChargesDetail, @ChargesAmt, @SNo, @Custom)";
                        await db.ExecuteAsync(sql, new { 
                            CustomInvoice = dto.CustomInvoiceNo,
                            oc.ChargesDetail, oc.ChargesAmt, oc.SNo, oc.Custom
                        }, tr);
                    }
                    else
                    {
                        var sql = @"UPDATE CustomInvoiceOtherCharges SET 
                                        ChargesDetail = @ChargesDetail, ChargesAmt = @ChargesAmt, 
                                        SNo = @SNo, Custom = @Custom
                                    WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(sql, oc, tr);
                    }
                }

                // Insert Advances (deleted entirely and re-inserted)
                foreach (var adv in dto.AdvancePayments)
                {
                    if (adv.AdjustNow > 0)
                    {
                        var sql = @"INSERT INTO FCustAdvancePaymentsAdjusted (CustomInvoice, RefID, Amount)
                                    VALUES (@CustomInvoice, @RefID, @Amount)";
                        await db.ExecuteAsync(sql, new { 
                            CustomInvoice = dto.CustomInvoiceNo,
                            RefID = adv.EntryID,
                            Amount = adv.AdjustNow
                        }, tr);
                    }
                }

                // Legacy update procedures
                // Wait, in legacy, it updates CustomInvoiceItems Qty by summing CustomPList
                await db.ExecuteAsync(@"
                    UPDATE CustomInvoiceItems 
                    SET CustomInvoiceItems.Qty = T1.Qty 
                    FROM CustomInvoiceItems 
                    INNER JOIN (SELECT OrderItemID, SUM(((CartonTo-CartonFrom)+1)*Qty) AS Qty FROM CustomPList WHERE CustomInvoice=@CustomInvoiceNo GROUP BY OrderItemID) T1 ON CustomInvoiceItems.EntryID = T1.OrderItemID
                    WHERE CustomInvoiceItems.CustomInvoice = @CustomInvoiceNo", 
                new { CustomInvoiceNo = dto.CustomInvoiceNo }, tr);

                // Proforma update procedure
                await db.ExecuteAsync(@"
                    UPDATE FProformaOrders 
                    SET InvQty = InvQty - IsNull((SELECT Qty FROM CustomInvoiceItems WHERE CustomInvoice=@CustomInvoiceNo AND RefID=FProformaOrders.EntryID), 0)
                    WHERE EntryID IN (SELECT RefID FROM CustomInvoiceItems WHERE CustomInvoice=@CustomInvoiceNo)",
                new { CustomInvoiceNo = dto.CustomInvoiceNo }, tr);

                tr.Commit();
                return true;
            }
            catch (Exception)
            {
                tr.Rollback();
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // DELETE INVOICE
        // ──────────────────────────────────────────────────────────────────────
        public async Task<bool> DeleteCustomInvoiceAsync(string customInvoiceNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var tr = db.BeginTransaction();
            try
            {
                // Restore Proforma quantities before deleting
                await db.ExecuteAsync(@"
                    UPDATE FProformaOrders 
                    SET InvQty = InvQty + IsNull((SELECT Qty FROM CustomInvoiceItems WHERE CustomInvoice=@CustomInvoiceNo AND RefID=FProformaOrders.EntryID), 0)
                    WHERE EntryID IN (SELECT RefID FROM CustomInvoiceItems WHERE CustomInvoice=@CustomInvoiceNo)",
                new { CustomInvoiceNo = customInvoiceNo }, tr);

                await db.ExecuteAsync("DELETE FROM CustomInvoice WHERE CustomInvoice = @CustomInvoiceNo", new { CustomInvoiceNo = customInvoiceNo }, tr);
                tr.Commit();
                return true;
            }
            catch (Exception)
            {
                tr.Rollback();
                throw;
            }
        }
    }
}
