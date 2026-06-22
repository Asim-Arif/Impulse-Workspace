using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace DataAccessLibrary.DAC.Export
{
    public class ForeignCustomerDataAccess : IForeignCustomerDataAccess
    {
        private readonly string _connectionString;

        public ForeignCustomerDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        // ------------------------------------------------------------------ //
        //  READ
        // ------------------------------------------------------------------ //

        public async Task<ForeignCustomerViewModel?> GetCustomerAsync(string custId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CustCode,Name,City,Country,Curr,
                       Phone1,Phone2,Phone3,Fax1,Fax2,Fax3,EMail1,EMail2,URL,
                       Address,FakeAddress,PaymentTerms,TradeTerms,DTFormat,
                       FinishingQuality,DefaultShipMethod,CustomerSource,
                       DefaultPort,OuterPackingLabel,InnerPackingLabel,
                       Inner_Label_Manual_I,Inner_Label_Manual_II,
                       Stamps,RcvdVia,SpecialInstructions,FC_Note_I,FC_Note_II,
                       PaymentDays,MaxDiscount,
                       AcceptsExtraQty,LateOrderAlerts,ShowCustomerRef,Active,
                       Cont1Name,Cont1Designation,Cont1Email,cont1Phone,Cont1Mobile,Cont1Skype,
                       Cont2Name,Cont2Designation,Cont2Email,Cont2Phone,Cont2Mobile,Cont2Skype,
                       NearestAirport,NearestRailwaystation,FedexNo,DHLNo,
                       AgentName,AgentEmail,AgentPhone1,AgentPhone2,AgentFax1,AgentFax2,
                       AgentAddress,AgentURL,ACName,ACEmail,ACPhone,ACMobile,
                       AccNo
                FROM VFCustomers
                WHERE CustCode+Country = @CustID";

            var model = await db.QueryFirstOrDefaultAsync<ForeignCustomerViewModel>(sql,
                new { CustID = custId });

            if (model == null) return null;

            // Eager-load child collections
            model.Banks = (await db.QueryAsync<ForeignCustomerBankViewModel>(@"
                SELECT BankID,BankName,BankCountry AS Country,Phone1,Phone2,
                       Fax1,Fax2,Email1,Email2,Address1,Address2,DefaultBank
                FROM FCustBanks
                WHERE CustCode+Country = @CustID", new { CustID = custId })).ToList();

            model.ShippingAddresses = (await db.QueryAsync<ForeignCustomerAddressViewModel>(@"
                SELECT EntryID, ShippingAddress
                FROM ForeignCustomersShippingAddresses
                WHERE CustCode+Country = @CustID", new { CustID = custId })).ToList();

            model.Events = (await db.QueryAsync<ForeignCustomerEventViewModel>(@"
                SELECT EntryID, Description, EventDT, RepeatedType, RepeatDescription
                FROM VForeignCustomers_Events
                WHERE CustCode+Country = @CustID", new { CustID = custId })).ToList();

            model.InvoiceTo = (await db.QueryAsync<ForeignCustomerInvoiceToViewModel>(@"
                SELECT EntryID, InvoiceTo
                FROM ForeignCustomers_InvoiceTo
                WHERE CustCode+Country = @CustID", new { CustID = custId })).ToList();

            return model;
        }

        // ------------------------------------------------------------------ //
        //  WRITE — Customer
        // ------------------------------------------------------------------ //

        public async Task<string> InsertCustomerAsync(ForeignCustomerViewModel m, string userName, string computerName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                INSERT INTO ForeignCustomers
                (CustCode,Name,City,Country,Curr,
                 Phone1,Phone2,Phone3,Fax1,Fax2,Fax3,EMail1,EMail2,URL,
                 Address,FakeAddress,PaymentTerms,TradeTerms,DTFormat,
                 FinishingQuality,DefaultShipMethod,CustomerSource,
                 DefaultPort,OuterPackingLabel,InnerPackingLabel,
                 Inner_Label_Manual_I,Inner_Label_Manual_II,
                 Stamps,RcvdVia,SpecialInstructions,FC_Note_I,FC_Note_II,
                 PaymentDays,MaxDiscount,
                 AcceptsExtraQty,LateOrderAlerts,ShowCustomerRef,Active,
                 Cont1Name,Cont1Designation,Cont1Email,cont1Phone,Cont1Mobile,Cont1Skype,
                 Cont2Name,Cont2Designation,Cont2Email,Cont2Phone,Cont2Mobile,Cont2Skype,
                 NearestAirport,NearestRailwaystation,FedexNo,DHLNo,
                 AgentName,AgentEmail,AgentPhone1,AgentPhone2,AgentFax1,AgentFax2,
                 AgentAddress,AgentURL,ACName,ACEmail,ACPhone,ACMobile,
                 AccNo)
                VALUES
                (@CustCode,@Name,@City,@Country,@Curr,
                 @Phone1,@Phone2,@Phone3,@Fax1,@Fax2,@Fax3,@EMail1,@EMail2,@URL,
                 @Address,@FakeAddress,@PaymentTerms,@TradeTerms,@DTFormat,
                 @FinishingQuality,@DefaultShipMethod,@CustomerSource,
                 @DefaultPort,@OuterPackingLabel,@InnerPackingLabel,
                 @Inner_Label_Manual_I,@Inner_Label_Manual_II,
                 @Stamps,@RcvdVia,@SpecialInstructions,@FC_Note_I,@FC_Note_II,
                 @PaymentDays,@MaxDiscount,
                 @AcceptsExtraQty,@LateOrderAlerts,@ShowCustomerRef,@Active,
                 @Cont1Name,@Cont1Designation,@Cont1Email,@cont1Phone,@Cont1Mobile,@Cont1Skype,
                 @Cont2Name,@Cont2Designation,@Cont2Email,@Cont2Phone,@Cont2Mobile,@Cont2Skype,
                 @NearestAirport,@NearestRailwaystation,@FedexNo,@DHLNo,
                 @AgentName,@AgentEmail,@AgentPhone1,@AgentPhone2,@AgentFax1,@AgentFax2,
                 @AgentAddress,@AgentURL,@ACName,@ACEmail,@ACPhone,@ACMobile,
                 @AccNo)";

            await db.ExecuteAsync(sql, m);
            return m.CustCode + m.Country;
        }

        public async Task UpdateCustomerAsync(ForeignCustomerViewModel m, string custId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                UPDATE ForeignCustomers SET
                 Name=@Name,City=@City,Country=@Country,Curr=@Curr,
                 Phone1=@Phone1,Phone2=@Phone2,Phone3=@Phone3,
                 Fax1=@Fax1,Fax2=@Fax2,Fax3=@Fax3,
                 EMail1=@EMail1,EMail2=@EMail2,URL=@URL,
                 Address=@Address,FakeAddress=@FakeAddress,
                 PaymentTerms=@PaymentTerms,TradeTerms=@TradeTerms,DTFormat=@DTFormat,
                 FinishingQuality=@FinishingQuality,DefaultShipMethod=@DefaultShipMethod,
                 CustomerSource=@CustomerSource,
                 DefaultPort=@DefaultPort,OuterPackingLabel=@OuterPackingLabel,InnerPackingLabel=@InnerPackingLabel,
                 Inner_Label_Manual_I=@Inner_Label_Manual_I,Inner_Label_Manual_II=@Inner_Label_Manual_II,
                 Stamps=@Stamps,RcvdVia=@RcvdVia,
                 SpecialInstructions=@SpecialInstructions,FC_Note_I=@FC_Note_I,FC_Note_II=@FC_Note_II,
                 PaymentDays=@PaymentDays,MaxDiscount=@MaxDiscount,
                 AcceptsExtraQty=@AcceptsExtraQty,LateOrderAlerts=@LateOrderAlerts,
                 ShowCustomerRef=@ShowCustomerRef,Active=@Active,
                 Cont1Name=@Cont1Name,Cont1Designation=@Cont1Designation,
                 Cont1Email=@Cont1Email,cont1Phone=@cont1Phone,
                 Cont1Mobile=@Cont1Mobile,Cont1Skype=@Cont1Skype,
                 Cont2Name=@Cont2Name,Cont2Designation=@Cont2Designation,
                 Cont2Email=@Cont2Email,Cont2Phone=@Cont2Phone,
                 Cont2Mobile=@Cont2Mobile,Cont2Skype=@Cont2Skype,
                 NearestAirport=@NearestAirport,NearestRailwaystation=@NearestRailwaystation,
                 FedexNo=@FedexNo,DHLNo=@DHLNo,
                 AgentName=@AgentName,AgentEmail=@AgentEmail,
                 AgentPhone1=@AgentPhone1,AgentPhone2=@AgentPhone2,
                 AgentFax1=@AgentFax1,AgentFax2=@AgentFax2,
                 AgentAddress=@AgentAddress,AgentURL=@AgentURL,
                 ACName=@ACName,ACEmail=@ACEmail,ACPhone=@ACPhone,ACMobile=@ACMobile,
                 AccNo=@AccNo
                WHERE CustCode+Country = @CustId";

            await db.ExecuteAsync(sql, new
            {
                m.Name, m.City, m.Country, m.Curr,
                m.Phone1, m.Phone2, m.Phone3, m.Fax1, m.Fax2, m.Fax3,
                m.EMail1, m.EMail2, m.URL, m.Address, m.FakeAddress,
                m.PaymentTerms, m.TradeTerms, m.DTFormat,
                m.FinishingQuality, m.DefaultShipMethod, m.CustomerSource,
                m.DefaultPort, m.OuterPackingLabel, m.InnerPackingLabel,
                m.Inner_Label_Manual_I, m.Inner_Label_Manual_II,
                m.Stamps, m.RcvdVia, m.SpecialInstructions, m.FC_Note_I, m.FC_Note_II,
                m.PaymentDays, m.MaxDiscount,
                m.AcceptsExtraQty, m.LateOrderAlerts, m.ShowCustomerRef, m.Active,
                m.Cont1Name, m.Cont1Designation, m.Cont1Email, m.cont1Phone, m.Cont1Mobile, m.Cont1Skype,
                m.Cont2Name, m.Cont2Designation, m.Cont2Email, m.Cont2Phone, m.Cont2Mobile, m.Cont2Skype,
                m.NearestAirport, m.NearestRailwaystation, m.FedexNo, m.DHLNo,
                m.AgentName, m.AgentEmail, m.AgentPhone1, m.AgentPhone2, m.AgentFax1, m.AgentFax2,
                m.AgentAddress, m.AgentURL, m.ACName, m.ACEmail, m.ACPhone, m.ACMobile,
                m.AccNo,
                CustId = custId
            });
        }

        // ------------------------------------------------------------------ //
        //  WRITE — Banks
        // ------------------------------------------------------------------ //

        public async Task SaveBanksAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedBankIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            if (removedBankIds.Any())
            {
                await db.ExecuteAsync(
                    "DELETE FROM FCustBanks WHERE BankID IN @Ids",
                    new { Ids = removedBankIds });
            }

            foreach (var b in model.Banks)
            {
                if (b.BankID == 0)
                {
                    await db.ExecuteAsync(@"
                        INSERT INTO FCustBanks
                        (CustCode,Country,BankName,BankCountry,Phone1,Phone2,Fax1,Fax2,
                         Email1,Email2,Address1,Address2,DefaultBank)
                        VALUES(@CustCode,@Country,@BankName,@BankCountry,@Phone1,@Phone2,
                               @Fax1,@Fax2,@Email1,@Email2,@Address1,@Address2,@DefaultBank)",
                        new
                        {
                            CustCode = custCode, Country = country,
                            b.BankName, BankCountry = b.Country,
                            b.Phone1, b.Phone2, b.Fax1, b.Fax2,
                            b.Email1, b.Email2, b.Address1, b.Address2, b.DefaultBank
                        });
                }
                else
                {
                    await db.ExecuteAsync(@"
                        UPDATE FCustBanks SET
                         BankName=@BankName,BankCountry=@BankCountry,
                         Phone1=@Phone1,Phone2=@Phone2,Fax1=@Fax1,Fax2=@Fax2,
                         Email1=@Email1,Email2=@Email2,Address1=@Address1,Address2=@Address2,
                         DefaultBank=@DefaultBank
                        WHERE BankID=@BankID",
                        new
                        {
                            b.BankName, BankCountry = b.Country,
                            b.Phone1, b.Phone2, b.Fax1, b.Fax2,
                            b.Email1, b.Email2, b.Address1, b.Address2, b.DefaultBank,
                            b.BankID
                        });
                }
            }
        }

        // ------------------------------------------------------------------ //
        //  WRITE — Shipping Addresses
        // ------------------------------------------------------------------ //

        public async Task SaveShippingAddressesAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            if (removedIds.Any())
            {
                await db.ExecuteAsync(
                    "DELETE FROM ForeignCustomersShippingAddresses WHERE EntryID IN @Ids",
                    new { Ids = removedIds });
            }

            foreach (var a in model.ShippingAddresses)
            {
                if (a.EntryID == 0)
                {
                    await db.ExecuteAsync(@"
                        INSERT INTO ForeignCustomersShippingAddresses(CustCode,Country,ShippingAddress)
                        VALUES(@CustCode,@Country,@ShippingAddress)",
                        new { CustCode = custCode, Country = country, a.ShippingAddress });
                }
                else
                {
                    await db.ExecuteAsync(@"
                        UPDATE ForeignCustomersShippingAddresses
                        SET ShippingAddress=@ShippingAddress
                        WHERE EntryID=@EntryID",
                        new { a.ShippingAddress, a.EntryID });
                }
            }
        }

        // ------------------------------------------------------------------ //
        //  WRITE — Events
        // ------------------------------------------------------------------ //

        public async Task SaveEventsAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            if (removedIds.Any())
            {
                await db.ExecuteAsync(
                    "DELETE FROM ForeignCustomers_Events WHERE EntryID IN @Ids",
                    new { Ids = removedIds });
            }

            foreach (var e in model.Events.Where(x => x.EntryID == 0))
            {
                await db.ExecuteAsync(@"
                    INSERT INTO ForeignCustomers_Events(CustCode,Country,EventDT,RepeatedType,Description)
                    VALUES(@CustCode,@Country,@EventDT,@RepeatedType,@Description)",
                    new { CustCode = custCode, Country = country, e.EventDT, e.RepeatedType, e.Description });
            }
        }

        // ------------------------------------------------------------------ //
        //  WRITE — Invoice To
        // ------------------------------------------------------------------ //

        public async Task SaveInvoiceToAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            if (removedIds.Any())
            {
                await db.ExecuteAsync(
                    "DELETE FROM ForeignCustomers_InvoiceTo WHERE EntryID IN @Ids",
                    new { Ids = removedIds });
            }

            foreach (var it in model.InvoiceTo)
            {
                if (it.EntryID == 0)
                {
                    await db.ExecuteAsync(@"
                        INSERT INTO ForeignCustomers_InvoiceTo(CustCode,Country,InvoiceTo)
                        VALUES(@CustCode,@Country,@InvoiceTo)",
                        new { CustCode = custCode, Country = country, it.InvoiceTo });
                }
                else
                {
                    await db.ExecuteAsync(@"
                        UPDATE ForeignCustomers_InvoiceTo SET InvoiceTo=@InvoiceTo
                        WHERE EntryID=@EntryID",
                        new { it.InvoiceTo, it.EntryID });
                }
            }
        }
    }
}
