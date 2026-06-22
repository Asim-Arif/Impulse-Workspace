using DataAccessLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface
{
    internal interface ICIPData
    {

    }
}

namespace DataAccessLibrary
{
    public interface ICIPData
    {
        Task<List<CipLoungeEntry>> GetCustomersAsync(); // (Optional) Retrieve all Customers
        Task InsertCIPAsync(CipLoungeEntry item);            // Method to save the item
        Task<CipLoungeEntry> GetCIPByInvoiceNosAsync(string invoiceno);        // (Optional) Retrieve item by ID
        Task<IEnumerable<CipLoungeEntry>> GetCIPsAsync();    // (Optional) Retrieve all items
        Task<string> GetmaxCIPInvoiceNumberAsync(DateTime CipDT);
        //Task<List<CipLoungeEntry>> GetCIPEntriesListAsync();
        //Task<List<CipLoungeEntry>> GetCIPEntriesListAsync(DateTime fromDate, DateTime toDate,string sourcetype);
        Task<List<CipLoungeEntry>> GetCIPEntriesListAsync(DateTime fromDate, DateTime toDate, string sourcetype,string searchByField, string txtsearch, string selectsource, string sourcetypevalue);
        Task<string> getBankName(int BankID);
        //Task<IEnumerable<CipLoungeEntry>> GetDirectors();
        Task<List<CipLoungeEntry>> GetDirectors();
        Task<List<CipLoungeEntry>> GetEmployees();
        Task<List<CipLoungeEntry>> GetAirlines();
        Task<List<CipLoungeEntry>> GetSourceTypes();
        Task<List<CipLoungeEntry>> getseats();
        Task<List<CipLoungeEntry>> getflights();
        Task<List<CipLoungeEntry>> getbankcards(int bankid);
        Task<List<CipLoungeEntry>> GetBanks();

        Task<List<CipLoungeEntry>> GetCIPEntriesAsync_cash_credit(DateTime fromDate, DateTime toDate, string sourcetype, bool p_bshowprocessed);
        
        Task<string> GetmaxCIPInvoiceNumberAsync_typewise(DateTime CipDT, string sourcetype);
        Task SaveCIPInvoicesTypewiseAsync(CipLoungeEntry item);            // Method to save the item

        Task<List<CipLoungeEntry>> GetCipDetailsByInvoiceNoAsync(string invoiceNo);
        Task<string> GetSingleStringValueAsync(string field, string table, string cond);
        Task<string> ExecuteMonthlyCIPSummaryAsync(DateTime mydtfrom, DateTime mydtto);
        //Task ImportAndSaveExcelAsync(Stream fileStream);
        //Task<List<CipLoungeEntry>> getbanks();
        Task<List<CIPBanks>> getbanks();
        Task InsertImportedDirector(CipLoungeEntry item);            // Method to save the item
        Task UpdateDirector(int foliono);            // Method to update directors
        Task<List<CipLoungeEntry>> GetCIPEntriesListAsync_typewise(DateTime fromDate, DateTime toDate, string sourcetype);
        
        Task<List<CipLoungeEntry>> GetCIPInvoiceTypewiseDetail(string invoiceNo);
    }
}
