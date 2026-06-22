using DataAccessLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface
{
    internal interface IPhotocopyData
    {

    }
}

namespace DataAccessLibrary
{
    public interface IPhotocopyData
    {
        Task<string> GetmaxInvoiceNumberAsync(DateTime CipDT);
        Task<string> GetmaxInvoiceNumberAsync_ForInvoice(DateTime CipDT);
        Task<List<CopyEntry>> GetCopyTypes();
        //Task<List<CopyEntry>> GetPhotoCopyEntriesListAsync(DateTime DT);
        Task InsertPhotoCopyAsync(CopyEntry item);
        Task<List<CopyEntry>> GetPhotoCopyEntriesListAsync(DateTime fromDate, DateTime toDate, int i_bshowprocessed=0);
        
        Task SavePhotoCopyInvoicesAsync(CopyEntry item);            // Method to save the item
    }
}
