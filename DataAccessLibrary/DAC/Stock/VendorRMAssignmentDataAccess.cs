using Dapper;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendorRMAssignmentDataAccess : IVendorRMAssignmentDataAccess
    {
        private readonly IDBHelper _dbHelper;

        public VendorRMAssignmentDataAccess(IDBHelper dbHelper)
        {
            _dbHelper = dbHelper;
        }

        public async Task<List<NewVendorViewModel>> GetVendorsAsync()
        {
            string sql = "SELECT AccNo, AccTitle as VendorName FROM VVendersANDMakers WHERE Active=1 ORDER BY AccTitle";
            var result = await _dbHelper.getListasync<NewVendorViewModel>(sql);
            return result.ToList();
        }

        public async Task<List<UnassignedRMLookupModel>> GetUnassignedRMsAsync(string vendId)
        {
            string sql = @"
                SELECT RMID, RMName + ' {' + RMID1 + '}' AS RMName 
                FROM RM 
                WHERE RMID NOT IN (SELECT RMID FROM VenderAssItems WHERE VendID=@VendID)  
                ORDER BY RMName";
            var result = await _dbHelper.getListasync<UnassignedRMLookupModel>(sql, new { VendID = vendId });
            return result.ToList();
        }

        public async Task<List<VendorRMAssignmentModel>> GetAssignedRMsAsync(string vendId)
        {
            string sql = @"
                SELECT EntryID, RMID, RMID1, RMName, GroupName, Rate, Remarks 
                FROM VVenderAssItems 
                WHERE VendID=@VendID
                ORDER BY EntryID";
            var result = await _dbHelper.getListasync<VendorRMAssignmentModel>(sql, new { VendID = vendId });
            return result.ToList();
        }

        public async Task AssignRMAsync(string vendId, int rmId)
        {
            string sql = "INSERT INTO VenderAssItems(VendID, RMID) VALUES(@VendID, @RMID)";
            await _dbHelper.ExecuteAsync(sql, new { VendID = vendId, RMID = rmId });
        }

        public async Task<bool> HasPendingReceivablesAsync(int entryId)
        {
            string sql = @"
                SELECT COUNT(1) 
                FROM Vvendissddetail v
                INNER JOIN VenderAssItems a ON v.VendID = a.VendID AND v.Itemcode = CAST(a.RMID as varchar)
                WHERE a.EntryID = @EntryID AND v.Issqty > v.rcvdqty";
            
            var result = await _dbHelper.getListasync<int>(sql, new { EntryID = entryId });
            return result.FirstOrDefault() > 0;
        }

        public async Task UnassignRMAsync(int entryId)
        {
            string sql = "DELETE FROM VenderAssItems WHERE EntryID=@EntryID";
            await _dbHelper.ExecuteAsync(sql, new { EntryID = entryId });
        }

        public async Task UpdateRMAssignmentAsync(int entryId, decimal rate, string remarks, string userName)
        {
            string sql = "UPDATE VenderAssItems SET Rate=@Rate, Remarks=@Remarks WHERE EntryID=@EntryID";
            await _dbHelper.ExecuteAsync(sql, new { EntryID = entryId, Rate = rate, Remarks = remarks ?? string.Empty });

            string historySql = @"
                INSERT INTO VenderAssItems_Revisions(VAI_RefID, Rate, UserName, MachineName, Remarks) 
                VALUES(@EntryID, @Rate, @UserName, @MachineName, @Remarks)";
            
            await _dbHelper.ExecuteAsync(historySql, new { 
                EntryID = entryId, 
                Rate = rate, 
                UserName = userName, 
                MachineName = Environment.MachineName, 
                Remarks = remarks ?? string.Empty 
            });
        }

        public async Task<List<VendorRMAssignmentHistoryModel>> GetHistoryAsync(int entryId)
        {
            string sql = @"
                SELECT EntryID, DTEntry, UserName, Rate, Remarks 
                FROM VenderAssItems_Revisions 
                WHERE VAI_RefID=@EntryID 
                ORDER BY EntryID";
            var result = await _dbHelper.getListasync<VendorRMAssignmentHistoryModel>(sql, new { EntryID = entryId });
            return result.ToList();
        }
    }
}
