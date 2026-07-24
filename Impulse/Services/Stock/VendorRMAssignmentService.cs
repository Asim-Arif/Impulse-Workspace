using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class VendorRMAssignmentService : IVendorRMAssignmentService
    {
        private readonly IVendorRMAssignmentDataAccess _dataAccess;

        public VendorRMAssignmentService(IVendorRMAssignmentDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<NewVendorViewModel>> GetVendorsAsync()
        {
            return await _dataAccess.GetVendorsAsync();
        }

        public async Task<List<UnassignedRMLookupModel>> GetUnassignedRMsAsync(string vendId)
        {
            if (string.IsNullOrWhiteSpace(vendId)) return new List<UnassignedRMLookupModel>();
            return await _dataAccess.GetUnassignedRMsAsync(vendId);
        }

        public async Task<List<VendorRMAssignmentModel>> GetAssignedRMsAsync(string vendId)
        {
            if (string.IsNullOrWhiteSpace(vendId)) return new List<VendorRMAssignmentModel>();
            return await _dataAccess.GetAssignedRMsAsync(vendId);
        }

        public async Task AssignRMAsync(string vendId, int rmId)
        {
            if (string.IsNullOrWhiteSpace(vendId) || rmId <= 0)
                throw new ArgumentException("Invalid Vendor or Raw Material");
            
            await _dataAccess.AssignRMAsync(vendId, rmId);
        }

        public async Task UnassignRMAsync(int entryId)
        {
            if (entryId <= 0) throw new ArgumentException("Invalid assignment entry.");

            bool hasPending = await _dataAccess.HasPendingReceivablesAsync(entryId);
            if (hasPending)
            {
                throw new InvalidOperationException("Can't Unassign this item. It's Receivable.");
            }

            await _dataAccess.UnassignRMAsync(entryId);
        }

        public async Task UpdateRMAssignmentAsync(int entryId, decimal rate, string remarks, string userName)
        {
            if (entryId <= 0) throw new ArgumentException("Invalid assignment entry.");
            await _dataAccess.UpdateRMAssignmentAsync(entryId, rate, remarks, userName);
        }

        public async Task<List<VendorRMAssignmentHistoryModel>> GetHistoryAsync(int entryId)
        {
            return await _dataAccess.GetHistoryAsync(entryId);
        }
    }
}
