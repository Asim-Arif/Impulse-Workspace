using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public class WastageTypeService : IWastageTypeService
    {
        private readonly IWastageTypeDataAccess _dataAccess;
        private readonly IAuditService _auditService;

        public WastageTypeService(IWastageTypeDataAccess dataAccess, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
        }

        public Task<List<WastageTypeModel>> GetWastageTypesAsync() => _dataAccess.GetWastageTypesAsync();

        public Task<WastageTypeDto?> GetWastageTypeByIdAsync(int entryId) => _dataAccess.GetWastageTypeByIdAsync(entryId);

        public async Task<int> SaveWastageTypeAsync(WastageTypeDto model)
        {
            Validate(model);
            return await _dataAccess.SaveWastageTypeAsync(model);
        }

        public async Task<bool> UpdateWastageTypeAsync(WastageTypeDto model)
        {
            Validate(model);
            return await _dataAccess.UpdateWastageTypeAsync(model);
        }

        public Task<(bool CanDelete, string Reason)> CanDeleteWastageTypeAsync(int entryId) => _dataAccess.CanDeleteWastageTypeAsync(entryId);

        public Task<bool> DeleteWastageTypeAsync(int entryId) => _dataAccess.DeleteWastageTypeAsync(entryId);

        private static void Validate(WastageTypeDto model)
        {
            if (string.IsNullOrWhiteSpace(model.WastageName))
                throw new InvalidOperationException("Please enter a Wastage Type name / description.");

            if (model.WastageType < 0 || model.WastageType > 3)
                throw new InvalidOperationException("Please select a valid Cost Allocation Strategy (Store, Vendor, Originator, or Ask).");
        }
    }
}
