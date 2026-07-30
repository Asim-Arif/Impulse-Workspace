using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IClearShortTermDataAccess
    {
        /// <summary>
        /// Returns the employee's open short-term advances
        /// (Advances WHERE EmpID=@id AND Type=0 AND Cleared=0 ORDER BY EntryID).
        /// </summary>
        Task<List<ShortTermAdvanceItem>> GetOpenAdvancesAsync(string empId);

        /// <summary>
        /// Records the clearance in ONE SQL transaction:
        /// inserts the ShortTermAdvanceClearance header (Type=0), then per
        /// selected advance sets Advances.Cleared=2 and inserts a
        /// ShortTermAdvanceClearanceDetail row (with the advance's Amount).
        /// The voucher (if any) has already been posted by the service layer
        /// and its VchrNo is passed in here (empty when no voucher was posted).
        /// </summary>
        Task SaveClearanceAsync(ClearShortTermDto dto, string vchrNo, string userName, string machineName);
    }
}
