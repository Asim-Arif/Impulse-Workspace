using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface ICustomerComplaintDataAccess
    {
        Task<List<CustomerComplaintListItemModel>> GetComplaintsAsync(CustomerComplaintFilterDto filter);
        Task<List<string>> GetCustomerCodesAsync();
        Task<List<string>> GetCountriesAsync(string? custCode = null);
        Task<string?> GetNCNoAsync(int entryId);
        Task<string?> GetCAPANoAsync(int entryId);
        Task<int> GetNCEntryIdAsync(string ncNo);
        Task<int> GetCAPAEntryIdAsync(string capaNo);

        // Add / Edit & Detail Methods
        Task<CustomerComplaintModel?> GetComplaintByIdAsync(int entryId);
        Task<List<CustomerComplaintDetailModel>> GetComplaintDetailsAsync(int refId);
        Task<int> SaveComplaintAsync(CustomerComplaintModel model, List<CustomerComplaintDetailModel> details, string currentUserName, string? machineName);
        Task<string> GetNextComplaintNoAsync(DateTime date);
        Task<string> GetNextFeedbackNoAsync(DateTime date);
        Task<string> GetNextAuditNoAsync(DateTime date);

        // Non-Compliance (NC) Methods
        Task<NonComplianceModel?> GetNCByIdAsync(int entryId);
        Task<NonComplianceModel?> GetNCByComplaintIdAsync(int complaintId);
        Task<string> GetNextNCNoAsync(DateTime date);
        Task<int> SaveNCAsync(NonComplianceModel model, int? complaintId, string currentUserName);

        // CAPA Methods
        Task<CAPAModel?> GetCAPAByIdAsync(int entryId);
        Task<CAPAModel?> GetCAPAByComplaintIdAsync(int complaintId);
        Task<string> GetNextCAPANoAsync(DateTime date);
        Task<int> SaveCAPAAsync(CAPAModel model, int? complaintId, string currentUserName);

        // Credit Note Methods
        Task<string> GetNextCreditNoteNoAsync(DateTime date);
        Task<bool> SaveCreditNoteAsync(int complaintId, string creditNoteNo, List<CustomerComplaintDetailModel> details, string currentUserName);

        // Lookups
        Task<List<CompanyLookupDto>> GetCompaniesAsync();
        Task<List<DepartmentLookupDto>> GetDepartmentsAsync();
        Task<List<QMSReasonLookupDto>> GetQMSReasonsAsync();
        Task<List<EmployeeLookupDto>> GetEmployeesByDeptAsync(string deptId);
        Task<List<string>> GetInvoicesByCustomerAsync(string custCode, string? country);
        Task<List<InvoiceItemLookupDto>> GetInvoiceItemsAsync(string custCode, string? country, string? invoiceNo);
        Task<List<string>> GetLotsAsync(string invoiceNo, string itemId);
    }
}
