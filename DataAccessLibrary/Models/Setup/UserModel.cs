using System;

namespace DataAccessLibrary.Models.Setup
{
    public class UserModel
    {
        public int UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string? Password { get; set; }
        public string? FullUserName { get; set; }
        public string? EmpID { get; set; }
        public bool InActive { get; set; }
        public System.Collections.Generic.List<string> AssignedRoles { get; set; } = new();

        // Joined Employee Details
        public string? EmployeeName { get; set; }
        public string? Designation { get; set; }
        public string? DepartmentName { get; set; }
        public string? CellNo { get; set; }

        // Core Management Permissions
        public bool UserManagement { get; set; }
        public bool ChangePassword { get; set; }

        // Module Access Flags
        public bool CompanyMainLink { get; set; }
        public bool FinancialMainLink { get; set; }
        public bool PayrollMainLink { get; set; }
        public bool ExportMainLink { get; set; }
        public bool StockMainLink { get; set; }
        public bool ProductionMainLink { get; set; }
        public bool DashBoardMainLink { get; set; }
        public bool QMSMainLink { get; set; }
        public bool FixedAssetsMainLink { get; set; }
        public bool SamplingMainLink { get; set; }
        public bool HelpMainLink { get; set; }
        public bool IntraOfficeMainLink { get; set; }
        public bool OpenCommandCenter { get; set; }

        // Item Profile Access Controls
        public bool RestrictedItemProfile { get; set; }
        public bool GeneralInfoItemProfile { get; set; }
        public bool ProcessesItemProfile { get; set; }
        public bool WeightItemProfile { get; set; }
        public bool PriceItemProfile { get; set; }
        public bool ReferencesItemProfile { get; set; }
        public bool PictureItemProfile { get; set; }
        public bool RMItemProfile { get; set; }
        public bool ShipInfoItemProfile { get; set; }
        public bool RestrictCompanyCatalogEditing { get; set; }

        // Operational & Approval Authorizations
        public bool AuthorizeVouchers { get; set; }
        public bool AuthorizeIssuance { get; set; }
        public bool PostMakerBill { get; set; }
        public bool ChangeRateonIssuance { get; set; }
        public bool ChangeRateMakerAssign { get; set; }
        public bool MakerBill_EditRate { get; set; }
        public bool HideRateMakerAssign { get; set; }
        public bool AddEditCustomerComplaint { get; set; }
        public bool EditFollowUp { get; set; }
        public bool CloseCAPA { get; set; }

        // System & Maintenance
        public bool BackupData { get; set; }
        public bool RestoreData { get; set; }
        public bool HicoVisible { get; set; }
        public bool AddProdPlan { get; set; }
        public bool ShowTips { get; set; }
    }
}
