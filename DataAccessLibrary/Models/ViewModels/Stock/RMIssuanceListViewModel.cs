using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMIssuanceListRowModel
    {
        public string IssNo { get; set; }
        public DateTime? DT { get; set; }
        public string VenderName { get; set; }
        public string IssdToEmpID { get; set; }
        public string RM { get; set; }
        public double? Total_Issued_Qty { get; set; }
    }

    public class RMGroupLookupModel
    {
        public int ID { get; set; }
        public string Description { get; set; }
        public string Display => Description;
    }

    public class DepartmentLookupModel
    {
        public string DeptID { get; set; }
        public string Name { get; set; }
        public string Display => Name;
    }
}
