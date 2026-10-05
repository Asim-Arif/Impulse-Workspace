using System;

namespace DataAccessLibrary.Models.Setup
{
    public class UserCustomerPermissionModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public bool IsAssigned { get; set; }
    }

    public class UserStorePermissionModel
    {
        public int StoreId { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public bool IsAssigned { get; set; }
    }
}
