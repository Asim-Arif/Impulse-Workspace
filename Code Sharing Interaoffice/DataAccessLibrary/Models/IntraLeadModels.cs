using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class LeadModel
    {
        public int Id { get; set; }
        public string LeadNumber { get; set; } = string.Empty;
        public string CompanyName { get; set; } = string.Empty;
        public string ContactPerson { get; set; } = string.Empty;
        public string? Designation { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public string? WhatsApp { get; set; }
        public string Country { get; set; } = string.Empty;
        public string? City { get; set; }
        public string? Address { get; set; }
        public string? Website { get; set; }
        public string Source { get; set; } = "Exhibition";
        public string? Industry { get; set; }
        public string? ProductInterest { get; set; }
        public int EstimatedQuantity { get; set; }
        public decimal EstimatedValue { get; set; }
        public string Currency { get; set; } = "USD";
        public DateTime? ExpectedOrderDate { get; set; }
        public string Status { get; set; } = "New"; // New, Contacted, Qualified, Proposal, Won, Lost, Converted
        public string Priority { get; set; } = "Medium"; // Low, Medium, High, Urgent
        public string? Notes { get; set; }
        public DateTime? NextFollowUpDate { get; set; }
        public string? ConvertedCustCode { get; set; }
        public string? AssignedTo { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

        public List<LeadActivityModel> Activities { get; set; } = new();
    }

    public class LeadActivityModel
    {
        public int Id { get; set; }
        public int LeadId { get; set; }
        public string ActivityType { get; set; } = "Call"; // Call, Email, Meeting, WhatsApp, Note, Site Visit
        public string Description { get; set; } = string.Empty;
        public string PerformedBy { get; set; } = string.Empty;
        public DateTime ActivityDate { get; set; } = DateTime.UtcNow;
        public DateTime? NextFollowUpDate { get; set; }
    }

    public class ConvertLeadToCustomerModel
    {
        public int LeadId { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Address { get; set; }
        public string? City { get; set; }
        public string? Phone1 { get; set; }
        public string? Email1 { get; set; }
        public string? Cont1name { get; set; }
        public string? Cont1Mobile { get; set; }
        public string? CustomerSource { get; set; }

        public string CustomerName { get => Name; set => Name = value; }
        public string? ContactPerson { get => Cont1name; set => Cont1name = value; }
        public string? Phone { get => Phone1; set => Phone1 = value; }
        public string? Email { get => Email1; set => Email1 = value; }
        public string ConvertedBy { get; set; } = string.Empty;
    }
}
