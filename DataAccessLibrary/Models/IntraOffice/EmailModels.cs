using System;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class EmailConfiguration
    {
        public int Id { get; set; }
        public string SmtpServer { get; set; } = "smtp.mail.yahoo.com";
        public string SmtpHost { get => SmtpServer; set => SmtpServer = value; }
        public int SmtpPort { get; set; } = 587;
        public string SenderEmail { get; set; } = "";
        public string SenderName { get; set; } = "IntraCom-CRM";
        public string SenderDisplayName { get => SenderName; set => SenderName = value; }
        public string Username { get; set; } = "";
        public string EncryptedPassword { get; set; } = "";
        public bool EnableSsl { get; set; } = true;
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get => CreatedAt; set => CreatedAt = value; }
        public string? UpdatedBy { get; set; }
    }

    public class EmailTemplate
    {
        public int Id { get; set; }
        public string TemplateCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Category { get; set; } = "QuotationFollowUp";
        public string SubjectTemplate { get; set; } = string.Empty;
        public string BodyTemplate { get; set; } = string.Empty;
        public string? AvailablePlaceholders { get; set; } = "{CompanyName}, {ContactName}, {QuoteNumber}, {Amount}, {Currency}, {Incoterm}, {SalespersonName}";
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
