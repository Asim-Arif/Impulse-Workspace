-- ==============================================================================================
-- IntraOffice Missing Tables & Data Migration Script
-- Target: Live & Development SQL Server Databases
-- Created: 2026-09-23
-- Features: Email Templates & SMTP Configuration
-- ==============================================================================================

USE [SMBI_AWM];
GO

-- 1. Ensure EmailConfigurations Table Exists
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'EmailConfigurations')
BEGIN
    CREATE TABLE [dbo].[EmailConfigurations] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [SmtpServer] NVARCHAR(200) NOT NULL,
        [SmtpPort] INT NOT NULL DEFAULT 587,
        [SenderEmail] NVARCHAR(200) NOT NULL,
        [SenderName] NVARCHAR(200) NOT NULL,
        [Username] NVARCHAR(200) NOT NULL,
        [EncryptedPassword] NVARCHAR(500) NOT NULL,
        [EnableSsl] BIT NOT NULL DEFAULT 1,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: EmailConfigurations';
END
ELSE
BEGIN
    PRINT 'Table EmailConfigurations already exists.';
END
GO

-- 2. Ensure EmailTemplates Table Exists
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'EmailTemplates')
BEGIN
    CREATE TABLE [dbo].[EmailTemplates] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TemplateCode] NVARCHAR(50) NOT NULL,
        [Name] NVARCHAR(150) NOT NULL,
        [Category] NVARCHAR(50) NOT NULL DEFAULT 'QuotationFollowUp',
        [SubjectTemplate] NVARCHAR(255) NOT NULL,
        [BodyTemplate] NVARCHAR(MAX) NOT NULL,
        [AvailablePlaceholders] NVARCHAR(500) NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: EmailTemplates';
END
ELSE
BEGIN
    PRINT 'Table EmailTemplates already exists.';
END
GO

-- 3. Seed Default Communication Templates
IF NOT EXISTS (SELECT 1 FROM [dbo].[EmailTemplates] WHERE [TemplateCode] = 'QT-FOLLOWUP')
BEGIN
    INSERT INTO [dbo].[EmailTemplates] (
        [TemplateCode], [Name], [Category], [SubjectTemplate], [BodyTemplate], [AvailablePlaceholders], [IsActive], [CreatedAt]
    ) VALUES (
        'QT-FOLLOWUP',
        'Official Quotation Follow-up & Commercial Review',
        'QuotationFollowUp',
        'Follow-up: Commercial Proposal for {CompanyName} - Ref #{QuoteNumber}',
        N'Dear {ContactName},

I hope this email finds you well.

Following our recent discussions, I am following up on our official quotation #{QuoteNumber} amounting to {Currency} {Amount} ({Incoterm} terms) submitted to {CompanyName}.

Our production line is currently scheduling manufacturing slots for the upcoming quarter. We would appreciate learning if you have any questions regarding the technical specifications, certifications, or delivery schedules.

Please let us know if you require any revisions or wish to proceed with issuing the purchase order.

Best regards,
{SalespersonName}
Export Sales Department
AWM Systems',
        '{CompanyName}, {ContactName}, {QuoteNumber}, {Amount}, {Currency}, {Incoterm}, {SalespersonName}',
        1,
        GETUTCDATE()
    );
    PRINT 'Seeded template: QT-FOLLOWUP';
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[EmailTemplates] WHERE [TemplateCode] = 'SAMPLE-DISPATCH')
BEGIN
    INSERT INTO [dbo].[EmailTemplates] (
        [TemplateCode], [Name], [Category], [SubjectTemplate], [BodyTemplate], [AvailablePlaceholders], [IsActive], [CreatedAt]
    ) VALUES (
        'SAMPLE-DISPATCH',
        'Export Product Sample Dispatch & Tracking Pre-Alert',
        'SampleNotice',
        'Sample Shipment Dispatched: {CompanyName} - AWB #{TrackingNumber}',
        N'Dear {ContactName},

We are pleased to inform you that your requested product samples have been carefully prepared, quality inspected, and dispatched via courier.

Shipment Details:
- Courier Service: {CourierName}
- Air Waybill / Tracking No: {TrackingNumber}
- Dispatch Date: {Date}
- Estimated Delivery: {DeliveryDate}
- Destination: {DestinationCountry}

You can track the live movement of the parcel on the courier website using the tracking number above. We look forward to receiving your feedback upon testing and evaluation.

Warm regards,
{SalespersonName}
Technical & Customer Relations
AWM Systems',
        '{CompanyName}, {ContactName}, {CourierName}, {TrackingNumber}, {Date}, {DeliveryDate}, {DestinationCountry}, {SalespersonName}',
        1,
        GETUTCDATE()
    );
    PRINT 'Seeded template: SAMPLE-DISPATCH';
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[EmailTemplates] WHERE [TemplateCode] = 'PAYMENT-REMINDER')
BEGIN
    INSERT INTO [dbo].[EmailTemplates] (
        [TemplateCode], [Name], [Category], [SubjectTemplate], [BodyTemplate], [AvailablePlaceholders], [IsActive], [CreatedAt]
    ) VALUES (
        'PAYMENT-REMINDER',
        'Commercial Invoice & Payment Due Notice',
        'PaymentReminder',
        'Statement Notice: Payment Due for {CompanyName} - Invoice Total {Currency} {Amount}',
        N'Dear {ContactName},

This is a courtesy reminder regarding the outstanding balance for {CompanyName}.

Payment Summary:
- Agreed Payment Terms: {PaymentTerms}
- Outstanding Amount: {Currency} {Amount}
- Due Date: {DueDate}

Please arrange the wire transfer to our designated beneficiary bank account and forward the swift transaction confirmation copy once processed.

If payment has already been remitted in the last 24-48 hours, please disregard this notice with our thanks.

Sincerely,
{SalespersonName}
Finance & Accounts Department
AWM Systems',
        '{CompanyName}, {ContactName}, {PaymentTerms}, {Currency}, {Amount}, {DueDate}, {SalespersonName}',
        1,
        GETUTCDATE()
    );
    PRINT 'Seeded template: PAYMENT-REMINDER';
END
GO

IF NOT EXISTS (SELECT 1 FROM [dbo].[EmailTemplates] WHERE [TemplateCode] = 'SHIPMENT-ADVICE')
BEGIN
    INSERT INTO [dbo].[EmailTemplates] (
        [TemplateCode], [Name], [Category], [SubjectTemplate], [BodyTemplate], [AvailablePlaceholders], [IsActive], [CreatedAt]
    ) VALUES (
        'SHIPMENT-ADVICE',
        'Export Container Shipping Pre-Alert & Logistics Advice',
        'ShipmentAdvice',
        'Shipping Advice Pre-Alert: Container Onboard for {CompanyName} - Vessel {VesselName}',
        N'Dear {ContactName},

We are pleased to advise that your export cargo has been successfully loaded on board the vessel and is en route to its destination port.

Logistics Overview:
- Ocean Vessel: {VesselName} (Voyage: {VoyageNumber})
- Port of Loading: {PortOfLoading}
- Port of Discharge: {PortOfDischarge}
- Estimated Time of Departure (ETD): {ETD}
- Estimated Time of Arrival (ETA): {ETA}

Attached please find non-negotiable copies of the Bill of Lading, Packing List, Certificate of Origin, and Commercial Invoice for import pre-clearance procedures.

Best regards,
{SalespersonName}
Logistics & Shipping Operations
AWM Systems',
        '{CompanyName}, {ContactName}, {VesselName}, {VoyageNumber}, {PortOfLoading}, {PortOfDischarge}, {ETD}, {ETA}, {SalespersonName}',
        1,
        GETUTCDATE()
    );
    PRINT 'Seeded template: SHIPMENT-ADVICE';
END
GO

PRINT '==============================================================================================';
PRINT 'IntraOffice Email Templates & SMTP Configuration script executed successfully!';
PRINT '==============================================================================================';
GO
