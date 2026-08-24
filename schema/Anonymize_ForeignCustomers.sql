-- ==============================================================================
-- Script: Anonymize_ForeignCustomers.sql
-- Purpose: Anonymize 'Name' and 'Cont1Name' in ForeignCustomers table while
--          preserving all 'CustCode', 'Country', and other foreign key links.
-- Execution: Run in SQL Server Management Studio (SSMS) against your database.
-- ==============================================================================

USE [SMBI_AWM]; -- Replace with your actual database name if different
GO

BEGIN TRANSACTION;

-- Preview existing data before anonymization
PRINT '=== PRE-UPDATE BACKUP SNAPSHOT ===';
SELECT CustCode, Country, Name AS [OldName], Cont1Name AS [OldContact] 
FROM ForeignCustomers
ORDER BY CustCode;

-- Temporary mapping table for realistic anonymized international names
DECLARE @CustomerMap TABLE (
    CustCode VARCHAR(50) PRIMARY KEY,
    AnonymizedName VARCHAR(255),
    AnonymizedContact VARCHAR(100)
);

INSERT INTO @CustomerMap (CustCode, AnonymizedName, AnonymizedContact)
VALUES
    ('A-700', 'Apex Dental & Surgical Instruments', 'Alexander Vance'),
    ('A-701', 'MedTech International Italia S.R.L.', 'Gianluca Rossi'),
    ('A-704', 'Central America Dental Supplies S.A.', 'Alberto Morales'),
    ('A-707', 'Vanguard Medical Innovations LLC', 'Rebecca Jenkins'),
    ('A-711', 'Lombardia Precision Tools S.p.A.', 'Barbara Moretti'),
    ('A-713', 'Elysium Medical Group Italia', 'Claudia Mancini'),
    ('A-714', 'Equine Care & Dental Equipment S.A.S.', 'Henri Dubois'),
    ('A-715', 'Pacific Health Products Inc.', 'Becky Thompson'),
    ('A-717', 'Sul Dental Equipamentos Ltda.', 'Carlos Silva'),
    ('A-719', 'NovaCare Surgical Corp.', 'Edgar Ramirez'),
    ('A-721', 'Nordic SurgiLine Sp. z o.o.', 'Monika Wisniewska'),
    ('A-725', 'Aeterna Dental Solutions S.r.l.', 'Antonio Ferrara'),
    ('A-726', 'Bavaria Dental Instrumente GmbH', 'Anna-Maria Schmidt'),
    ('A-727', 'Lusitania Medicos & Saude Lda.', 'Serafim Ferreira'),
    ('A-729', 'Alpha Diamond Instruments S.R.L.S.', 'Matteo Bianchi'),
    ('A-730', 'Horizon Medical Technologies LLC', 'Frank Miller'),
    ('A-732', 'Medica Dynamic Care S.r.l.', 'Marco De Luca'),
    ('A-734', 'Riviera Medical Supply', 'Alberto Conti'),
    ('A-735', 'Helvetia Dental Specialties AG', 'Lucas Weber'),
    ('A-736', 'Atelier Clinica Dentaria Lda.', 'Joao Santos'),
    ('A-737', 'Medici & Chirurgia Strumenti S.r.l.', 'Massimo Ricci'),
    ('A-738', 'Anatolia Medikal Cihazlar Ltd.', 'Sureyya Yilmaz'),
    ('A-740', 'Venezia Dental Care', 'Diego Colombo'),
    ('A-741', 'Noordzee Tandartsen B.V.', 'Pieter van Dijk'),
    ('A-742', 'Rheinland Dentaltechnik GmbH', 'Klaus Becker'),
    ('A-743', 'Atlantic Surgical Instruments Corp.', 'Ramesh Patel'),
    ('D-003', 'Victoria Health & Dental Care Ltd.', 'Brian Wong'),
    ('PARTS', 'PARTS', ''),
    ('SAMPLES', 'SAMPLES', ''),
    ('STOCK', 'STOCK', '');

-- Update ForeignCustomers using the mapped anonymized values
UPDATE fc
SET 
    fc.Name = m.AnonymizedName,
    fc.Cont1Name = m.AnonymizedContact
FROM ForeignCustomers fc
INNER JOIN @CustomerMap m ON fc.CustCode = m.CustCode;

-- Fallback for any customer code not explicitly in the static map (dynamic anonymizer)
UPDATE ForeignCustomers
SET 
    Name = 'International Client ' + CustCode,
    Cont1Name = 'Contact ' + CustCode
WHERE CustCode NOT IN (SELECT CustCode FROM @CustomerMap)
  AND CustCode NOT IN ('PARTS', 'SAMPLES', 'STOCK');

-- Verify the updated values
PRINT '=== POST-UPDATE VERIFIED RECORDS ===';
SELECT CustCode, Country, Name AS [AnonymizedName], Cont1Name AS [AnonymizedContact] 
FROM ForeignCustomers
ORDER BY CustCode;

-- Commit transaction
COMMIT TRANSACTION;
PRINT 'ForeignCustomers anonymization completed successfully!';
GO
