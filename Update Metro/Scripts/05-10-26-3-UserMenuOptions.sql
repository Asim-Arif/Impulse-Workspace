
-- ==============================================================================
-- Deduplicate and Harmonize MenuOptions with UserMenuOptions
-- Realigns Blazor OptionIDs with canonical legacy OptionIDs where user rights reside
-- ==============================================================================

-- 1. Ensure any permissions granted under duplicate OptionIDs are preserved in canonical OptionIDs
INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CompanyCatalog'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'CmpCompanyCatalog'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CompanyCatalog');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CustomerCatalog'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'CmpCustomerCatalog'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CustomerCatalog');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'OrderItemList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpOrderItemList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'OrderItemList');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CustomInvoice'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpCustomInvoice'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CustomInvoice');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'NewCustomInvoice'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpNewCustomInvoice'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'NewCustomInvoice');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrintValuationForm'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpValuationForm'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrintValuationForm');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkMaterialGroup'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkRMGroups'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkMaterialGroup');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkVenderBilling'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkVendorBilling'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkVenderBilling');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkVenderBillingList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkVendorBillingList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkVenderBillingList');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlTakeAttendanceEx'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayAttendanceManual'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlTakeAttendanceEx');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlAbsentSheet'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayAbsentSheet'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlAbsentSheet');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlSocialSecurity'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PaySocialSecurity'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlSocialSecurity');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlEOBI'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayEOBI'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlEOBI');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlPayrollPolicies'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayPolicies'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlPayrollPolicies');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrdReceivingAgainstPO'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PrdReceivePO'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrdReceivingAgainstPO');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'AccMakerList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PrdMakerList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'AccMakerList');
GO

-- 2. Remove orphaned duplicate assignments from UserMenuOptions
DELETE FROM UserMenuOptions 
WHERE OptionID IN (
    'CmpCompanyCatalog', 'CmpCustomerCatalog', 'ExpOrderItemList', 'ExpCustomInvoice',
    'ExpNewCustomInvoice', 'ExpValuationForm', 'StkRMGroups', 'StkVendorBilling',
    'StkVendorBillingList', 'PayAttendanceManual', 'PayAbsentSheet', 'PaySocialSecurity',
    'PayEOBI', 'PayPolicies', 'PrdReceivePO', 'PrdMakerList'
);
GO

-- 3. Remove orphaned duplicate rows from MenuOptions
DELETE FROM MenuOptions 
WHERE OptionID IN (
    'CmpCompanyCatalog', 'CmpCustomerCatalog', 'ExpOrderItemList', 'ExpCustomInvoice',
    'ExpNewCustomInvoice', 'ExpValuationForm', 'StkRMGroups', 'StkVendorBilling',
    'StkVendorBillingList', 'PayAttendanceManual', 'PayAbsentSheet', 'PaySocialSecurity',
    'PayEOBI', 'PayPolicies', 'PrdReceivePO', 'PrdMakerList'
);
GO

-- 4. Disambiguate identical display names that represent different screens
UPDATE MenuOptions SET OptionName = 'Commercial Packing List' WHERE OptionID = 'ComPackingList' AND OptionName = 'Packing List';
UPDATE MenuOptions SET OptionName = 'Custom Packing List' WHERE OptionID = 'CustomPackingList' AND OptionName = 'Packing List';
UPDATE MenuOptions SET OptionName = 'Change Locations (Raw Material)' WHERE OptionID = 'StkChangeLocations' AND OptionName = 'Change Locations';
UPDATE MenuOptions SET OptionName = 'Change Locations (Semi-Finished)' WHERE OptionID = 'StkChangeLocationsSF' AND OptionName = 'Change Locations';
GO
