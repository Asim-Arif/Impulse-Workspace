-- ============================================================================
-- Gate Management: Screen & Menu Options Synchronization
-- ============================================================================
DELETE FROM dbo.UserMenuOptions 
WHERE OptionID IN (
    'GateMaterialInward', 'GateMaterialInwardEntry', 
    'GateMaterialOutward', 'GateMaterialOutwardEntry',
    'GateInwardList', 'GateInwardEntry', 'GateInwardReceive'
);

DELETE FROM dbo.MenuOptions 
WHERE OptionID IN (
    'GateMaterialInward', 'GateMaterialInwardEntry', 
    'GateMaterialOutward', 'GateMaterialOutwardEntry',
    'GateInwardList', 'GateInwardEntry', 'GateInwardReceive'
);

MERGE dbo.MenuOptions AS target
USING (VALUES
    ('GateHub', 'Gate Management Hub', 'GateManagement', 'GateManagementHub'),
    ('GateDashboard', 'Gate Operations Dashboard', 'GateManagement', 'GateDashboard'),
    ('GateIn', 'Gate In (Inward Register)', 'GateManagement', 'GateInList'),
    ('GateInEntry', 'Gate In Pass Entry', 'GateManagement', 'GateInEntry'),
    ('GatePassPermanent', 'Gate Pass Permanent (Outward)', 'GateManagement', 'GatePassPermanentList'),
    ('GatePassPermanentEntry', 'Permanent Gate Pass Entry', 'GateManagement', 'GatePassPermanentEntry'),
    ('GatePassReturnable', 'Gate pass Returnable', 'GateManagement', 'GatePassReturnableList'),
    ('GatePassReturnableEntry', 'Returnable Gate Pass Entry', 'GateManagement', 'GatePassReturnableEntry'),
    ('CourierIssuance', 'Courier Issuance', 'GateManagement', 'CourierIssuanceList'),
    ('CourierIssuanceEntry', 'New Courier Pass', 'GateManagement', 'CourierIssuanceEntry'),
    ('GateVisitors', 'Visitor Management Register', 'GateManagement', 'VisitorList'),
    ('GateVisitorEntry', 'Visitor Pass Entry', 'GateManagement', 'VisitorEntry'),
    ('GateVisitorCheckout', 'Visitor Gate Out Checkout', 'GateManagement', 'VisitorGateOut'),
    ('GateVehicles', 'Vehicle Movement Log', 'GateManagement', 'VehicleList'),
    ('GateVehicleEntry', 'Vehicle Entry / Exit', 'GateManagement', 'VehicleEntry'),
    ('GatePassTypes', 'Gate Pass Types Setup', 'GateManagement', 'GatePassTypes'),
    ('GateReports', 'Gate Security Reports', 'GateManagement', 'GateReports')
) AS source (OptionID, OptionName, ModuleName, FormName)
ON target.OptionID = source.OptionID
WHEN MATCHED THEN
    UPDATE SET 
        target.OptionName = source.OptionName,
        target.ModuleName = source.ModuleName,
        target.FormName = source.FormName
WHEN NOT MATCHED THEN
    INSERT (OptionID, OptionName, ModuleName, FormName)
    VALUES (source.OptionID, source.OptionName, source.ModuleName, source.FormName);

INSERT INTO dbo.UserMenuOptions (UserID, OptionID)
SELECT DISTINCT u.UserID, mo.OptionID
FROM dbo.Users u
CROSS JOIN dbo.MenuOptions mo
WHERE mo.ModuleName = 'GateManagement'
  AND (
      u.UserName = 'Administrator'
      OR ISNULL(u.GateMainLink, 0) = 1
      OR ISNULL(u.UserManagement, 0) = 1
      OR EXISTS (
          SELECT 1 FROM dbo.UserMenuOptions umo2 
          INNER JOIN dbo.MenuOptions mo2 ON umo2.OptionID = mo2.OptionID
          WHERE umo2.UserID = u.UserID AND mo2.ModuleName = 'GateManagement'
      )
  )
  AND NOT EXISTS (
      SELECT 1 FROM dbo.UserMenuOptions existing 
      WHERE existing.UserID = u.UserID AND existing.OptionID = mo.OptionID
  );
GO